using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using OMNIX.Core.AiGateway;
using OMNIX.Core.Context;
using OMNIX.Core.Errors;
using OMNIX.Core.Logging;
using OMNIX.Core.Settings;
using OMNIX.Core.Storage;
using OMNIX.Core.Tools;
using OMNIX.Core.Ui.Dialogs;
using OMNIX.Core.Util;

namespace OMNIX.Core.Ui
{
    /// <summary>
    /// Per-window controller: every Office document window owns its own WorkspaceView, AI Gateway,
    /// provider adapters, privacy-session approval, cancellation token and conversation state.
    ///
    /// The per-window gateway is intentional: provider adapters keep request configuration in
    /// memory. Sharing one mutable adapter/gateway across two Office windows can race credentials,
    /// model selection and AskBeforeSending callbacks. Isolation prevents one document window from
    /// borrowing another window's provider state or cloud-consent session.
    /// </summary>
    public sealed class WorkspaceController : IDisposable
    {
        private readonly IHostAdapter _adapter;
        private readonly AiGateway.AiGateway _gateway;
        private readonly ChatHistoryStore _historyStore;
        private readonly ToolExecutor _toolExecutor;

        private CancellationTokenSource _cts;
        private List<ChatTurn> _turns = new List<ChatTurn>();
        private string _docKey = "unnamed";
        private bool _busy;
        private bool _disposed;
        private bool _pendingContextRefresh;
        private int _documentScopeVersion;

        public WorkspaceView View { get; private set; }
        public AiGateway.AiGateway Gateway { get { return _gateway; } }
        public ProviderRegistry GatewayRegistry { get { return _gateway.Registry; } }

        public WorkspaceController(IHostAdapter adapter, ChatHistoryStore historyStore)
        {
            if (adapter == null) throw new ArgumentNullException("adapter");
            if (historyStore == null) throw new ArgumentNullException("historyStore");

            _adapter = adapter;
            _historyStore = historyStore;

            // Per-workspace registry/adapters/gateway: no mutable provider state is shared between
            // two open documents. Local runtime discovery is asynchronous and never blocks pane UI.
            _gateway = new AiGateway.AiGateway(new ProviderRegistry());

            _toolExecutor = new ToolExecutor();
            _toolExecutor.WriteConfirmation = preview =>
                Application.Current != null
                    ? RunOnUiThread(() => OmnixDialogs.ConfirmWritePreview(preview))
                    : Task.FromResult(false);

            // This callback now belongs only to THIS workspace's PrivacyGate, so "remember for this
            // session" cannot silently approve a different document window.
            _gateway.Privacy.CloudConfirmationCallback = providerName =>
                RunOnUiThread(() =>
                {
                    var ctx = _adapter.ReadContext();
                    return OmnixDialogs.ConfirmCloudSend(providerName, ctx.ContextBarText);
                });

            View = new WorkspaceView(this);
            Theming.ThemeManager.Instance.ApplyTo(View);
            View.Resources.MergedDictionaries.Add(Localization.Strings.Dictionary);
            Theming.ThemeManager.Instance.ThemeChanged += OnThemeChanged;
            RefreshContextBar(initial: true);
            ObserveBackground(Task.Run(() => _gateway.ProbeLocalAsync()), "initial local provider probe");
        }

        // ------------------------------------------------------------------ lifecycle

        public void RefreshContextBar(bool initial = false)
        {
            if (_disposed) return;
            try
            {
                var ctx = _adapter.ReadContext();
                string newKey = StableDocumentKey(ctx);
                bool docChanged = !string.Equals(newKey, _docKey, StringComparison.OrdinalIgnoreCase);

                // Any observed document identity change invalidates the previous request scope
                // immediately. Streaming callbacks can check this integer without touching COM.
                if (docChanged)
                    Interlocked.Increment(ref _documentScopeVersion);

                // Never retarget an in-flight request/history to a different Office document.
                // Cancel the old request and defer the history switch until its async continuation
                // has unwound. This also causes the ToolExecutor request-scope guard to fail closed.
                if (_busy && docChanged)
                {
                    _pendingContextRefresh = true;
                    CancelActiveRequest();
                    View.Chat.SetContextText(ctx.ContextBarText);
                    Logger.Ui("Active Office document changed during AI request; request cancelled and context switch deferred.");
                    return;
                }

                _docKey = newKey;

                if (docChanged || initial)
                {
                    _turns = _historyStore.Load(_docKey);
                    View.Chat.ReloadMessages(_turns);
                    if (docChanged) _gateway.Privacy.ResetSession();
                }

                View.Chat.SetContextText(ctx.ContextBarText);
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "RefreshContextBar failed", ex);
            }
        }

        public void OnPaneClosing()
        {
            CancelActiveRequest();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Interlocked.Increment(ref _documentScopeVersion);

            // Cancel but do not dispose an in-flight CTS here. SendMessage owns the CTS lifetime
            // and disposes it in finally after provider/tool continuations have unwound.
            CancelActiveRequest();
            _cts = null;

            _gateway.Privacy.ResetSession();
            _gateway.Privacy.CloudConfirmationCallback = null;
            Theming.ThemeManager.Instance.ThemeChanged -= OnThemeChanged;
        }

        private void OnThemeChanged()
        {
            if (_disposed) return;
            try { Theming.ThemeManager.Instance.ApplyTo(View); } catch { }
        }

        private static string StableDocumentKey(OfficeContext ctx)
        {
            if (ctx == null) return DocKeySanitizer.StableKey("office", null, "unnamed");
            return DocKeySanitizer.StableKey(
                ctx.Host.ToString(),
                ctx.DocumentPath,
                ctx.DocumentName);
        }

        /// <summary>
        /// Fast non-COM check used by high-frequency streaming callbacks. RefreshContextBar and
        /// Dispose increment the version as soon as a document/window scope change is observed.
        /// </summary>
        private bool IsRequestScopeVersionValid(string requestDocKey, int requestScopeVersion)
        {
            return !_disposed &&
                   Volatile.Read(ref _documentScopeVersion) == requestScopeVersion &&
                   string.Equals(_docKey, requestDocKey, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Deep check used at low-frequency security boundaries (tool execution and post-provider
        /// completion). It confirms both the in-memory scope generation and actual Office context.
        /// A COM/context failure is treated as scope loss.
        /// </summary>
        private bool ValidateCurrentOfficeDocumentScope(string requestDocKey, int requestScopeVersion)
        {
            if (!IsRequestScopeVersionValid(requestDocKey, requestScopeVersion)) return false;
            try
            {
                var ctx = _adapter.ReadContext();
                string currentKey = StableDocumentKey(ctx);
                return string.Equals(currentKey, requestDocKey, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        // ------------------------------------------------------------------ chat actions

        public async void SendMessage(string text, ImageAttachment image)
        {
            if (_disposed || _busy) return;
            text = (text ?? "").Trim();
            if (text.Length == 0 && image == null) return;

            // Re-read context immediately before binding the request. Office activation/selection
            // events are best-effort; Send must not rely on an earlier event having already fired.
            RefreshContextBar();
            if (_disposed || _busy) return;
            string requestDocKey = _docKey;
            int requestScopeVersion = Volatile.Read(ref _documentScopeVersion);

            _busy = true;
            View.Chat.SetBusy(true);

            var userTurn = new ChatTurn
            {
                Role = ChatRole.User,
                Text = text,
                TimestampUtc = DateTime.UtcNow
            };
            if (image != null && image.PngBytes != null && image.PngBytes.Length > 0)
                userTurn.Images = new List<ImageAttachment> { image };

            _turns.Add(userTurn);
            View.Chat.AppendTurn(userTurn);
            Persist(requestDocKey);

            var assistantTurn = new ChatTurn
            {
                Role = ChatRole.Assistant,
                Text = "",
                TimestampUtc = DateTime.UtcNow
            };
            var bubble = View.Chat.AppendTurn(assistantTurn);
            bubble.AppendText("…");

            if (_cts != null)
            {
                try { _cts.Cancel(); } catch { }
            }

            var requestCts = new CancellationTokenSource();
            _cts = requestCts;
            var ct = requestCts.Token;
            var sb = new System.Text.StringBuilder();

            // AiGateway owns the provider/tool loop, while this per-window executor owns the
            // Office boundary. The validator performs a deep Office identity check only when a
            // tool is actually about to cross into the host adapter, not for every streamed token.
            _toolExecutor.RequestCancellationTokenProvider = () => requestCts.Token;
            _toolExecutor.RequestScopeValidator = () =>
                ValidateCurrentOfficeDocumentScope(requestDocKey, requestScopeVersion);

            try
            {
                var request = new ChatRequest
                {
                    SystemPrompt = AiGateway.AiGateway.BuildSystemPrompt(_adapter, _adapter.ReadContext()),
                    History = _turns.Take(_turns.Count - 1).ToList(),
                    UserTurn = userTurn
                };

                var response = await _gateway.ChatAsync(
                    request,
                    _adapter,
                    delta =>
                    {
                        // Network adapters may emit deltas from a non-UI continuation. Never call
                        // Office COM here. The version check is in-memory and the UI update is
                        // dispatched only while the originating document scope remains valid.
                        if (!IsRequestScopeVersionValid(requestDocKey, requestScopeVersion)) return;
                        var app = Application.Current;
                        if (app == null) return;
                        app.Dispatcher.BeginInvoke(new Action(delegate
                        {
                            if (!IsRequestScopeVersionValid(requestDocKey, requestScopeVersion)) return;
                            sb.Append(delta);
                            bubble.ReplaceText(sb.ToString());
                        }));
                    },
                    _toolExecutor,
                    ct).ConfigureAwait(true);

                // A provider can race cancellation and return a completed response. Re-check both
                // token and actual Office document identity before touching UI/history after await.
                ct.ThrowIfCancellationRequested();
                if (!ValidateCurrentOfficeDocumentScope(requestDocKey, requestScopeVersion))
                    throw new OperationCanceledException("Office document changed during the AI request.", ct);
                if (_disposed) return;

                if (response != null && !string.IsNullOrEmpty(response.Text) && sb.Length == 0)
                    sb.Append(response.Text);

                assistantTurn.Text = sb.Length > 0 ? sb.ToString() : Localization.Strings.T("S.Chat.Cancelled");
                bubble.ReplaceText(assistantTurn.Text);
                _turns.Add(assistantTurn);
                Persist(requestDocKey);
            }
            catch (OperationCanceledException)
            {
                // Pane disposal or document switching invalidates this request completely. Do not
                // write a late cancellation bubble/history entry into a closed or different doc.
                if (_disposed || !ValidateCurrentOfficeDocumentScope(requestDocKey, requestScopeVersion)) return;

                // A user pressing Stop in the SAME document remains a normal visible cancellation.
                assistantTurn.Text = sb.ToString() + Environment.NewLine + Localization.Strings.T("S.Chat.Cancelled");
                bubble.ReplaceText(assistantTurn.Text);
                _turns.Add(assistantTurn);
                Persist(requestDocKey);
            }
            catch (OmnixException ex)
            {
                if (_disposed || !ValidateCurrentOfficeDocumentScope(requestDocKey, requestScopeVersion)) return;
                bubble.ReplaceText("");
                View.Chat.ShowError(ErrorPresenter.Format(ex));
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "SendMessage failed", ex);
                if (_disposed || !ValidateCurrentOfficeDocumentScope(requestDocKey, requestScopeVersion)) return;
                bubble.ReplaceText("");
                View.Chat.ShowError(ErrorPresenter.Format(ex));
            }
            finally
            {
                _busy = false;

                // Clear request bindings only if they still belong to this request. A controller
                // currently permits one request at a time, but reference checks keep this robust if
                // that UI policy changes later.
                if (ReferenceEquals(_cts, requestCts)) _cts = null;
                _toolExecutor.RequestCancellationTokenProvider = null;
                _toolExecutor.RequestScopeValidator = null;
                try { requestCts.Dispose(); } catch { }

                if (!_disposed)
                {
                    View.Chat.SetBusy(false);
                    if (_pendingContextRefresh)
                    {
                        _pendingContextRefresh = false;
                        RefreshContextBar();
                    }
                }
            }
        }

        public void StopStreaming()
        {
            CancelActiveRequest();
        }

        private void CancelActiveRequest()
        {
            try { if (_cts != null) _cts.Cancel(); } catch { }
        }

        public void NewChat()
        {
            if (_disposed || _busy) return;
            _turns = new List<ChatTurn>();
            _gateway.Privacy.ResetSession();
            View.Chat.ReloadMessages(_turns);
            View.Chat.SetStatus(Localization.Strings.T("S.Chat.NewSession"));
        }

        public void ClearChat()
        {
            if (_disposed || _busy) return;
            _turns = new List<ChatTurn>();
            _historyStore.Delete(_docKey);
            _gateway.Privacy.ResetSession();
            View.Chat.ReloadMessages(_turns);
            View.Chat.SetStatus(Localization.Strings.T("S.Chat.Cleared"));
        }

        public void CopyLastAnswer()
        {
            if (_disposed) return;
            var last = _turns.LastOrDefault(t => t.Role == ChatRole.Assistant && !string.IsNullOrEmpty(t.Text));
            if (last == null)
            {
                View.Chat.SetStatus(Localization.Strings.T("S.Chat.NothingToCopy"));
                return;
            }
            try
            {
                Clipboard.SetText(last.Text);
                View.Chat.SetStatus(Localization.Strings.T("S.Chat.Copied"));
            }
            catch { }
        }

        public void RetryLast()
        {
            if (_disposed || _busy) return;
            int idx = _turns.FindLastIndex(t => t.Role == ChatRole.User);
            if (idx < 0)
            {
                View.Chat.SetStatus(Localization.Strings.T("S.Chat.RetryEmpty"));
                return;
            }

            var resend = _turns[idx];
            _turns = _turns.Take(idx).ToList();
            View.Chat.ReloadMessages(_turns);
            ImageAttachment img = resend.HasImages ? resend.Images.FirstOrDefault(i => i != null && i.PngBytes != null && i.PngBytes.Length > 0) : null;
            SendMessage(resend.Text, img);
        }

        public void AttachImageFromDocument()
        {
            if (_disposed) return;
            try
            {
                byte[] png = _adapter.CaptureCurrentViewAsImage();
                if (png == null || png.Length == 0)
                {
                    View.Chat.SetStatus("No capturable Office view is available.");
                    return;
                }
                ImageNormalizer.ValidatePngBytes(png, _adapter.HostDisplayName + " capture");
                View.Chat.SetPendingImage(new ImageAttachment
                {
                    PngBytes = png,
                    FileName = "document-capture.png",
                    SourceLabel = _adapter.HostDisplayName
                });
            }
            catch (OmnixException ex)
            {
                View.Chat.ShowError(ErrorPresenter.Format(ex));
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "AttachImageFromDocument failed", ex);
                View.Chat.SetStatus("Capture failed");
            }
        }

        public void AttachImageFromDisk()
        {
            if (_disposed) return;
            try
            {
                var dlg = new System.Windows.Forms.OpenFileDialog
                {
                    Title = Localization.Strings.T("S.Chat.UploadImage"),
                    Filter = "Images (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp"
                };
                if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

                // All supported source formats are decoded and re-encoded as a bounded PNG.
                // Provider adapters can therefore truthfully send image/png instead of labeling
                // raw JPEG/BMP bytes as PNG. Invalid/oversized images fail before entering chat.
                byte[] png = ImageNormalizer.LoadFileAsPng(dlg.FileName);
                View.Chat.SetPendingImage(new ImageAttachment
                {
                    PngBytes = png,
                    FileName = System.IO.Path.GetFileNameWithoutExtension(dlg.FileName) + ".png",
                    SourceLabel = "file"
                });
            }
            catch (System.IO.InvalidDataException ex)
            {
                View.Chat.SetStatus(ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "AttachImageFromDisk failed", ex);
                View.Chat.SetStatus("Upload failed");
            }
        }

        public void SaveSettingsFromUi()
        {
            if (_disposed) return;
            SettingsManager.Instance.Save();
            _gateway.Privacy.ResetSession();
            View.Chat.SetStatus(Localization.Strings.T("S.Settings.Saved"));
        }

        private void Persist(string requestDocKey)
        {
            if (_disposed || string.IsNullOrWhiteSpace(requestDocKey)) return;
            _historyStore.Save(requestDocKey, _turns);
        }

        private static Task<T> RunOnUiThread<T>(Func<T> action)
        {
            var app = Application.Current;
            if (app == null) return Task.FromResult(default(T));
            return Task.FromResult(app.Dispatcher.Invoke(action));
        }

        private static async void ObserveBackground(Task task, string operation)
        {
            if (task == null) return;
            try { await task.ConfigureAwait(false); }
            catch (Exception ex) { Logger.Error("gateway", "Background " + operation + " failed", ex); }
        }
    }
}
