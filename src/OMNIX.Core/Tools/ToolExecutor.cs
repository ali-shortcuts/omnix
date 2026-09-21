using System;
using System.Threading;
using System.Threading.Tasks;
using OMNIX.Core.Context;
using OMNIX.Core.Errors;
using OMNIX.Core.Logging;
using OMNIX.Core.Security;
using OMNIX.Core.Util;

namespace OMNIX.Core.Tools
{
    /// <summary>
    /// Layer 7 — executes ONLY whitelisted tools. Read tools run directly; write tools first
    /// produce a preview, require explicit user confirmation, then apply through the host
    /// adapter so native Office undo (Ctrl+Z) keeps working.
    ///
    /// Tool execution is cancellation/scope-aware. This is a security boundary, not only a UX
    /// feature: when an Office document/window changes while an AI request is in flight, no late
    /// provider tool call may read from or mutate the newly-active document.
    /// </summary>
    public sealed class ToolExecutor
    {
        /// <summary>UI wires this: returns true when the user confirmed the change.</summary>
        public Func<string, string> ConversationSearch { get; set; }

        public Action<string> Progress { get; set; }

        public Func<WritePreview, Task<bool>> WriteConfirmation { get; set; }

        /// <summary>
        /// Per-workspace request token provider. WorkspaceController sets this only while one
        /// request owns the executor. Existing non-UI callers can leave it null.
        /// </summary>
        public Func<CancellationToken> RequestCancellationTokenProvider { get; set; }

        /// <summary>
        /// Fail-closed request scope check, normally bound to the document identity captured when
        /// Send was pressed. Returning false aborts the tool loop before Office COM is touched.
        /// </summary>
        public Func<bool> RequestScopeValidator { get; set; }

        /// <summary>
        /// Compatibility overload used by AiGateway and deterministic callers. Interactive
        /// WorkspaceController binds this executor to its current request token and document scope.
        /// </summary>
        public Task<ToolResult> ExecuteAsync(ToolCall call, IHostAdapter adapter)
        {
            CancellationToken ct = CancellationToken.None;
            var tokenProvider = RequestCancellationTokenProvider;
            if (tokenProvider != null)
            {
                try { ct = tokenProvider(); }
                catch { throw new OperationCanceledException("OMNIX request scope is no longer available."); }
            }
            return ExecuteAsync(call, adapter, ct);
        }

        public async Task<ToolResult> ExecuteAsync(ToolCall call, IHostAdapter adapter, CancellationToken ct)
        {
            EnsureRequestScope(ct);

            if (call == null || !ToolNames.IsWhitelisted(call.Name))
                return ToolResult.Fail("Tool not whitelisted: " + (call != null ? call.Name : "(null)"));

            try
            {
                EnsureRequestScope(ct);
                ReportProgress("Running: " + call.Name);
                ToolResult result = ToolNames.IsWriteTool(call.Name)
                    ? await ExecuteWriteAsync(call, adapter, ct).ConfigureAwait(true)
                    : ExecuteRead(call, adapter, ct);
                ReportProgress((result.Success ? "Completed: " : "Failed: ") + call.Name);
                return result;
            }
            catch (OperationCanceledException)
            {
                // Cancellation/scope loss is a request boundary. Never convert it into a
                // model-visible TOOL ERROR because the caller must abort the entire tool loop.
                throw;
            }
            catch (OmnixException ex)
            {
                return ToolResult.Fail("TOOL ERROR [" + ex.Code + "]: " + ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Error("gateway", "Tool execution failed: " + call.Name, ex);
                return ToolResult.Fail("TOOL ERROR: " + ex.Message);
            }
        }

        private void ReportProgress(string message)
        {
            try { if (Progress != null) Progress(message); }
            catch { /* Display failures must not interrupt Office operations. */ }
        }

        private void EnsureRequestScope(CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();

            var validator = RequestScopeValidator;
            if (validator == null) return;

            bool valid = false;
            try { valid = validator(); }
            catch { valid = false; }
            if (!valid)
                throw new OperationCanceledException("The active Office document changed while the AI request was running.", ct);
        }

        private ToolResult ExecuteRead(ToolCall call, IHostAdapter adapter, CancellationToken ct)
        {
            EnsureRequestScope(ct);

            switch (call.Name)
            {
                case ToolNames.SearchConversation:
                    return ConversationSearch != null
                        ? ToolResult.Ok(ConversationSearch(ToolArguments.Parse(call.ArgumentsJson).Get("query", "")))
                        : ToolResult.Fail("Conversation search is unavailable in this workspace.");
                case ToolNames.SearchOfficeReference:
                {
                    var args = ToolArguments.Parse(call.ArgumentsJson);
                    return ToolResult.Ok(Reference.OfficeReference.Search(args.Get("host", adapter.HostDisplayName), args.Get("query", ""), args.Integer("offset", 0, 0, 10000)));
                }
                case ToolNames.ReadDocumentMap:
                case ToolNames.ReadDocumentSection:
                {
                    var indexed = adapter as IIndexedHostAdapter;
                    if (indexed == null) return ToolResult.Fail("Structured navigation is unavailable for this host.");
                    var args = ToolArguments.Parse(call.ArgumentsJson);
                    EnsureRequestScope(ct);
                    string data = call.Name == ToolNames.ReadDocumentMap
                        ? indexed.ReadDocumentMap(args.Integer("offset", 0, 0, 1000000))
                        : indexed.ReadDocumentSection(args);
                    return ToolResult.Ok(UntrustedData.Wrap("BOUNDED DOCUMENT NAVIGATION RESULT", data));
                }
                case ToolNames.ReadSelection:
                {
                    EnsureRequestScope(ct);
                    string text = adapter.ReadSelection();
                    return ToolResult.Ok(UntrustedData.Wrap("READ_SELECTION RESULT", text));
                }
                case ToolNames.ReadDocument:
                case ToolNames.ReadPresentation:
                {
                    EnsureRequestScope(ct);
                    string text = adapter.ReadDocument(6000);
                    return ToolResult.Ok(UntrustedData.Wrap("READ_DOCUMENT RESULT", text));
                }
                case ToolNames.CaptureChartAsImage:
                {
                    var args = ToolArguments.Parse(call.ArgumentsJson);
                    EnsureRequestScope(ct);
                    byte[] png = adapter.CaptureChartAsImage(args.Get("chart", ""));
                    if (png == null || png.Length == 0) return ToolResult.Fail("No chart found to capture.");
                    return VisionCaptureResult("Excel chart", call.Name, png);
                }
                case ToolNames.CaptureSlideAsImage:
                {
                    var args = ToolArguments.Parse(call.ArgumentsJson);
                    int slide = 0;
                    int.TryParse(args.Get("slide", "0"), out slide);
                    EnsureRequestScope(ct);
                    byte[] png = adapter.CaptureSlideAsImage(slide);
                    if (png == null || png.Length == 0) return ToolResult.Fail("No slide available to capture.");
                    return VisionCaptureResult("PowerPoint slide", call.Name, png);
                }
                case ToolNames.CaptureCurrentViewAsImage:
                {
                    EnsureRequestScope(ct);
                    byte[] png = adapter.CaptureCurrentViewAsImage();
                    if (png == null || png.Length == 0)
                        return ToolResult.Fail("The current Office view could not be captured as an image.");
                    return VisionCaptureResult(adapter.HostDisplayName + " current view/selection", call.Name, png);
                }
                default:
                    return ToolResult.Fail("Unhandled read tool: " + call.Name);
            }
        }

        private static ToolResult VisionCaptureResult(string source, string toolName, byte[] png)
        {
            return new ToolResult
            {
                Success = true,
                ContentForModel = "OMNIX captured the " + source + ". The PNG attached to this tool-result message is the visual source; inspect it directly and combine it with the structured Office context. Do not claim to see anything outside this captured view.",
                UiNote = source + " captured for Vision",
                CapturedPng = png
            };
        }

        private async Task<ToolResult> ExecuteWriteAsync(ToolCall call, IHostAdapter adapter, CancellationToken ct)
        {
            EnsureRequestScope(ct);

            WritePreview preview;
            try
            {
                preview = adapter.PrepareWrite(call.Name, call.ArgumentsJson);
            }
            catch (OmnixException ex)
            {
                return ToolResult.Fail("PREVIEW ERROR [" + ex.Code + "]: " + ex.Message);
            }

            // The active Office document may have changed while PrepareWrite inspected it.
            EnsureRequestScope(ct);

            if (WriteConfirmation == null)
                return ToolResult.Fail("Write confirmation dialog is unavailable; change was NOT applied.");

            bool confirmed;
            try
            {
                confirmed = await WriteConfirmation(preview).ConfigureAwait(true);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "Write confirmation handler failed", ex);
                confirmed = false;
            }

            // Confirmation can be modal and long-lived. Re-check cancellation/scope immediately
            // before any mutation so a document switch cannot redirect an approved old preview.
            EnsureRequestScope(ct);

            if (!confirmed)
            {
                return ToolResult.Fail("The user reviewed the preview and CANCELLED the change. Do not retry the same write without asking why.");
            }

            EnsureRequestScope(ct);
            adapter.ApplyWrite(call.Name, call.ArgumentsJson);
            string hint = call.Name == ToolNames.CreateDataTable
                ? "New worksheet and data table created; headers, cell values and row count verified. To reverse this operation, delete the new worksheet; native Ctrl+Z is not guaranteed."
                : Localization.Strings.T("S.Tools.Applied");
            return ToolResult.Ok("CHANGE APPLIED. " + hint, hint);
        }
    }
}
