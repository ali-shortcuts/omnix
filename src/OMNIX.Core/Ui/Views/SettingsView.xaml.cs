using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using OMNIX.Core.AiGateway;
using OMNIX.Core.Errors;
using OMNIX.Core.Logging;
using OMNIX.Core.Settings;
using OMNIX.Core.Theming;
using OMNIX.Core.Ui.Dialogs;

namespace OMNIX.Core.Ui
{
    /// <summary>
    /// Settings page: provider dropdown, DPAPI-protected API key, dynamic models,
    /// categorized diagnostics, verified official provider setup links, explicit access/free-tier
    /// information, local AI probing, Privacy Mode, appearance and history limits.
    /// </summary>
    public partial class SettingsView : UserControl
    {
        private WorkspaceController _controller;
        private bool _loading;
        private string _displayedProviderId;
        private CancellationTokenSource _providerOperation;
        private readonly List<string> _discoveredModels = new List<string>();
        private readonly Dictionary<string, ModelVerificationResult> _modelVerification =
            new Dictionary<string, ModelVerificationResult>(StringComparer.Ordinal);

        private static readonly string[] AllowedOfficialHosts =
        {
            "docs.siliconflow.com", "cloud.siliconflow.com",
            "developers.cloudflare.com", "dash.cloudflare.com",
            "ai.google.dev",
            "aistudio.google.com",
            "groq.com",
            "console.groq.com",
            "openrouter.ai",
            "mistral.ai",
            "www.mistral.ai",
            "docs.mistral.ai",
            "console.mistral.ai",
            "huggingface.co",
            "www.huggingface.co",
            "cerebras.ai",
            "www.cerebras.ai",
            "inference-docs.cerebras.ai",
            "cloud.cerebras.ai",
            "ollama.com",
            "docs.ollama.com",
            "lmstudio.ai", "cloud.sambanova.ai", "docs.sambanova.ai", "build.nvidia.com", "docs.api.nvidia.com"
        };

        public SettingsView()
        {
            InitializeComponent();
            Unloaded += (sender, args) => CancelProviderOperation();
            ApiKeyBox.PasswordChanged += (sender, args) => InvalidateVerification();
            CustomBaseUrlBox.TextChanged += (sender, args) => InvalidateVerification();
            CloudflareAccountBox.TextChanged += (sender, args) => InvalidateVerification();
            CustomApiTypeCombo.SelectionChanged += (sender, args) => InvalidateVerification();
        }

        private void InvalidateVerification()
        {
            if (_loading) return;
            _modelVerification.Clear();
            VerifiedModelsPanel.Children.Clear();
            ModelVerificationText.Text = "";
            ModelVerificationScroll.Visibility = Visibility.Collapsed;
            WorkingModelsOnlyCheck.IsChecked = false;
            WorkingModelsOnlyCheck.Visibility = Visibility.Collapsed;
        }

        public void Initialize(WorkspaceController controller)
        {
            _controller = controller;
        }

        public void OnShown()
        {
            if (!_loading) LoadFromSettings();
        }

        private void LoadFromSettings()
        {
            CancelProviderOperation();
            _loading = true;
            try
            {
                var settings = SettingsManager.Instance.Settings;
                var registry = _controller != null ? _controller.GatewayRegistry : null;

                ProviderCombo.ItemsSource = registry != null ? registry.All.Select(p => p.Info).OrderBy(p => p.Id == "custom" ? 0 : (p.Id == "ollama" || p.Id == "lmstudio" ? 2 : 1)).ToList() : null;
                var selected = registry != null ? registry.Get(settings.SelectedProviderId) : null;
                _displayedProviderId = selected != null ? selected.Info.Id : null;
                _discoveredModels.Clear();
                _modelVerification.Clear();
                WorkingModelsOnlyCheck.IsChecked = false;
                WorkingModelsOnlyCheck.Visibility = Visibility.Collapsed;
                ModelVerificationScroll.Visibility = Visibility.Collapsed;
                ModelVerificationText.Text = "";
            VerifiedModelsPanel.Children.Clear();
                VerifiedModelsPanel.Children.Clear();
                CloudflareAccountBox.Text = settings.CloudflareAccountId ?? "";
                ConfirmWritesCheck.IsChecked = settings.ConfirmEveryWrite;
                VisibleStepsCheck.IsChecked = settings.ExecutionStepDelayMs > 0;
                BusinessLocaleBox.Text = settings.BusinessLocale ?? "Afghanistan; Dari; currency AFN";
                ModelCombo.ItemsSource = new[] { "Custom Model" };
                ManualModelBox.Visibility = Visibility.Collapsed;
                if (selected != null)
                {
                    ProviderCombo.SelectedItem = selected.Info;
                    UpdateProviderUi(selected.Info);
                    TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
                    TestResultText.Text = BuildProviderSummary(selected.Info);
                }

                string model;
                ModelCombo.Text = settings.Models != null && settings.Models.TryGetValue(settings.SelectedProviderId ?? "", out model) ? model : "";

                ApiKeyBox.Clear();
                if (selected == null) KeyStateText.Visibility = Visibility.Collapsed;

                var cp = settings.EndpointConfig(_displayedProviderId);
                CustomNameBox.Text = cp != null ? cp.Name : "";
                CustomBaseUrlBox.Text = cp != null ? cp.BaseUrl : "";
                CustomApiTypeCombo.SelectedIndex = cp != null && cp.ApiType == "Anthropic" ? 1 : 0;

                PrivacyLocalOnly.IsChecked = settings.Privacy == PrivacyMode.LocalOnly;
                PrivacyCloudAllowed.IsChecked = settings.Privacy == PrivacyMode.CloudAllowed;
                PrivacyAsk.IsChecked = settings.Privacy == PrivacyMode.AskBeforeSending;

                ThemeCombo.SelectedIndex = (int)settings.Theme;
                PreferLocalCheck.IsChecked = settings.PreferLocalWhenAvailable;
                MaxMessagesBox.Text = settings.HistoryMaxMessages.ToString();
                MaxDaysBox.Text = settings.HistoryMaxAgeDays.ToString();

                RefreshModelOptions(ModelCombo.Text);
                UpdateLocalStatusText();
            }
            finally
            {
                _loading = false;
            }
        }

        private AiGateway.AiGateway Gateway
        {
            get { return _controller != null ? _controller.Gateway : null; }
        }

        private void UpdateLocalStatusText()
        {
            var registry = _controller != null ? _controller.GatewayRegistry : null;
            if (registry == null) return;
            string ollama = registry.IsLocalAvailable("ollama") ? Localization.Strings.T("S.Settings.Available") : Localization.Strings.T("S.Settings.NotAvailable");
            string lm = registry.IsLocalAvailable("lmstudio") ? Localization.Strings.T("S.Settings.Available") : Localization.Strings.T("S.Settings.NotAvailable");
            LocalStatusText.Text = "Ollama (11434): " + ollama + "\nLM Studio (1234): " + lm;
        }

        private void UpdateProviderUi(ProviderInfo info)
        {
            if (info == null) return;

            CloudflareSection.Visibility = info.Id == "cloudflare" ? Visibility.Visible : Visibility.Collapsed;
            CustomProviderSection.Visibility = (info.Id == "custom" || info.Id == "agentrouter") ? Visibility.Visible : Visibility.Collapsed;
            LocalProviderSection.Visibility = info.Id == "ollama" || info.Id == "lmstudio" ? Visibility.Visible : Visibility.Collapsed;
            bool needsKey = info.RequiresApiKey;
            bool hasKey = needsKey && SettingsManager.Instance.HasApiKey(info.Id);

            ApiKeyLabel.Visibility = needsKey ? Visibility.Visible : Visibility.Collapsed;
            ApiKeyBox.Visibility = needsKey ? Visibility.Visible : Visibility.Collapsed;

            if (needsKey)
            {
                KeyStateText.SetResourceReference(TextBlock.TextProperty, "S.Settings.ApiKeyStored");
                KeyStateText.Visibility = hasKey ? Visibility.Visible : Visibility.Collapsed;
            }
            else
            {
                KeyStateText.SetResourceReference(TextBlock.TextProperty, "S.Settings.NoApiKeyRequired");
                KeyStateText.Visibility = Visibility.Visible;
            }

            GetApiKeyButton.Visibility = needsKey && !string.IsNullOrWhiteSpace(info.ApiKeyUrl)
                ? Visibility.Visible : Visibility.Collapsed;
            ProviderDocsButton.Visibility = !string.IsNullOrWhiteSpace(info.DocumentationUrl)
                ? Visibility.Visible : Visibility.Collapsed;

            bool hasOfficialLink = GetApiKeyButton.Visibility == Visibility.Visible ||
                                   ProviderDocsButton.Visibility == Visibility.Visible;
            ProviderLinksPanel.Visibility = hasOfficialLink ? Visibility.Visible : Visibility.Collapsed;
            ProviderLinkNote.Visibility = Visibility.Collapsed;
        }

        private static string BuildProviderSummary(ProviderInfo info)
        {
            return string.Empty;
        }

        private string EffectiveModelId
        {
            get { return (ModelCombo.Text == "Custom Model" ? ManualModelBox.Text : ModelCombo.Text).Trim(); }
        }

        private void OnModelChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ManualModelBox == null) return;
            bool manual = string.Equals(ModelCombo.SelectedItem as string, "Custom Model", StringComparison.Ordinal);
            ManualModelBox.Visibility = manual ? Visibility.Visible : Visibility.Collapsed;
            if (manual) ManualModelBox.Focus();
        }

        private void OnProvidersOpened(object sender, EventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (ProviderCombo.Items.Count == 0) return;
                var item = ProviderCombo.ItemContainerGenerator.ContainerFromIndex(0) as FrameworkElement;
                if (item != null) item.BringIntoView();
            }));
        }

        private void OnProviderChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loading) return;
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            if (info == null) return;

            CancelProviderOperation();
            _discoveredModels.Clear();
            _modelVerification.Clear();
            WorkingModelsOnlyCheck.IsChecked = false;
            WorkingModelsOnlyCheck.Visibility = Visibility.Collapsed;
            ModelVerificationScroll.Visibility = Visibility.Collapsed;
            ModelVerificationText.Text = "";
            VerifiedModelsPanel.Children.Clear();
            SaveDisplayedProviderFields();
            var settings = SettingsManager.Instance.Settings;
            _displayedProviderId = info.Id;
            var endpoint = settings.EndpointConfig(info.Id);
            if (endpoint != null) {
                CustomNameBox.Text = endpoint.Name;
                CustomBaseUrlBox.Text = endpoint.BaseUrl;
                CustomApiTypeCombo.SelectedIndex = endpoint.ApiType == "Anthropic" ? 1 : 0;
            }
            settings.SelectedProviderId = info.Id;
            string model;
            ModelCombo.ItemsSource = new[] { "Custom Model" };
                ManualModelBox.Visibility = Visibility.Collapsed;
            ModelCombo.Text = settings.Models != null && settings.Models.TryGetValue(info.Id, out model) ? model : info.DefaultModel;

            RefreshModelOptions(ModelCombo.Text);
            ApiKeyBox.Clear();
            UpdateProviderUi(info);
            TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
            TestResultText.Text = BuildProviderSummary(info);
        }

        private void SetProviderOperationBusy(bool busy)
        {
            LoadModelsButton.IsEnabled = !busy;
            TestButton.IsEnabled = !busy;
            TestModelButton.IsEnabled = !busy;
            VerifyModelsButton.IsEnabled = !busy;
            WorkingModelsOnlyCheck.IsEnabled = !busy;
            CancelTestButton.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            CloudflareAccountBox.IsEnabled = !busy;
            ModelCombo.IsEnabled = !busy;
            ManualModelBox.IsEnabled = !busy;
            ApiKeyBox.IsEnabled = !busy;
            CustomNameBox.IsEnabled = !busy;
            CustomBaseUrlBox.IsEnabled = !busy;
            CustomApiTypeCombo.IsEnabled = !busy;
        }

        private void CancelProviderOperation()
        {
            var operation = _providerOperation;
            _providerOperation = null;
            if (operation != null) operation.Cancel();
            SetProviderOperationBusy(false);
        }

        private CancellationTokenSource BeginProviderOperation(int timeoutSeconds)
        {
            CancelProviderOperation();
            _providerOperation = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));
            SetProviderOperationBusy(true);
            return _providerOperation;
        }

        private void EndProviderOperation(CancellationTokenSource operation)
        {
            if (ReferenceEquals(_providerOperation, operation))
            {
                _providerOperation = null;
                SetProviderOperationBusy(false);
            }
            operation.Dispose();
        }

        private void OnGetApiKey(object sender, RoutedEventArgs e)
        {
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            if (info == null || string.IsNullOrWhiteSpace(info.ApiKeyUrl)) return;
            OpenVerifiedOfficialUrl(info.ApiKeyUrl, info.Id, "API key page");
        }

        private void OnProviderDocs(object sender, RoutedEventArgs e)
        {
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            if (info == null || string.IsNullOrWhiteSpace(info.DocumentationUrl)) return;
            OpenVerifiedOfficialUrl(info.DocumentationUrl, info.Id, "documentation");
        }

        private void OpenVerifiedOfficialUrl(string url, string providerId, string purpose)
        {
            try
            {
                Uri uri;
                if (!Uri.TryCreate(url, UriKind.Absolute, out uri) ||
                    !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
                    !AllowedOfficialHosts.Any(h => string.Equals(uri.Host, h, StringComparison.OrdinalIgnoreCase)))
                {
                    Logger.Error("ui", "Blocked non-official provider URL: " + url, null);
                    TestResultText.Text = "Blocked an unverified provider link. No page was opened.";
                    TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = uri.AbsoluteUri,
                    UseShellExecute = true
                });
                Logger.Gateway("Opened verified official " + purpose + " for provider=" + providerId + " host=" + uri.Host);
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "Failed to open official provider link", ex);
                TestResultText.Text = "Could not open the provider's official page: " + ex.Message;
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
            }
        }

        private async void OnLoadModels(object sender, RoutedEventArgs e)
        {
            try { await OfficeUi.RunAsync(Dispatcher, OnLoadModelsCore); }
            catch (Exception ex) { Logger.Error("ui", "Settings operation failed", ex); }
        }

        private async Task OnLoadModelsCore()
        {
            var gateway = Gateway;
            if (gateway == null) return;
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            if (info == null) return;

            TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
            TestResultText.Text = "Loading models…";
            SaveProviderFields();
            var operation = BeginProviderOperation(25);
            try
            {
                // Diagnostics must not reconfigure the adapter used by an active Office chat.
                var adapter = new ProviderRegistry().Get(info.Id);
                if (adapter == null) return;
                adapter.Configure(gateway.Router.BuildCredentials(info.Id));
                var models = await adapter.ListModelsAsync(operation.Token);
                if (!ReferenceEquals(_providerOperation, operation)) return;
                if (models == null || models.Count == 0)
                {
                    TestResultText.Text = "The server returned no model catalog. You can enter an exact model ID and use Test model.";
                    return;
                }
                string current = EffectiveModelId;
                _discoveredModels.Clear();
                _discoveredModels.AddRange(models.Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.Ordinal));
                _modelVerification.Clear();
                WorkingModelsOnlyCheck.IsChecked = false;
                WorkingModelsOnlyCheck.Visibility = Visibility.Collapsed;
                ModelVerificationScroll.Visibility = Visibility.Collapsed;
                ModelVerificationText.Text = "";
            VerifiedModelsPanel.Children.Clear();
                RefreshModelOptions(current);
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Success");

                TestResultText.Text = models.Count + " models detected. Detection is only a catalog result; use Test model or Verify models to prove inference.";
            }
            catch (OmnixException ex)
            {
                if (ReferenceEquals(_providerOperation, operation))
                    TestResultText.Text = ex.Message + " Enter a model ID manually.";
            }
            catch (OperationCanceledException)
            {
                if (ReferenceEquals(_providerOperation, operation))
                    TestResultText.Text = "Model discovery timed out. Your model ID was kept.";
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "LoadModels failed", ex);
                if (ReferenceEquals(_providerOperation, operation))
                    TestResultText.Text = "Loading models failed: " + ex.Message;
            }
            finally
            {
                EndProviderOperation(operation);
            }
        }

        private async void OnTestConnection(object sender, RoutedEventArgs e)
        {
            try { await OfficeUi.RunAsync(Dispatcher, OnTestConnectionCore); }
            catch (Exception ex) { Logger.Error("ui", "Settings operation failed", ex); }
        }

        private async Task OnTestConnectionCore()
        {
            var gateway = Gateway;
            if (gateway == null) return;
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            if (info == null) return;

            SaveProviderFields();
            var operation = BeginProviderOperation(30);
            TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
            TestResultText.Text = "Testing provider connection/authentication without a model…";
            try
            {
                var adapter = new ProviderRegistry().Get(info.Id);
                if (adapter == null) return;
                var credentials = gateway.Router.BuildCredentials(info.Id);
                var result = await ProviderDiagnostics.TestConnectionOnlyAsync(adapter, credentials, operation.Token);
                if (!ReferenceEquals(_providerOperation, operation)) return;

                TestResultText.Text = result.Summary + " (" + result.LatencyMs + " ms)";
                if (result.State == ConnectionDiagnosticState.Connected)
                    TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Success");
                else if (result.EndpointReachable)
                    TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
                else
                    TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
            }
            catch (OperationCanceledException)
            {
                if (ReferenceEquals(_providerOperation, operation))
                {
                    TestResultText.Text = "Connection test cancelled or timed out. No model was tested.";
                    TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "Connection test failed", ex);
                if (ReferenceEquals(_providerOperation, operation))
                {
                    TestResultText.Text = ErrorPresenter.Format(ex);
                    TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
                }
            }
            finally
            {
                EndProviderOperation(operation);
            }
        }

        private async void OnTestModel(object sender, RoutedEventArgs e)
        {
            try { await OfficeUi.RunAsync(Dispatcher, OnTestModelCore); }
            catch (Exception ex) { Logger.Error("ui", "Settings operation failed", ex); }
        }

        private async Task OnTestModelCore()
        {
            var gateway = Gateway;
            if (gateway == null) return;
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            if (info == null) return;

            SaveProviderFields();
            string model = EffectiveModelId;
            if (string.IsNullOrWhiteSpace(model))
            {
                TestResultText.Text = "Select or enter a model ID first. Test connection does not require a model; Test model does.";
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
                return;
            }

            var operation = BeginProviderOperation(35);
            TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
            TestResultText.Text = "Testing model '" + model + "' with a tiny document-free OMNIX tool-calling probe…";
            try
            {
                var credentials = gateway.Router.BuildCredentials(info.Id);
                credentials.Model = model;
                var result = await ProviderDiagnostics.TestSelectedModelAsync(info.Id, credentials, operation.Token);
                if (!ReferenceEquals(_providerOperation, operation)) return;

                _modelVerification[model] = result;
                RenderVerificationSummary(1, 1);
                WorkingModelsOnlyCheck.Visibility = Visibility.Visible;
                ModelVerificationScroll.Visibility = Visibility.Visible;
                TestResultText.Text = model + ": " + result.Summary;
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty,
                    result.Working ? "B.Success" :
                    result.State == ModelVerificationState.TextOnly ? "B.ForegroundDim" : "B.Danger");
            }
            catch (OperationCanceledException)
            {
                if (ReferenceEquals(_providerOperation, operation))
                {
                    TestResultText.Text = "Model test cancelled or timed out.";
                    TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
                }
            }
            finally
            {
                EndProviderOperation(operation);
            }
        }

        private async void OnVerifyModels(object sender, RoutedEventArgs e)
        {
            try { await OfficeUi.RunAsync(Dispatcher, OnVerifyModelsCore); }
            catch (Exception ex) { Logger.Error("ui", "Settings operation failed", ex); }
        }

        private async Task OnVerifyModelsCore()
        {
            var gateway = Gateway;
            if (gateway == null) return;
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            if (info == null) return;

            if (_discoveredModels.Count == 0)
            {
                TestResultText.Text = "Detect models first. Verification tests only the catalog returned for this exact provider/API configuration.";
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
                return;
            }

            SaveProviderFields();
            var operation = BeginProviderOperation(300);
            _modelVerification.Clear();
            WorkingModelsOnlyCheck.Visibility = Visibility.Collapsed;
            ModelVerificationScroll.Visibility = Visibility.Visible;
            ModelVerificationText.Text = "";
            VerifiedModelsPanel.Children.Clear();
            TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
            TestResultText.Text = "Testing models…";

            try
            {
                var credentials = gateway.Router.BuildCredentials(info.Id);
                var results = await ProviderDiagnostics.VerifyModelsAsync(
                    info.Id,
                    credentials,
                    _discoveredModels,
                    (result, completed, total) =>
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (!ReferenceEquals(_providerOperation, operation)) return;
                            _modelVerification[result.ModelId] = result;
                            RenderVerificationSummary(completed, total);
                        }));
                    },
                    operation.Token);

                if (!ReferenceEquals(_providerOperation, operation)) return;
                foreach (var result in results) _modelVerification[result.ModelId] = result;
                RenderVerificationSummary(results.Count, Math.Min(100, _discoveredModels.Count));
                WorkingModelsOnlyCheck.Visibility = Visibility.Visible;

                int working = results.Count(x => x.Working);
                int textOnly = results.Count(x => x.State == ModelVerificationState.TextOnly);
                TestResultText.Text = "Verification complete: " + working + " OMNIX tool-compatible, " +
                                      textOnly + " text-only, " + results.Count + " tested.";
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty,
                    working > 0 ? "B.Success" : textOnly > 0 ? "B.ForegroundDim" : "B.Danger");
                RefreshModelOptions(EffectiveModelId);
            }
            catch (OperationCanceledException)
            {
                if (ReferenceEquals(_providerOperation, operation))
                {
                    TestResultText.Text = "Stopped. Results kept.";
                    TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
                    WorkingModelsOnlyCheck.Visibility = _modelVerification.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
                    RenderVerificationSummary(_modelVerification.Count, Math.Min(100, _discoveredModels.Count));
                }
            }
            finally
            {
                EndProviderOperation(operation);
            }
        }

        private void OnWorkingModelsFilterChanged(object sender, RoutedEventArgs e)
        {
            if (_loading) return;
            RefreshModelOptions(EffectiveModelId);
        }

        private void RefreshModelOptions(string current)
        {
            var settings = SettingsManager.Instance.Settings;
            List<string> saved;
            IEnumerable<string> source = _discoveredModels;
            if (settings.SavedModels != null && settings.SavedModels.TryGetValue(_displayedProviderId ?? "", out saved))
                source = source.Concat(saved);
            if (WorkingModelsOnlyCheck.IsChecked == true)
                source = source.Where(id =>
                {
                    ModelVerificationResult result;
                    return _modelVerification.TryGetValue(id, out result) && result.Working;
                });

            string catalogCurrent = WorkingModelsOnlyCheck.IsChecked == true ? null : current;
            var options = new[] { "Custom Model" }
                .Concat(ProviderDiagnostics.ModelOptions(source, catalogCurrent))
                .Distinct(StringComparer.Ordinal)
                .ToList();
            ModelCombo.ItemsSource = options;
            // The editable text is preserved even if the working-only dropdown hides this ID.
            // This avoids silently changing the user's configured model.
            ModelCombo.Text = current ?? "";
        }

        private void RenderVerificationSummary(int completed, int total)
        {
            var ordered = _modelVerification.Values
                .OrderByDescending(x => x.Working)
                .ThenBy(x => x.ModelId, StringComparer.OrdinalIgnoreCase)
                .Take(100)
                .ToList();

            int working = ordered.Count(x => x.Working);
            int textOnly = ordered.Count(x => x.State == ModelVerificationState.TextOnly);
            int denied = ordered.Count(x => x.State == ModelVerificationState.AccessDenied);
            int unavailable = ordered.Count(x => x.State == ModelVerificationState.NotFoundOrUnavailable);
            int limited = ordered.Count(x => x.State == ModelVerificationState.RateLimited);
            int incompatible = ordered.Count(x => x.State == ModelVerificationState.Incompatible);
            int timedOut = ordered.Count(x => x.State == ModelVerificationState.TimedOut);

            var lines = new List<string>
            {
                "Progress " + completed + "/" + total + " · tool-compatible=" + working +
                " · text-only=" + textOnly + " · denied=" + denied + " · unavailable=" + unavailable +
                " · rate-limited=" + limited + " · incompatible=" + incompatible +
                " · timeout=" + timedOut
            };
            ModelVerificationText.Text = completed + "/" + total + " · tools " + working + " · text " + textOnly;
            VerifiedModelsPanel.Children.Clear();
            foreach (var result in ordered)
            {
                var captured = result;
                bool selectable = result.Working || result.State == ModelVerificationState.TextOnly;
                var row = new StackPanel { Orientation = Orientation.Horizontal };
                var keep = new CheckBox { VerticalAlignment = VerticalAlignment.Center, IsEnabled = selectable,
                    ToolTip = "Keep this model", Margin = new Thickness(0, 0, 6, 0) };
                var settings = SettingsManager.Instance.Settings;
                List<string> saved;
                keep.IsChecked = settings.SavedModels != null &&
                    settings.SavedModels.TryGetValue(_displayedProviderId ?? "", out saved) && saved.Contains(result.ModelId);
                keep.Checked += (sender, args) => RememberModel(captured.ModelId, true);
                keep.Unchecked += (sender, args) => RememberModel(captured.ModelId, false);
                var select = new Button { Content = result.ModelId, IsEnabled = selectable,
                    ToolTip = result.State + (result.ToolCallingVerified ? " · tools" : " · no verified tools"),
                    Margin = new Thickness(0, 2, 0, 2) };
                select.SetResourceReference(Control.ForegroundProperty, "B.Foreground");
                select.SetResourceReference(Control.BackgroundProperty, "B.Surface");
                select.Click += (sender, args) =>
                {
                    RememberModel(captured.ModelId, true);
                    ModelCombo.Text = captured.ModelId;
                    ManualModelBox.Visibility = Visibility.Collapsed;
                    TestResultText.Text = "Selected: " + captured.ModelId;
                };
                row.Children.Add(keep);
                row.Children.Add(select);
                var state = new TextBlock { Text = "  " + result.State, VerticalAlignment = VerticalAlignment.Center };
                state.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
                row.Children.Add(state);
                VerifiedModelsPanel.Children.Add(row);
            }
            ModelVerificationScroll.Visibility = ordered.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void RememberModel(string model, bool keep)
        {
            var settings = SettingsManager.Instance.Settings;
            if (settings.SavedModels == null) settings.SavedModels = new Dictionary<string, List<string>>();
            List<string> models;
            if (!settings.SavedModels.TryGetValue(_displayedProviderId, out models))
                settings.SavedModels[_displayedProviderId] = models = new List<string>();
            if (keep && !models.Contains(model)) models.Add(model);
            if (!keep) models.Remove(model);
            RefreshModelOptions(EffectiveModelId);
        }

        private async void OnProbeLocal(object sender, RoutedEventArgs e)
        {
            try { await OfficeUi.RunAsync(Dispatcher, OnProbeLocalCore); }
            catch (Exception ex) { Logger.Error("ui", "Settings operation failed", ex); }
        }

        private async Task OnProbeLocalCore()
        {
            var gateway = Gateway;
            if (gateway == null) return;
            ProbeButton.IsEnabled = false;
            LocalStatusText.Text = "Probing local AI…";
            try { await gateway.ProbeLocalAsync(); } catch { }
            UpdateLocalStatusText();
            ProbeButton.IsEnabled = true;
        }

        private void OnThemeChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loading) return;
            var item = ThemeCombo.SelectedItem as ComboBoxItem;
            if (item == null) return;
            var settings = SettingsManager.Instance.Settings;
            settings.Theme = (ThemeMode)ThemeCombo.SelectedIndex;
            SettingsManager.Instance.Save();
            Theming.ThemeManager.Instance.ApplyTo(ParentWorkspace());
            Theming.ThemeManager.Instance.NotifySettingsChanged();
        }

        private System.Windows.FrameworkElement ParentWorkspace()
        {
            DependencyObject d = this;
            while (d != null && !(d is WorkspaceView)) d = System.Windows.Media.VisualTreeHelper.GetParent(d);
            return d as WorkspaceView;
        }

        private void OnCancelTest(object sender, RoutedEventArgs e) { CancelProviderOperation(); TestResultText.Text = "Cancelled."; }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            SaveProviderFields();
            SaveGeneralFields();
            SettingsManager.Instance.Save();
            if (!SettingsManager.Instance.LastSaveSucceeded)
            {
                TestResultText.Text = "Could not save settings. Please try again.";
                return;
            }
            if (_controller != null) _controller.SaveSettingsFromUi();
            var workspace = ParentWorkspace() as WorkspaceView;
            if (workspace != null) workspace.ShowChatTab();
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            if (info != null) UpdateProviderUi(info);
        }

        private void SaveProviderFields()
        {
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            var settings = SettingsManager.Instance.Settings;

            SaveDisplayedProviderFields();
            if (info != null) settings.SelectedProviderId = info.Id;
            SettingsManager.Instance.Save();
        }

        private void SaveDisplayedProviderFields()
        {
            if (string.IsNullOrEmpty(_displayedProviderId)) return;
            var settings = SettingsManager.Instance.Settings;

            var customConfig = settings.EndpointConfig(_displayedProviderId);
            if (customConfig != null && (_displayedProviderId == "custom" || _displayedProviderId == "agentrouter"))
            {
                string apiType = CustomApiTypeCombo.SelectedIndex == 1 ? "Anthropic" : "OpenAI";
                if (!string.Equals(customConfig.ApiType, apiType, StringComparison.Ordinal) ||
                    !string.IsNullOrWhiteSpace(ApiKeyBox.Password) || !string.Equals(customConfig.BaseUrl, CustomBaseUrlBox.Text.Trim(), StringComparison.Ordinal) ||
                    (!string.Equals(customConfig.Model, EffectiveModelId, StringComparison.Ordinal)))
                    customConfig.SupportsVision = null;
                customConfig.ApiType = apiType;
                customConfig.Name = CustomNameBox.Text.Trim();
                customConfig.BaseUrl = CustomBaseUrlBox.Text.Trim();
                customConfig.Model = EffectiveModelId;
            }

            if (_displayedProviderId == "cloudflare") settings.CloudflareAccountId = CloudflareAccountBox.Text.Trim();
            settings.Models[_displayedProviderId] = EffectiveModelId;
            string key = ApiKeyBox.Password;
            if (!string.IsNullOrWhiteSpace(key))
                SettingsManager.Instance.SetApiKey(_displayedProviderId, key.Trim());
        }

        private void SaveGeneralFields()
        {
            var settings = SettingsManager.Instance.Settings;
            var previousPrivacy = settings.Privacy;

            if (PrivacyLocalOnly.IsChecked == true) settings.Privacy = PrivacyMode.LocalOnly;
            else if (PrivacyCloudAllowed.IsChecked == true) settings.Privacy = PrivacyMode.CloudAllowed;
            else settings.Privacy = PrivacyMode.AskBeforeSending;

            settings.ConfirmEveryWrite = ConfirmWritesCheck.IsChecked == true;
            settings.ExecutionStepDelayMs = VisibleStepsCheck.IsChecked == true ? 350 : 0;
            settings.BusinessLocale = BusinessLocaleBox.Text.Trim();
            settings.PreferLocalWhenAvailable = PreferLocalCheck.IsChecked == true;

            int msgs, days;
            settings.HistoryMaxMessages = int.TryParse(MaxMessagesBox.Text, out msgs) ? Math.Max(10, msgs) : 500;
            settings.HistoryMaxAgeDays = int.TryParse(MaxDaysBox.Text, out days) ? Math.Max(1, days) : 30;

            var gateway = Gateway;
            if (gateway != null && previousPrivacy != settings.Privacy) gateway.Privacy.ResetSession();
        }
    }
}
