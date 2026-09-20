using System;
using System.Diagnostics;
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

        private static readonly string[] AllowedOfficialHosts =
        {
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

                ProviderCombo.ItemsSource = registry != null ? registry.All.Select(p => p.Info).ToList() : null;
                var selected = registry != null ? registry.Get(settings.SelectedProviderId) : null;
                _displayedProviderId = selected != null ? selected.Info.Id : null;
                ModelCombo.ItemsSource = null;
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

        private void OnProviderChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loading) return;
            var info = ProviderCombo.SelectedItem as ProviderInfo;
            if (info == null) return;

            CancelProviderOperation();
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
            ModelCombo.ItemsSource = null;
            ModelCombo.Text = settings.Models != null && settings.Models.TryGetValue(info.Id, out model) ? model : info.DefaultModel;

            ApiKeyBox.Clear();
            UpdateProviderUi(info);
            TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
            TestResultText.Text = BuildProviderSummary(info);
        }

        private void SetProviderOperationBusy(bool busy)
        {
            LoadModelsButton.IsEnabled = !busy;
            TestButton.IsEnabled = !busy;
            CancelTestButton.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            ModelCombo.IsEnabled = !busy;
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
                    TestResultText.Text = "The server returned no model catalog. You can enter a model ID and use Test Connection.";
                    return;
                }
                string current = ModelCombo.Text;
                ModelCombo.ItemsSource = ProviderDiagnostics.ModelOptions(models, current);
                ModelCombo.Text = current; // Preserve a manually entered model ID.
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Success");

                TestResultText.Text = models.Count + " models found. Select one or enter an ID.";
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
            SaveGeneralFields();
            SettingsManager.Instance.Save();
            var operation = BeginProviderOperation(30);
            TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
            TestResultText.Text = "Testing…";
            try
            {
                var adapter = new ProviderRegistry().Get(info.Id);
                if (adapter == null) return;
                await ProviderDiagnostics.TestSyntheticModelAsync(adapter, gateway.Router.BuildCredentials(info.Id), operation.Token);
                if (!ReferenceEquals(_providerOperation, operation)) return;
                TestResultText.Text = "Connected. Text response received. Vision not tested.";
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Success");
            }
            catch (OmnixException ex)
            {
                if (!ReferenceEquals(_providerOperation, operation)) return;
                TestResultText.Text = ex.Message;
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
            }
            catch (OperationCanceledException)
            {
                if (!ReferenceEquals(_providerOperation, operation)) return;
                TestResultText.Text = Errors.ErrorPresenter.Format(OmnixException.Timeout(info.DisplayName + " connection test timed out."));
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "TestConnection failed", ex);
                if (!ReferenceEquals(_providerOperation, operation)) return;
                TestResultText.Text = Localization.Strings.T("S.Settings.TestFailed") + " — " + ex.Message;
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
            }
            finally
            {
                EndProviderOperation(operation);
            }
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
            if (_controller != null) _controller.SaveSettingsFromUi();
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
                customConfig.ApiType = CustomApiTypeCombo.SelectedIndex == 1 ? "Anthropic" : "OpenAI";
                if (!string.Equals(customConfig.BaseUrl, CustomBaseUrlBox.Text.Trim(), StringComparison.Ordinal) ||
                    (!string.Equals(customConfig.Model, ModelCombo.Text.Trim(), StringComparison.Ordinal)))
                    customConfig.SupportsVision = null;
                customConfig.Name = CustomNameBox.Text.Trim();
                customConfig.BaseUrl = CustomBaseUrlBox.Text.Trim();
                customConfig.Model = ModelCombo.Text.Trim();
            }

            settings.Models[_displayedProviderId] = ModelCombo.Text.Trim();
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

            settings.PreferLocalWhenAvailable = PreferLocalCheck.IsChecked == true;

            int msgs, days;
            settings.HistoryMaxMessages = int.TryParse(MaxMessagesBox.Text, out msgs) ? Math.Max(10, msgs) : 500;
            settings.HistoryMaxAgeDays = int.TryParse(MaxDaysBox.Text, out days) ? Math.Max(1, days) : 30;

            var gateway = Gateway;
            if (gateway != null && previousPrivacy != settings.Privacy) gateway.Privacy.ResetSession();
        }
    }
}
