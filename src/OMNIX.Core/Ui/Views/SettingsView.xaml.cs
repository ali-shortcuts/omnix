using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
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
            "lmstudio.ai"
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

                var cp = settings.CustomProvider;
                CustomNameBox.Text = cp != null ? cp.Name : "";
                CustomBaseUrlBox.Text = cp != null ? cp.BaseUrl : "";

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

            CustomProviderSection.Visibility = info.Id == "custom" ? Visibility.Visible : Visibility.Collapsed;
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
            ProviderLinkNote.Visibility = hasOfficialLink ? Visibility.Visible : Visibility.Collapsed;
        }

        private static string BuildProviderSummary(ProviderInfo info)
        {
            if (info == null) return string.Empty;
            string access;
            switch (info.AccessProfile)
            {
                case ProviderAccessProfile.LocalNoCost:
                    access = "Access: Local / no cloud token charge.";
                    break;
                case ProviderAccessProfile.FreeTierAvailable:
                    access = "Access: Free tier/mode currently available; provider limits apply.";
                    break;
                case ProviderAccessProfile.FreeModelsAvailable:
                    access = "Access: Free models currently available; provider capacity/limits can change.";
                    break;
                case ProviderAccessProfile.FreeCreditsAvailable:
                    access = "Access: Limited free credits/trial currently available; not unlimited free usage.";
                    break;
                case ProviderAccessProfile.CustomEndpoint:
                    access = "Access: Defined by your custom endpoint.";
                    break;
                case ProviderAccessProfile.AccountDependent:
                    access = "Access: Account/plan dependent; see the provider's current terms.";
                    break;
                default:
                    access = "Access: Unknown until provider/account details are checked.";
                    break;
            }

            string text = access;
            if (!string.IsNullOrWhiteSpace(info.AccessNotes)) text += "\n" + info.AccessNotes;
            if (!string.IsNullOrWhiteSpace(info.Notes)) text += "\n" + info.Notes;
            return text;
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
            ModelCombo.IsEnabled = !busy;
            ApiKeyBox.IsEnabled = !busy;
            CustomNameBox.IsEnabled = !busy;
            CustomBaseUrlBox.IsEnabled = !busy;
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

                TestResultText.Text = models.Count + " models loaded.";
                if (info.AccessProfile == ProviderAccessProfile.FreeModelsAvailable)
                    TestResultText.Text += " Free options are prioritized at the top of the list.";
                if (string.Equals(info.Id, "huggingface", StringComparison.OrdinalIgnoreCase))
                    TestResultText.Text += " Any currently-free provider routes reported by the live Hugging Face catalog are prioritized.";
                TestResultText.Text += "\nSelect a model, then use Test Connection. Model discovery does not verify Vision.";
            }
            catch (OmnixException ex)
            {
                if (ReferenceEquals(_providerOperation, operation))
                    TestResultText.Text = Errors.ErrorPresenter.Format(ex) + "\nYou can still enter a model ID and test it directly.";
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
                var privacy = new PrivacyGate
                {
                    CloudConfirmationCallback = providerName => System.Threading.Tasks.Task.FromResult(OmnixDialogs.ConfirmCloudSend(
                        providerName, "Synthetic connection test only. No Office document content is sent."))
                };
                await ProviderDiagnostics.TestModelAsync(adapter, gateway.Router.BuildCredentials(info.Id), privacy, operation.Token);
                if (!ReferenceEquals(_providerOperation, operation)) return;
                TestResultText.Text = "The selected model returned a text response.\nVision was not tested.\n" + BuildProviderSummary(adapter.Info);
                TestResultText.SetResourceReference(TextBlock.ForegroundProperty, "B.Success");
            }
            catch (OmnixException ex)
            {
                if (!ReferenceEquals(_providerOperation, operation)) return;
                TestResultText.Text = Errors.ErrorPresenter.Format(ex);
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

            if (settings.CustomProvider != null)
            {
                if (!string.Equals(settings.CustomProvider.BaseUrl, CustomBaseUrlBox.Text.Trim(), StringComparison.Ordinal) ||
                    (_displayedProviderId == "custom" && !string.Equals(settings.CustomProvider.Model, ModelCombo.Text.Trim(), StringComparison.Ordinal)))
                    settings.CustomProvider.SupportsVision = null;
                settings.CustomProvider.Name = CustomNameBox.Text.Trim();
                settings.CustomProvider.BaseUrl = CustomBaseUrlBox.Text.Trim();
                if (_displayedProviderId == "custom") settings.CustomProvider.Model = ModelCombo.Text.Trim();
            }

            settings.Models[_displayedProviderId] = ModelCombo.Text.Trim();
            string key = ApiKeyBox.Password;
            if (!string.IsNullOrWhiteSpace(key))
                SettingsManager.Instance.SetApiKey(_displayedProviderId, key.Trim());
        }

        private void SaveGeneralFields()
        {
            var settings = SettingsManager.Instance.Settings;

            if (PrivacyLocalOnly.IsChecked == true) settings.Privacy = PrivacyMode.LocalOnly;
            else if (PrivacyCloudAllowed.IsChecked == true) settings.Privacy = PrivacyMode.CloudAllowed;
            else settings.Privacy = PrivacyMode.AskBeforeSending;

            settings.PreferLocalWhenAvailable = PreferLocalCheck.IsChecked == true;

            int msgs, days;
            settings.HistoryMaxMessages = int.TryParse(MaxMessagesBox.Text, out msgs) ? Math.Max(10, msgs) : 500;
            settings.HistoryMaxAgeDays = int.TryParse(MaxDaysBox.Text, out days) ? Math.Max(1, days) : 30;

            var gateway = Gateway;
            if (gateway != null) gateway.Privacy.ResetSession();
        }
    }
}
