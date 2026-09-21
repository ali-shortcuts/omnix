using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using OMNIX.Core.Storage;

namespace OMNIX.Core.Ui
{
    /// <summary>
    /// Workspace root: three tabs (Chat default, Settings, About) inside a docked task pane.
    /// Never full-screen, never a separate window (Ironclad Rule 5).
    /// </summary>
    public partial class WorkspaceView : UserControl
    {
        private readonly WorkspaceController _controller;
        private bool _ready;

        public WorkspaceView(WorkspaceController controller)
        {
            InitializeComponent();
            _controller = controller;
            ChatPage.Initialize(controller);
            SettingsPage.Initialize(controller);
            _ready = true;
            TabChat.IsChecked = true;
        }

        public ChatView Chat { get { return ChatPage; } }
        public SettingsView Settings { get { return SettingsPage; } }

        public void ShowChatTab() { TabChat.IsChecked = true; }

        public void ShowSettingsTab()
        {
            TabSettings.IsChecked = true;
        }

        public void OnPaneClosing()
        {
            ChatPage.CancelPending();
        }

        private int _referenceOffset;
        private void ShowLearn(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;
            ChatPage.Visibility = SettingsPage.Visibility = AboutPage.Visibility = Visibility.Collapsed;
            LearnPage.Visibility = Visibility.Visible;
            UpdateReference();
        }
        private void ReferenceChanged(object sender, RoutedEventArgs e) { _referenceOffset = 0; UpdateReference(); }
        private void ReferencePrevious(object sender, RoutedEventArgs e) { _referenceOffset = Math.Max(0, _referenceOffset - 40); UpdateReference(); }
        private void ReferenceNext(object sender, RoutedEventArgs e) { _referenceOffset += 40; UpdateReference(); }
        private void UpdateReference()
        {
            if (!_ready || ReferenceText == null) return;
            var host = ReferenceHost.SelectedItem as ComboBoxItem;
            var languageItem = ReferenceLanguage.SelectedItem as ComboBoxItem;
            string language = languageItem != null && languageItem.Tag != null
                ? languageItem.Tag.ToString()
                : "en";
            bool persian = string.Equals(language, "fa", StringComparison.OrdinalIgnoreCase);

            ReferenceText.FlowDirection = persian ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;
            ReferenceQuery.FlowDirection = persian ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;
            ReferenceQuery.ToolTip = persian ? "جستجوی نام تابع یا موضوع" : "Search function name or topic";
            ReferencePreviousButton.Content = persian ? "قبلی" : "Previous";
            ReferenceNextButton.Content = persian ? "بعدی" : "Next";

            ReferenceText.Text = Reference.OfficeReference.Search(
                host != null ? host.Content.ToString() : "Excel",
                ReferenceQuery.Text,
                _referenceOffset,
                language);
        }

        private void ShowChat(object sender, RoutedEventArgs e) { if (!_ready) return; LearnPage.Visibility = Visibility.Collapsed; ChatPage.Visibility = Visibility.Visible; SettingsPage.Visibility = Visibility.Collapsed; AboutPage.Visibility = Visibility.Collapsed; }
        private void ShowSettings(object sender, RoutedEventArgs e) { if (!_ready) return; LearnPage.Visibility = Visibility.Collapsed; ChatPage.Visibility = Visibility.Collapsed; SettingsPage.Visibility = Visibility.Visible; AboutPage.Visibility = Visibility.Collapsed; SettingsPage.OnShown(); }
        private void ShowAbout(object sender, RoutedEventArgs e) { if (!_ready) return; LearnPage.Visibility = Visibility.Collapsed; ChatPage.Visibility = Visibility.Collapsed; SettingsPage.Visibility = Visibility.Collapsed; AboutPage.Visibility = Visibility.Visible; }
    }
}
