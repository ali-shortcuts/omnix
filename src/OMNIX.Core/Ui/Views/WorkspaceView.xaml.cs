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

        public void ShowSettingsTab()
        {
            TabSettings.IsChecked = true;
        }

        public void OnPaneClosing()
        {
            ChatPage.CancelPending();
        }

        private void ShowChat(object sender, RoutedEventArgs e) { if (!_ready) return; ChatPage.Visibility = Visibility.Visible; SettingsPage.Visibility = Visibility.Collapsed; AboutPage.Visibility = Visibility.Collapsed; }
        private void ShowSettings(object sender, RoutedEventArgs e) { if (!_ready) return; ChatPage.Visibility = Visibility.Collapsed; SettingsPage.Visibility = Visibility.Visible; AboutPage.Visibility = Visibility.Collapsed; SettingsPage.OnShown(); }
        private void ShowAbout(object sender, RoutedEventArgs e) { if (!_ready) return; ChatPage.Visibility = Visibility.Collapsed; SettingsPage.Visibility = Visibility.Collapsed; AboutPage.Visibility = Visibility.Visible; }
    }
}
