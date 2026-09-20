using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace OMNIX.Core.Ui.Dialogs
{
    /// <summary>
    /// Modal dialogs used by the AI Gateway (privacy confirmation, spec Layer 7.5) and the
    /// ToolExecutor (write preview before applying, spec Phase 15). Code-built WPF windows so
    /// no extra XAML resources are needed inside the Office process.
    /// </summary>
    public static class OmnixDialogs
    {
        private static Window BuildShell(string title)
        {
            var w = new Window
            {
                Title = title,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                SizeToContent = SizeToContent.WidthAndHeight,
                ShowInTaskbar = false,
                Topmost = true,
                MinWidth = 380,
                MaxWidth = 460,
                Background = Safe("B.Background"),
                UseLayoutRounding = true
            };
            return w;
        }

        private static Brush Safe(string key)
        {
            object v = Application.Current != null ? Application.Current.TryFindResource(key) : null;
            return v as Brush ?? Brushes.White;
        }

        private static Brush Safe(string key, Brush fallback)
        {
            object v = Application.Current != null ? Application.Current.TryFindResource(key) : null;
            return v as Brush ?? fallback;
        }

        private sealed class ThemeText : TextBlock
        {
            public ThemeText(string text, double size, bool bold = false, bool dim = false)
            {
                Text = text;
                FontSize = size;
                FontWeight = bold ? FontWeights.SemiBold : FontWeights.Normal;
                Foreground = Safe(dim ? "B.ForegroundDim" : "B.Foreground", Brushes.Black);
                TextWrapping = TextWrapping.Wrap;
                Margin = new Thickness(0, 2, 0, 2);
            }
        }

        private static Button MakeButton(string label, bool accent)
        {
            var b = new Button
            {
                Content = label,
                Padding = new Thickness(14, 5, 14, 5),
                Margin = new Thickness(4, 0, 4, 0),
                MinWidth = 84,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            if (accent)
            {
                b.Background = Safe("B.Accent", Brushes.MediumPurple);
                b.Foreground = Safe("B.AccentForeground", Brushes.White);
                b.BorderBrush = b.Background;
            }
            return b;
        }

        /// <summary>
        /// Privacy "Ask before sending" dialog. Returns Tuple(allowed, rememberForSession).
        /// </summary>
        public static Tuple<bool, bool> ConfirmCloudSend(string providerDisplayName, string detailLine)
        {
            bool allowed = false;
            bool remember = false;

            var window = BuildShell(Localization.Strings.T("S.Privacy.ConfirmTitle"));
            var panel = new StackPanel { Margin = new Thickness(16) };

            panel.Children.Add(new ThemeText(
                string.Format(Localization.Strings.T("S.Privacy.ConfirmText"), providerDisplayName), 13, bold: true));
            if (!string.IsNullOrEmpty(detailLine))
                panel.Children.Add(new ThemeText(detailLine, 11, dim: true));

            var rememberBox = new CheckBox
            {
                Content = Localization.Strings.T("S.Privacy.SessionRemember"),
                FontSize = 12,
                Margin = new Thickness(0, 8, 0, 4)
            };
            panel.Children.Add(rememberBox);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            var no = MakeButton("Cancel", false);
            var yes = MakeButton("Allow once", true);
            no.Click += delegate { window.DialogResult = false; window.Close(); };
            yes.Click += delegate { allowed = true; window.DialogResult = true; window.Close(); };
            rememberBox.Checked += delegate { remember = true; };
            rememberBox.Unchecked += delegate { remember = false; };
            buttons.Children.Add(no);
            buttons.Children.Add(yes);
            panel.Children.Add(buttons);

            window.Content = panel;
            window.Owner = GetActiveWindow();
            window.ShowDialog();
            return Tuple.Create(allowed, remember);
        }

        /// <summary>Write tool preview (before/after) — returns true when user applies.</summary>
        public static bool ConfirmWritePreview(Tools.WritePreview preview)
        {
            bool confirmed = false;
            var window = BuildShell(preview.Title ?? Localization.Strings.T("S.Tools.ConfirmTitle"));

            var panel = new StackPanel { Margin = new Thickness(16), MinWidth = 380 };
            panel.Children.Add(new ThemeText(preview.Title, 14, bold: true));

            panel.Children.Add(new ThemeText("Before:", 12, dim: true));
            var before = new TextBox
            {
                Text = preview.Before ?? "",
                IsReadOnly = true,
                FontSize = 11.5,
                TextWrapping = TextWrapping.Wrap,
                MaxHeight = 110,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Background = Safe("B.Surface")
            };
            panel.Children.Add(before);

            panel.Children.Add(new ThemeText("After:", 12, dim: true));
            var after = new TextBox
            {
                Text = preview.After ?? "",
                IsReadOnly = true,
                FontSize = 11.5,
                TextWrapping = TextWrapping.Wrap,
                MaxHeight = 110,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Background = Safe("B.Surface")
            };
            panel.Children.Add(after);

            panel.Children.Add(new ThemeText(preview.ToolName == Tools.ToolNames.CreateDataTable
                ? "To reverse this operation, delete the new worksheet. Ctrl+Z is not guaranteed."
                : Localization.Strings.T("S.Tools.UndoHint"), 10.5, dim: true));

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            var cancel = MakeButton(Localization.Strings.T("S.Tools.ConfirmCancel"), false);
            var apply = MakeButton(Localization.Strings.T("S.Tools.ConfirmApply"), true);
            cancel.Click += delegate { window.DialogResult = false; window.Close(); };
            apply.Click += delegate { confirmed = true; window.DialogResult = true; window.Close(); };
            buttons.Children.Add(cancel);
            buttons.Children.Add(apply);
            panel.Children.Add(buttons);

            window.Content = panel;
            window.Owner = GetActiveWindow();
            window.ShowDialog();
            return confirmed;
        }

        private static Window GetActiveWindow()
        {
            try
            {
                return Application.Current != null ? Application.Current.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive) : null;
            }
            catch
            {
                return null;
            }
        }
    }
}
