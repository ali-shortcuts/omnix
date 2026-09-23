using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using OMNIX.Core.Storage;
using OMNIX.Core.Ui.Markdown;

namespace OMNIX.Core.Ui
{
    /// <summary>
    /// One chat message (spec Section 5): user and AI bubbles with different colors and
    /// alignment; AI answers are rendered as Markdown (real tables, highlighted code blocks,
    /// monospace Excel formulas). Streaming appends text into the existing bubble.
    ///
    /// Deterministic UI Automation IDs are deliberately exposed on the bubble and message body.
    /// Real Office acceptance can therefore prove that a provider response actually reached the
    /// rendered OMNIX workspace instead of mistaking the Ribbon/task-pane shell for a chat PASS.
    /// </summary>
    public sealed class ChatBubble : Border
    {
        private readonly RichTextBox _body;
        private readonly TextBlock _header;
        private readonly Image _image;
        private readonly FlowDocument _doc;
        private string _rawText = "";

        public bool IsUser { get; private set; }

        public ChatBubble(ChatTurn turn)
        {
            IsUser = turn.Role == ChatRole.User;

            System.Windows.Automation.AutomationProperties.SetAutomationId(
                this, IsUser ? "OMNIX.UserBubble" : "OMNIX.AssistantBubble");
            System.Windows.Automation.AutomationProperties.SetName(
                this, IsUser ? "OMNIX user message" : "OMNIX assistant message");

            CornerRadius = new CornerRadius(10);
            Padding = new Thickness(8, 6, 8, 6);
            Margin = new Thickness(IsUser ? 28 : 4, 3, IsUser ? 4 : 28, 3);
            MaxWidth = 420;

            if (IsUser)
            {
                SetResourceReference(BackgroundProperty, "B.BubbleUser");
                HorizontalAlignment = HorizontalAlignment.Right;
            }
            else
            {
                SetResourceReference(BackgroundProperty, "B.BubbleAi");
                HorizontalAlignment = HorizontalAlignment.Left;
            }

            var stack = new StackPanel();

            _header = new TextBlock
            {
                FontSize = 10,
                FontWeight = FontWeights.SemiBold,
                Opacity = 0.8,
                Margin = new Thickness(0, 0, 0, 2)
            };
            System.Windows.Automation.AutomationProperties.SetAutomationId(
                _header, IsUser ? "OMNIX.UserMessageHeader" : "OMNIX.AssistantMessageHeader");
            _header.Text = (IsUser ? Localization.Strings.T("S.Chat.You") : Localization.Strings.T("S.Chat.Assistant"))
                           + "  ·  " + turn.TimestampUtc.ToLocalTime().ToString("HH:mm");
            _header.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
            stack.Children.Add(_header);

            _doc = new FlowDocument
            {
                PagePadding = new Thickness(0),
                ColumnWidth = double.PositiveInfinity
            };
            _doc.SetCurrentValue(System.Windows.Documents.TextElement.FontFamilyProperty, new FontFamily("Segoe UI"));
            _doc.SetCurrentValue(System.Windows.Documents.TextElement.FontSizeProperty, 12.0);
            _doc.SetResourceReference(TextElement.ForegroundProperty,
                IsUser ? "B.BubbleUserForeground" : "B.BubbleAiForeground");

            _body = new RichTextBox
            {
                Document = _doc,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                IsReadOnly = true,
                IsReadOnlyCaretVisible = false,
                Padding = new Thickness(0),
                VerticalContentAlignment = VerticalAlignment.Top,
                Cursor = Cursors.IBeam
            };
            System.Windows.Automation.AutomationProperties.SetAutomationId(
                _body, IsUser ? "OMNIX.UserMessageBody" : "OMNIX.AssistantMessageBody");
            System.Windows.Automation.AutomationProperties.SetName(
                _body, IsUser ? "OMNIX user message body" : "OMNIX assistant message body");

            // Width-adaptive wrapping inside the narrow pane (no horizontal scrolling).
            _body.Document.PageWidth = double.NaN;
            _body.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
            _body.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
            _body.IsDocumentEnabled = true;
            stack.Children.Add(_body);

            _image = new Image
            {
                MaxHeight = 160,
                Margin = new Thickness(0, 4, 0, 0),
                Stretch = Stretch.Uniform,
                Visibility = Visibility.Collapsed
            };
            System.Windows.Automation.AutomationProperties.SetAutomationId(
                _image, IsUser ? "OMNIX.UserMessageImage" : "OMNIX.AssistantMessageImage");
            stack.Children.Add(_image);

            _body.SetResourceReference(Control.ForegroundProperty,
                IsUser ? "B.BubbleUserForeground" : "B.BubbleAiForeground");
            Child = stack;

            if (turn.HasImages && turn.Images[0].PngBytes != null)
                SetImage(turn.Images[0].PngBytes);

            _rawText = turn.Text ?? "";
            Loaded += (sender, args) =>
            {
                ApplyResolvedThemeResources();
                AppendMarkdown(_rawText);
                Theming.ThemeManager.Instance.ThemeChanged += RefreshTheme;
            };
            Unloaded += (sender, args) => Theming.ThemeManager.Instance.ThemeChanged -= RefreshTheme;
        }

        private void RefreshTheme()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (!IsLoaded) return;
                ApplyResolvedThemeResources();
                AppendMarkdown(_rawText);
            }));
        }

        public void SetImage(byte[] png)
        {
            if (png == null) return;
            try
            {
                var image = new System.Windows.Media.Imaging.BitmapImage();
                using (var ms = new System.IO.MemoryStream(png))
                {
                    image.BeginInit();
                    image.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                    image.StreamSource = ms;
                    image.EndInit();
                }
                image.Freeze();
                _image.Source = image;
                _image.Visibility = Visibility.Visible;
            }
            catch { }
        }

        /// <summary>Streaming: appends chunk text and re-renders the markdown view.</summary>
        public void AppendText(string chunk)
        {
            _rawText += chunk ?? "";
            if (!IsLoaded) return;
            AppendMarkdown(_rawText);
            ScrollToEndSafe();
        }

        /// <summary>Incremental plain-text preview; Markdown is rendered once on completion.</summary>
        public void AppendStreamingText(string chunk)
        {
            if (string.IsNullOrEmpty(chunk)) return;
            _rawText += chunk;
            if (!IsLoaded) return;
            var paragraph = _doc.Blocks.LastBlock as Paragraph;
            if (paragraph == null) { paragraph = new Paragraph(); _doc.Blocks.Add(paragraph); }
            paragraph.Inlines.Add(new Run(chunk));
            if (_rawText.Length == chunk.Length)
                _doc.FlowDirection = System.Text.RegularExpressions.Regex.IsMatch(chunk, @"[\u0600-\u06ff]")
                    ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;
            ScrollToEndSafe();
        }

        public void ReplaceText(string fullText)
        {
            _rawText = fullText ?? "";
            if (!IsLoaded) return;
            AppendMarkdown(_rawText);
            ScrollToEndSafe();
        }

        private void AppendMarkdown(string text)
        {
            ApplyResolvedThemeResources();
            _doc.FlowDirection = System.Text.RegularExpressions.Regex.IsMatch(text ?? "", @"^[^A-Za-z\u0600-\u06ff]*[\u0600-\u06ff]")
                ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;
            MarkdownRenderer.Render(_doc, text);
        }

        /// <summary>
        /// FlowDocument resource lookup can be unreliable in VSTO because Office does not always
        /// create a normal WPF Application resource tree. Copy the brushes resolved from the actual
        /// task-pane element into the document itself before every render. This prevents the
        /// RichTextBox/Markdown defaults from silently falling back to black text in dark mode.
        /// </summary>
        private void ApplyResolvedThemeResources()
        {
            string[] keys =
            {
                "B.Foreground", "B.ForegroundDim", "B.Border", "B.Accent", "B.Link",
                "B.CodeBackground", "B.CodeKeyword", "B.CodeString", "B.CodeComment",
                "B.CodeNumber", "B.BubbleUserForeground", "B.BubbleAiForeground"
            };
            foreach (string key in keys)
            {
                object resolved = null;
                try { resolved = TryFindResource(key); } catch { }
                if (resolved is Brush) _doc.Resources[key] = resolved;
            }

            Brush fallback = IsUser ? Brushes.White : Brushes.Gainsboro;
            string foregroundKey = IsUser ? "B.BubbleUserForeground" : "B.BubbleAiForeground";
            Brush foreground = null;
            try { foreground = TryFindResource(foregroundKey) as Brush; } catch { }
            if (foreground == null && _doc.Resources.Contains(foregroundKey))
                foreground = _doc.Resources[foregroundKey] as Brush;
            if (foreground == null) foreground = fallback;

            _doc.Foreground = foreground;
            _body.Foreground = foreground;
        }

        private void ScrollToEndSafe()
        {
            _body.ScrollToEnd();
        }

        private static Brush Find(string key)
        {
            object v = Application.Current != null ? Application.Current.TryFindResource(key) : null;
            return v as Brush ?? Brushes.Transparent;
        }
    }
}
