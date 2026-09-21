using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using OMNIX.Core.Storage;

namespace OMNIX.Core.Ui
{
    /// <summary>
    /// Chat page (spec Section 5): bubbles, real streaming, real Stop (CancellationToken),
    /// New Chat / Copy / Retry / Stop / Clear, attach-image buttons and the context bar.
    /// Designed for 360px width.
    /// </summary>
    public partial class ChatView : UserControl
    {
        private WorkspaceController _controller;
        private ImageAttachment _pendingImage;
        private bool _busy;

        public ChatView()
        {
            InitializeComponent();
        }

        public void Initialize(WorkspaceController controller)
        {
            _controller = controller;
        }

        // ------------------------------------------------------------- message list

        public ChatBubble AppendTurn(ChatTurn turn)
        {
            var bubble = new ChatBubble(turn);
            MessagesPanel.Children.Add(bubble);
            ScrollToBottom();
            return bubble;
        }

        public void ReloadMessages(IEnumerable<ChatTurn> turns)
        {
            TranscriptBox.Text = _controller != null ? _controller.ConversationText() : "";
            MessagesPanel.Children.Clear();
            foreach (var t in turns)
            {
                var bubble = new ChatBubble(t);
                MessagesPanel.Children.Add(bubble);
            }
            ScrollToBottom();
        }

        private void ScrollToBottom()
        {
            MessagesScroll.UpdateLayout();
            MessagesScroll.ScrollToEnd();
        }

        // ------------------------------------------------------------- state

        public void SetBusy(bool busy)
        {
            _busy = busy;
            SendButton.Visibility = busy ? Visibility.Collapsed : Visibility.Visible;
            StopButton.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            NewChatButton.IsEnabled = !busy;
            ClearButton.IsEnabled = !busy;
            RetryButton.IsEnabled = !busy;
        }

        public void SetContextText(string text)
        {
            ContextText.Text = text ?? "—";
        }

        public void SetStatus(string message)
        {
            StatusText.Text = message ?? "";
            StatusText.SetResourceReference(TextBlock.ForegroundProperty, "B.ForegroundDim");
            StatusBorder.Visibility = string.IsNullOrEmpty(message) ? Visibility.Collapsed : Visibility.Visible;
        }

        public void ShowError(string message)
        {
            StatusText.Text = message ?? "";
            StatusText.SetResourceReference(TextBlock.ForegroundProperty, "B.Danger");
            StatusBorder.Visibility = Visibility.Visible;
        }

        private void OnStatusClose(object sender, RoutedEventArgs e)
        {
            SetStatus("");
        }

        // ------------------------------------------------------------- pending image

        public void SetPendingImage(ImageAttachment image)
        {
            _pendingImage = image;
            if (image != null && image.PngBytes != null)
            {
                var bmp = new BitmapImage();
                using (var ms = new System.IO.MemoryStream(image.PngBytes))
                {
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.StreamSource = ms;
                    bmp.EndInit();
                }
                bmp.Freeze();
                PendingImage.Source = bmp;
                PendingImageBorder.Visibility = Visibility.Visible;
            }
            else
            {
                PendingImageBorder.Visibility = Visibility.Collapsed;
            }
        }

        private void OnRemovePendingImage(object sender, RoutedEventArgs e)
        {
            SetPendingImage(null);
        }

        public void CancelPending()
        {
            if (_controller != null) _controller.StopStreaming();
        }

        // ------------------------------------------------------------- events

        private void OnSend(object sender, RoutedEventArgs e)
        {
            Send();
        }

        private void OnStop(object sender, RoutedEventArgs e)
        {
            if (_controller != null) _controller.StopStreaming();
        }

        private void OnNewChat(object sender, RoutedEventArgs e)
        {
            if (_controller != null) _controller.NewChat();
        }

        private void OnTranscript(object sender, RoutedEventArgs e)
        {
            bool show = TranscriptBox.Visibility != Visibility.Visible;
            TranscriptBox.Text = _controller != null ? _controller.ConversationText() : "";
            TranscriptBox.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            MessagesScroll.Visibility = show ? Visibility.Collapsed : Visibility.Visible;
            if (show) TranscriptBox.Focus();
        }

        private void OnImportText(object sender, RoutedEventArgs e)
        {
            var picker = new Microsoft.Win32.OpenFileDialog { Filter = "Text files (*.txt)|*.txt", Multiselect = false };
            if (picker.ShowDialog() != true) return;
            try
            {
                if (new System.IO.FileInfo(picker.FileName).Length > 262144)
                { ShowError("Text file is too large. Split it into smaller files."); return; }
                string content = System.IO.File.ReadAllText(picker.FileName, new System.Text.UTF8Encoding(false, true));
                if (InputBox.Text.Length + content.Length > 65536)
                { ShowError("Combined text exceeds 65,536 characters. Nothing was removed."); return; }
                InputBox.AppendText(content);
                InputBox.Focus();
            }
            catch { ShowError("Could not read this text file. Use UTF-8 text."); }
        }

        private void OnCopy(object sender, RoutedEventArgs e)
        {
            if (_controller != null) _controller.CopyConversation();
        }

        private void OnRetry(object sender, RoutedEventArgs e)
        {
            if (_controller != null) _controller.RetryLast();
        }

        private void OnClear(object sender, RoutedEventArgs e)
        {
            if (_controller != null) _controller.ClearChat();
        }

        private void OnAttachFromDocument(object sender, RoutedEventArgs e)
        {
            if (_controller != null) _controller.AttachImageFromDocument();
        }

        private void OnUploadImage(object sender, RoutedEventArgs e)
        {
            if (_controller != null) _controller.AttachImageFromDisk();
        }

        private void OnInputKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Shift) == 0)
            {
                e.Handled = true;
                Send();
            }
        }

        private void Send()
        {
            if (_controller == null || _busy) return;
            string text = InputBox.Text;
            if (string.IsNullOrWhiteSpace(text) && _pendingImage == null) return;
            if (text.Length > 64 * 1024)
            {
                ShowError("Message exceeds 65,536 characters. Your draft has been preserved; split it into smaller messages.");
                return;
            }
            InputBox.Clear();
            var image = _pendingImage;
            SetPendingImage(null);
            _controller.SendMessage(text, image);
        }
    }
}
