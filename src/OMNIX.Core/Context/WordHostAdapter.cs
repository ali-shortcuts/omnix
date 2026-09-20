using System;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using OMNIX.Core.Errors;
using OMNIX.Core.Tools;
using OMNIX.Core.Util;

namespace OMNIX.Core.Context
{
    /// <summary>
    /// Word adapter (spec Section 3, Layer 3): Document, Selection, Paragraphs, Headings, Tables,
    /// Track Changes/Comments. Write tool: rewrite_selected_text with a native Word UndoRecord.
    ///
    /// Performance invariant: large Word ranges are shortened through Range.Duplicate/End BEFORE
    /// the Text property is requested. Truncating a giant string after doc.Content.Text has already
    /// been materialized defeats the context limit and can pause Word on very large documents.
    /// </summary>
    public sealed class WordHostAdapter : IHostAdapter, IIndexedHostAdapter
    {
        private const int MaxRewriteSelectionChars = 50000;
        private const int MaxRewriteReplacementChars = 50000;
        private const int PreviewChars = 4000;

        private readonly Word.Application _app;
        private readonly Func<int> _maxChars;

        public WordHostAdapter(Word.Application app, Func<int> maxChars)
        {
            _app = app;
            _maxChars = maxChars;
        }

        public HostType Host { get { return HostType.Word; } }
        public string HostDisplayName { get { return "Word"; } }

        public OfficeContext ReadContext()
        {
            var ctx = new OfficeContext { Host = HostType.Word };
            try
            {
                var doc = _app.ActiveDocument;
                if (doc == null) return ctx;
                ctx.DocumentName = doc.Name;
                ctx.DocumentPath = doc.FullName;
                ctx.HasTrackChanges = doc.TrackRevisions;
                ctx.CommentCount = doc.Comments.Count;

                var sel = _app.Selection;
                if (sel != null && sel.Range != null)
                {
                    ctx.SelectionAddress = "chars " + sel.Start + "–" + sel.End;
                    ctx.SelectionText = ReadRangeTextBounded(sel.Range, 200);
                }
                ctx.HeadingsPreview = BuildHeadings(doc);
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "WordHostAdapter.ReadContext failed", ex);
            }
            return ctx;
        }

        public string ReadSelection()
        {
            try
            {
                var sel = _app.Selection;
                if (sel == null || sel.Range == null) return "(no selection)";
                int cap = Math.Max(1, _maxChars());
                string text = ReadRangeTextBounded(sel.Range, cap);
                bool truncated = (sel.End - sel.Start) > cap;
                return "Selection (" + sel.Start + "–" + sel.End + "):\n" + text +
                       (truncated ? Environment.NewLine + "…[selection truncated before Word text materialization]" : "");
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "WordHostAdapter.ReadSelection failed", ex);
                return "(unable to read selection)";
            }
        }

        public string ReadDocument(int maxChars)
        {
            try
            {
                var doc = _app.ActiveDocument;
                if (doc == null) return "(no document open)";

                int cap = Math.Max(1, Math.Min(maxChars, _maxChars()));
                var sb = new StringBuilder();
                sb.AppendLine("Document: " + doc.Name + " (" + doc.Paragraphs.Count + " paragraphs, "
                              + doc.Tables.Count + " tables)");
                if (doc.TrackRevisions) sb.AppendLine("[Track Changes is ON]");
                sb.AppendLine();

                int textBudget = Math.Max(1, cap - Math.Min(512, sb.Length + 128));
                var content = doc.Content;
                int originalLength = Math.Max(0, content.End - content.Start);
                sb.Append(ReadRangeTextBounded(content, textBudget));
                if (originalLength > textBudget)
                    sb.AppendLine(Environment.NewLine + "…[document text truncated before Word text materialization]");

                return TextUtil.Truncate(sb.ToString(), cap);
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "WordHostAdapter.ReadDocument failed", ex);
                return "(unable to read document)";
            }
        }

        public string ReadDocumentMap(int offset)
        {
            var doc = _app.ActiveDocument;
            var body = doc.Content;
            return "Word main story: start=" + body.Start + "; endExclusive=" + body.End
                + "; paragraphs=" + doc.Paragraphs.Count + "; tables=" + doc.Tables.Count
                + "; comments=" + doc.Comments.Count
                + ". Read using start/count. Headers, footers, text boxes, comments and footnotes are NOT included in main-story text.";
        }

        public string ReadDocumentSection(ToolArguments args)
        {
            var doc = _app.ActiveDocument;
            var body = doc.Content;
            int start = args.Integer("start", body.Start, body.Start, body.End);
            int count = args.Integer("count", 4000, 1, 4000);
            int end = Math.Min(body.End, start + count);
            string text = doc.Range(start, end).Text;
            return "Word main story [" + start + "," + end + "); nextStart="
                + (end < body.End ? end.ToString() : "none") + "\n" + text;
        }

        public byte[] CaptureChartAsImage(string chartName) { return null; }

        public byte[] CaptureSlideAsImage(int slideIndexOneBased) { return null; }

        public byte[] CaptureCurrentViewAsImage()
        {
            try
            {
                var sel = _app.Selection;
                if (sel == null || sel.Range == null) return null;
                return TempImageCapture.FromExporter(path =>
                {
                    sel.Range.CopyAsPicture();
                    System.Windows.IDataObject data = System.Windows.Clipboard.GetDataObject();
                    if (data != null && data.GetDataPresent(System.Windows.DataFormats.Bitmap))
                    {
                        var bmp = data.GetData(System.Windows.DataFormats.Bitmap) as System.Windows.Media.Imaging.BitmapSource;
                        if (bmp != null)
                        {
                            var enc = new System.Windows.Media.Imaging.PngBitmapEncoder();
                            enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bmp));
                            using (var fs = System.IO.File.Create(path))
                                enc.Save(fs);
                            return;
                        }
                    }
                    throw new InvalidOperationException("No bitmap on clipboard after CopyAsPicture.");
                });
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "WordHostAdapter.CaptureCurrentViewAsImage failed", ex);
                return null;
            }
        }

        public WritePreview PrepareWrite(string toolName, string argumentsJson)
        {
            if (toolName != ToolNames.RewriteSelectedText)
                throw new OmnixException(ErrorCode.CORE_ERROR,
                    "Tool '" + toolName + "' is not supported by Word.",
                    "WordHostAdapter.PrepareWrite", "Use rewrite_selected_text.");

            var args = ToolArguments.Parse(argumentsJson);
            var sel = RequireBoundedRewriteSelection(toolName);
            string after = RequireBoundedReplacement(args, toolName);
            int selectedChars = Math.Max(0, sel.End - sel.Start);
            string before = ReadRangeTextBounded(sel.Range, PreviewChars);

            return new WritePreview
            {
                ToolName = toolName,
                Title = "Word — rewrite selected text",
                Before = "Selected characters: " + selectedChars + Environment.NewLine +
                         TextUtil.Truncate(before, PreviewChars) +
                         (selectedChars > PreviewChars ? Environment.NewLine + "…[preview truncated; mutation remains bounded to the selected range]" : ""),
                After = "Replacement characters: " + after.Length + Environment.NewLine + TextUtil.Truncate(after, PreviewChars) +
                        (after.Length > PreviewChars ? Environment.NewLine + "…[replacement preview truncated]" : ""),
                ArgumentsJson = argumentsJson
            };
        }

        public void ApplyWrite(string toolName, string argumentsJson)
        {
            if (toolName != ToolNames.RewriteSelectedText)
                throw new OmnixException(ErrorCode.CORE_ERROR, "Unknown Word write tool: " + toolName, "", "");

            var args = ToolArguments.Parse(argumentsJson);
            string newText = RequireBoundedReplacement(args, toolName);
            var sel = RequireBoundedRewriteSelection(toolName);

            Word.UndoRecord undo = null;
            try
            {
                undo = _app.UndoRecord;
                if (undo != null) undo.StartCustomRecord("OMNIX rewrite_selected_text");
                sel.Text = newText;
            }
            finally
            {
                if (undo != null)
                {
                    try { undo.EndCustomRecord(); } catch { }
                }
            }
            Logging.Logger.Install("Word write tool applied: rewrite_selected_text (replacement chars=" + newText.Length + ")");
        }

        private Word.Selection RequireBoundedRewriteSelection(string toolName)
        {
            var sel = _app.Selection;
            if (sel == null || sel.Range == null)
                throw new OmnixException(ErrorCode.CORE_ERROR, "No active selection in Word.", toolName, "Select the text to rewrite first.");

            int length = Math.Max(0, sel.End - sel.Start);
            if (length <= 0)
                throw new OmnixException(ErrorCode.CORE_ERROR,
                    "rewrite_selected_text requires a non-empty selection.",
                    "The current Word selection is only a caret position.",
                    "Select the exact text that OMNIX should rewrite, then try again.");
            if (length > MaxRewriteSelectionChars)
                throw new OmnixException(ErrorCode.CORE_ERROR,
                    "The selected Word range is too large for one AI-approved rewrite.",
                    "Selected characters=" + length + "; limit=" + MaxRewriteSelectionChars + ".",
                    "Select a smaller section and approve it separately.");
            return sel;
        }

        private static string RequireBoundedReplacement(ToolArguments args, string toolName)
        {
            string text = args.Get("text", args.Get("value", "")) ?? "";
            if (text.Length > MaxRewriteReplacementChars)
                throw new OmnixException(ErrorCode.CORE_ERROR,
                    "Replacement text is too large for one AI-approved rewrite.",
                    "Replacement characters=" + text.Length + "; limit=" + MaxRewriteReplacementChars + ".",
                    "Split the rewrite into smaller explicit changes.");
            return text;
        }

        private static string ReadRangeTextBounded(Word.Range range, int maxChars)
        {
            if (range == null) return string.Empty;
            int cap = Math.Max(1, maxChars);
            Word.Range bounded = null;
            try
            {
                bounded = range.Duplicate;
                int desiredEnd = bounded.Start + cap;
                if (bounded.End > desiredEnd) bounded.End = desiredEnd;
                return bounded.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                if (bounded != null)
                {
                    try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(bounded); } catch { }
                }
            }
        }

        private string BuildHeadings(Word.Document doc)
        {
            var sb = new StringBuilder();
            try
            {
                int count = 0;
                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    if (count++ >= 200) { sb.AppendLine("…"); break; }
                    string styleName = null;
                    try { styleName = p.Range.ParagraphStyle != null ? ((Word.Style)p.Range.ParagraphStyle).NameLocal : null; }
                    catch { }
                    if ((!string.IsNullOrEmpty(styleName) && styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) ||
                        (styleName != null && styleName.StartsWith("عنوان", StringComparison.Ordinal)))
                    {
                        string text = ReadRangeTextBounded(p.Range, 300).Trim('\r', '\a');
                        sb.AppendLine(styleName + ": " + text);
                        if (sb.Length >= 2000) { sb.AppendLine("…"); break; }
                    }
                }
            }
            catch { }
            return sb.Length == 0 ? "(no headings)" : TextUtil.Truncate(sb.ToString(), 2200);
        }
    }
}
