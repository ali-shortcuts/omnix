using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using OMNIX.Core.Tools;
using OMNIX.Core.Util;

namespace OMNIX.Core.Context
{
    /// <summary>
    /// Read-only Excel decorator. Existing ExcelHostAdapter behavior remains the authority for
    /// context, images and mutations; this layer adds bounded native Range.Find retrieval.
    /// </summary>
    public sealed class SearchableExcelHostAdapter : IHostAdapter, IDocumentSearchProvider
    {
        private readonly ExcelHostAdapter _inner;
        private readonly Excel.Application _app;
        private readonly Func<int> _maxChars;

        public SearchableExcelHostAdapter(Excel.Application app, Func<int> maxCells, Func<int> maxChars)
        {
            if (app == null) throw new ArgumentNullException("app");
            _app = app;
            _maxChars = maxChars ?? delegate { return 6000; };
            _inner = new ExcelHostAdapter(app, maxCells, maxChars);
        }

        public HostType Host { get { return _inner.Host; } }
        public string HostDisplayName { get { return _inner.HostDisplayName; } }
        public OfficeContext ReadContext() { return _inner.ReadContext(); }
        public string ReadSelection() { return _inner.ReadSelection(); }
        public string ReadDocument(int maxChars) { return _inner.ReadDocument(maxChars); }
        public byte[] CaptureChartAsImage(string chartName) { return _inner.CaptureChartAsImage(chartName); }
        public byte[] CaptureSlideAsImage(int slideIndexOneBased) { return _inner.CaptureSlideAsImage(slideIndexOneBased); }
        public byte[] CaptureCurrentViewAsImage() { return _inner.CaptureCurrentViewAsImage(); }
        public WritePreview PrepareWrite(string toolName, string argumentsJson) { return _inner.PrepareWrite(toolName, argumentsJson); }
        public void ApplyWrite(string toolName, string argumentsJson) { _inner.ApplyWrite(toolName, argumentsJson); }

        public string SearchDocument(string query, int maxResults, int maxChars)
        {
            query = (query ?? string.Empty).Trim();
            if (query.Length == 0) return "(empty search query)";

            int resultCap = Math.Max(1, Math.Min(20, maxResults));
            int charCap = Math.Max(256, Math.Min(maxChars, Math.Max(256, _maxChars())));
            var wb = _app.ActiveWorkbook;
            if (wb == null) return "(no workbook open)";

            string findText = EscapeExcelFindText(query);
            var output = new StringBuilder();
            output.AppendLine("Workbook search: \"" + TextUtil.Truncate(query, 200) + "\"");
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int count = 0;

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (count >= resultCap || output.Length >= charCap) break;
                Excel.Range used = null;
                try
                {
                    used = ws.UsedRange;
                    FindExcelMatches(used, ws.Name, findText, query, Excel.XlFindLookIn.xlValues,
                        seen, output, resultCap, charCap, ref count);
                    if (count < resultCap && output.Length < charCap)
                        FindExcelMatches(used, ws.Name, findText, query, Excel.XlFindLookIn.xlFormulas,
                            seen, output, resultCap, charCap, ref count);
                }
                catch (Exception ex)
                {
                    Logging.Logger.Error("startup-debug", "Excel search skipped sheet '" + SafeLabel(ws.Name, 80) + "'", ex);
                }
                finally
                {
                    ReleaseCom(used);
                }
            }

            if (count == 0) output.AppendLine("No matches found.");
            else if (count >= resultCap) output.AppendLine("…[result limit reached: " + resultCap + "]");
            return TextUtil.Truncate(output.ToString(), charCap);
        }

        private static void FindExcelMatches(
            Excel.Range used,
            string sheetName,
            string findText,
            string originalQuery,
            Excel.XlFindLookIn lookIn,
            HashSet<string> seen,
            StringBuilder output,
            int resultCap,
            int charCap,
            ref int count)
        {
            if (used == null || count >= resultCap || output.Length >= charCap) return;

            Excel.Range found = null;
            try
            {
                found = used.Find(findText, Type.Missing, lookIn, Excel.XlLookAt.xlPart,
                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext,
                    false, Type.Missing, false);
                if (found == null) return;

                string firstAddress = SafeExcelAddress(found);
                int safety = 0;
                while (found != null && count < resultCap && output.Length < charCap && safety++ < 10000)
                {
                    string address = SafeExcelAddress(found);
                    string key = sheetName + "!" + address;
                    if (seen.Add(key))
                    {
                        string value = SafeCellText(delegate { return found.Value2; });
                        string formula = SafeCellText(delegate { return found.Formula; });
                        string line = "• " + SafeLabel(sheetName, 100) + "!" + address +
                                      " | value: " + TextUtil.Truncate(NormalizeInline(value), 240);
                        if (!string.IsNullOrEmpty(formula) && !string.Equals(formula, value, StringComparison.Ordinal))
                            line += " | formula: " + TextUtil.Truncate(NormalizeInline(formula), 240);
                        AppendBoundedLine(output, line, charCap);
                        count++;
                    }

                    Excel.Range next = null;
                    try
                    {
                        next = used.Find(findText, found, lookIn, Excel.XlLookAt.xlPart,
                            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext,
                            false, Type.Missing, false);
                    }
                    catch { next = null; }

                    if (next == null)
                    {
                        ReleaseCom(found);
                        found = null;
                        break;
                    }

                    string nextAddress = SafeExcelAddress(next);
                    ReleaseCom(found);
                    found = next;
                    if (string.Equals(nextAddress, firstAddress, StringComparison.OrdinalIgnoreCase)) break;
                }
            }
            finally
            {
                ReleaseCom(found);
            }
        }

        private static string EscapeExcelFindText(string value)
        {
            return (value ?? string.Empty).Replace("~", "~~").Replace("*", "~*").Replace("?", "~?");
        }

        private static string SafeExcelAddress(Excel.Range range)
        {
            try { return range != null ? range.Address[false, false] : "?"; }
            catch { return "?"; }
        }

        private static string SafeCellText(Func<object> getter)
        {
            try
            {
                object value = getter != null ? getter() : null;
                return value == null ? string.Empty : (Convert.ToString(value) ?? string.Empty);
            }
            catch { return string.Empty; }
        }

        private static void ReleaseCom(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return;
            try { Marshal.FinalReleaseComObject(value); } catch { }
        }

        private static string SafeLabel(string value, int maxChars)
        {
            value = value ?? string.Empty;
            return value.Length <= maxChars ? value : value.Substring(0, maxChars);
        }

        private static string NormalizeInline(string value)
        {
            return (value ?? string.Empty).Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ').Trim();
        }

        private static void AppendBoundedLine(StringBuilder output, string line, int charCap)
        {
            if (output == null || output.Length >= charCap) return;
            int remaining = charCap - output.Length;
            string text = (line ?? string.Empty) + Environment.NewLine;
            if (text.Length > remaining) text = text.Substring(0, remaining);
            output.Append(text);
        }
    }

    /// <summary>
    /// Read-only Word decorator. Uses native Range.Find with wdFindStop and advances the range after
    /// every match so large documents can be searched without materializing Document.Content.Text.
    /// </summary>
    public sealed class SearchableWordHostAdapter : IHostAdapter, IDocumentSearchProvider
    {
        private readonly WordHostAdapter _inner;
        private readonly Word.Application _app;
        private readonly Func<int> _maxChars;

        public SearchableWordHostAdapter(Word.Application app, Func<int> maxChars)
        {
            if (app == null) throw new ArgumentNullException("app");
            _app = app;
            _maxChars = maxChars ?? delegate { return 6000; };
            _inner = new WordHostAdapter(app, maxChars);
        }

        public HostType Host { get { return _inner.Host; } }
        public string HostDisplayName { get { return _inner.HostDisplayName; } }
        public OfficeContext ReadContext() { return _inner.ReadContext(); }
        public string ReadSelection() { return _inner.ReadSelection(); }
        public string ReadDocument(int maxChars) { return _inner.ReadDocument(maxChars); }
        public byte[] CaptureChartAsImage(string chartName) { return _inner.CaptureChartAsImage(chartName); }
        public byte[] CaptureSlideAsImage(int slideIndexOneBased) { return _inner.CaptureSlideAsImage(slideIndexOneBased); }
        public byte[] CaptureCurrentViewAsImage() { return _inner.CaptureCurrentViewAsImage(); }
        public WritePreview PrepareWrite(string toolName, string argumentsJson) { return _inner.PrepareWrite(toolName, argumentsJson); }
        public void ApplyWrite(string toolName, string argumentsJson) { _inner.ApplyWrite(toolName, argumentsJson); }

        public string SearchDocument(string query, int maxResults, int maxChars)
        {
            query = (query ?? string.Empty).Trim();
            if (query.Length == 0) return "(empty search query)";

            var doc = _app.ActiveDocument;
            if (doc == null) return "(no document open)";

            int resultCap = Math.Max(1, Math.Min(20, maxResults));
            int charCap = Math.Max(256, Math.Min(maxChars, Math.Max(256, _maxChars())));
            int documentStart = doc.Content.Start;
            int documentEnd = doc.Content.End;
            var output = new StringBuilder();
            output.AppendLine("Document search: \"" + TextUtil.Truncate(query, 200) + "\"");

            Word.Range searchRange = null;
            int count = 0;
            int safety = 0;
            try
            {
                searchRange = doc.Range(documentStart, documentEnd);
                while (count < resultCap && output.Length < charCap && searchRange != null && safety++ < 10000)
                {
                    Word.Find finder = searchRange.Find;
                    finder.ClearFormatting();
                    finder.Text = query;
                    finder.Forward = true;
                    finder.Wrap = Word.WdFindWrap.wdFindStop;
                    finder.Format = false;
                    finder.MatchCase = false;
                    finder.MatchWholeWord = false;
                    finder.MatchWildcards = false;

                    bool matched = finder.Execute();
                    if (!matched) break;

                    int matchStart = searchRange.Start;
                    int matchEnd = searchRange.End;
                    int snippetStart = Math.Max(documentStart, matchStart - 120);
                    int snippetEnd = Math.Min(documentEnd, matchEnd + 180);
                    Word.Range snippet = null;
                    string text = string.Empty;
                    try
                    {
                        snippet = doc.Range(snippetStart, snippetEnd);
                        text = snippet.Text ?? string.Empty;
                    }
                    finally { ReleaseWordCom(snippet); }

                    AppendWordLine(output,
                        "• chars " + matchStart + "–" + matchEnd + ": " +
                        TextUtil.Truncate(NormalizeWordText(text), 420), charCap);
                    count++;

                    int nextStart = Math.Max(matchEnd, matchStart + 1);
                    if (nextStart >= documentEnd) break;
                    Word.Range next = doc.Range(nextStart, documentEnd);
                    ReleaseWordCom(searchRange);
                    searchRange = next;
                }
            }
            finally { ReleaseWordCom(searchRange); }

            if (count == 0) output.AppendLine("No matches found.");
            else if (count >= resultCap) output.AppendLine("…[result limit reached: " + resultCap + "]");
            return TextUtil.Truncate(output.ToString(), charCap);
        }

        private static string NormalizeWordText(string value)
        {
            return (value ?? string.Empty)
                .Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ').Replace('\a', ' ')
                .Trim();
        }

        private static void AppendWordLine(StringBuilder output, string line, int charCap)
        {
            if (output == null || output.Length >= charCap) return;
            int remaining = charCap - output.Length;
            string text = (line ?? string.Empty) + Environment.NewLine;
            if (text.Length > remaining) text = text.Substring(0, remaining);
            output.Append(text);
        }

        private static void ReleaseWordCom(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return;
            try { Marshal.FinalReleaseComObject(value); } catch { }
        }
    }

    /// <summary>
    /// Read-only PowerPoint decorator. Searches text frames, table cells and speaker notes with
    /// bounded TextRange.Find calls and returns slide/shape coordinates for follow-up inspection.
    /// </summary>
    public sealed class SearchablePowerPointHostAdapter : IHostAdapter, IDocumentSearchProvider
    {
        private readonly PowerPointHostAdapter _inner;
        private readonly Ppt.Application _app;
        private readonly Func<int> _maxChars;

        public SearchablePowerPointHostAdapter(Ppt.Application app, Func<int> maxChars)
        {
            if (app == null) throw new ArgumentNullException("app");
            _app = app;
            _maxChars = maxChars ?? delegate { return 6000; };
            _inner = new PowerPointHostAdapter(app, maxChars);
        }

        public HostType Host { get { return _inner.Host; } }
        public string HostDisplayName { get { return _inner.HostDisplayName; } }
        public OfficeContext ReadContext() { return _inner.ReadContext(); }
        public string ReadSelection() { return _inner.ReadSelection(); }
        public string ReadDocument(int maxChars) { return _inner.ReadDocument(maxChars); }
        public byte[] CaptureChartAsImage(string chartName) { return _inner.CaptureChartAsImage(chartName); }
        public byte[] CaptureSlideAsImage(int slideIndexOneBased) { return _inner.CaptureSlideAsImage(slideIndexOneBased); }
        public byte[] CaptureCurrentViewAsImage() { return _inner.CaptureCurrentViewAsImage(); }
        public WritePreview PrepareWrite(string toolName, string argumentsJson) { return _inner.PrepareWrite(toolName, argumentsJson); }
        public void ApplyWrite(string toolName, string argumentsJson) { _inner.ApplyWrite(toolName, argumentsJson); }

        public string SearchDocument(string query, int maxResults, int maxChars)
        {
            query = (query ?? string.Empty).Trim();
            if (query.Length == 0) return "(empty search query)";

            var pres = _app.ActivePresentation;
            if (pres == null) return "(no presentation open)";

            int resultCap = Math.Max(1, Math.Min(20, maxResults));
            int charCap = Math.Max(256, Math.Min(maxChars, Math.Max(256, _maxChars())));
            var output = new StringBuilder();
            output.AppendLine("Presentation search: \"" + TextUtil.Truncate(query, 200) + "\"");
            int count = 0;

            for (int slideIndex = 1; slideIndex <= pres.Slides.Count && count < resultCap && output.Length < charCap; slideIndex++)
            {
                Ppt.Slide slide = pres.Slides[slideIndex];
                foreach (Ppt.Shape shape in slide.Shapes)
                {
                    if (count >= resultCap || output.Length >= charCap) break;
                    SearchPowerPointShape(query, slideIndex, shape, output, resultCap, charCap, ref count);
                }

                if (count < resultCap && output.Length < charCap)
                    SearchPowerPointNotes(query, slide, slideIndex, output, resultCap, charCap, ref count);
            }

            if (count == 0) output.AppendLine("No matches found.");
            else if (count >= resultCap) output.AppendLine("…[result limit reached: " + resultCap + "]");
            return TextUtil.Truncate(output.ToString(), charCap);
        }

        private static void SearchPowerPointShape(
            string query,
            int slideIndex,
            Ppt.Shape shape,
            StringBuilder output,
            int resultCap,
            int charCap,
            ref int count)
        {
            if (shape == null || count >= resultCap || output.Length >= charCap) return;

            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                    shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    SearchPowerPointRange(query, "Slide " + slideIndex + " / " + SafePptLabel(shape.Name, 100),
                        shape.TextFrame.TextRange, output, resultCap, charCap, ref count);
                }
            }
            catch { }

            if (count >= resultCap || output.Length >= charCap) return;
            try
            {
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    var table = shape.Table;
                    for (int row = 1; row <= table.Rows.Count && count < resultCap; row++)
                    {
                        for (int col = 1; col <= table.Columns.Count && count < resultCap; col++)
                        {
                            var cellShape = table.Cell(row, col).Shape;
                            if (cellShape != null && cellShape.HasTextFrame == Office.MsoTriState.msoTrue &&
                                cellShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                            {
                                SearchPowerPointRange(query,
                                    "Slide " + slideIndex + " / " + SafePptLabel(shape.Name, 80) +
                                    " table R" + row + "C" + col,
                                    cellShape.TextFrame.TextRange, output, resultCap, charCap, ref count);
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private static void SearchPowerPointNotes(
            string query,
            Ppt.Slide slide,
            int slideIndex,
            StringBuilder output,
            int resultCap,
            int charCap,
            ref int count)
        {
            try
            {
                var placeholders = slide.NotesPage.Shapes.Placeholders;
                for (int i = 1; i <= placeholders.Count && count < resultCap; i++)
                {
                    var shape = placeholders[i];
                    if (shape == null || shape.HasTextFrame != Office.MsoTriState.msoTrue ||
                        shape.TextFrame.HasText != Office.MsoTriState.msoTrue)
                        continue;
                    SearchPowerPointRange(query, "Slide " + slideIndex + " / speaker notes",
                        shape.TextFrame.TextRange, output, resultCap, charCap, ref count);
                }
            }
            catch { }
        }

        private static void SearchPowerPointRange(
            string query,
            string label,
            Ppt.TextRange range,
            StringBuilder output,
            int resultCap,
            int charCap,
            ref int count)
        {
            if (range == null || count >= resultCap || output.Length >= charCap) return;
            int after = 0;
            int safety = 0;
            while (count < resultCap && output.Length < charCap && safety++ < 1000)
            {
                Ppt.TextRange hit = null;
                try
                {
                    hit = range.Find(query, after, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
                    if (hit == null) break;

                    int rangeLength = 0;
                    try { rangeLength = range.Length; } catch { }
                    int hitStart = Math.Max(1, hit.Start);
                    int hitLength = Math.Max(1, hit.Length);
                    int contextStart = Math.Max(1, hitStart - 80);
                    int contextEnd = rangeLength > 0
                        ? Math.Min(rangeLength, hitStart + hitLength + 120)
                        : hitStart + hitLength;
                    int contextLength = Math.Max(1, contextEnd - contextStart + 1);

                    string context = string.Empty;
                    try { context = range.Characters(contextStart, contextLength).Text ?? string.Empty; }
                    catch
                    {
                        try { context = hit.Text ?? string.Empty; } catch { context = string.Empty; }
                    }

                    AppendPptLine(output, "• " + label + ": " +
                        TextUtil.Truncate(NormalizePptText(context), 420), charCap);
                    count++;

                    int nextAfter = hitStart + hitLength - 1;
                    if (nextAfter <= after) nextAfter = after + 1;
                    after = nextAfter;
                    if (rangeLength > 0 && after >= rangeLength) break;
                }
                finally { ReleasePptCom(hit); }
            }
        }

        private static string NormalizePptText(string value)
        {
            return (value ?? string.Empty).Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ').Trim();
        }

        private static string SafePptLabel(string value, int maxChars)
        {
            value = value ?? string.Empty;
            return value.Length <= maxChars ? value : value.Substring(0, maxChars);
        }

        private static void AppendPptLine(StringBuilder output, string line, int charCap)
        {
            if (output == null || output.Length >= charCap) return;
            int remaining = charCap - output.Length;
            string text = (line ?? string.Empty) + Environment.NewLine;
            if (text.Length > remaining) text = text.Substring(0, remaining);
            output.Append(text);
        }

        private static void ReleasePptCom(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return;
            try { Marshal.FinalReleaseComObject(value); } catch { }
        }
    }
}
