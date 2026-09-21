using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using OMNIX.Core.Errors;
using OMNIX.Core.Tools;
using OMNIX.Core.Util;

namespace OMNIX.Core.Context
{
    /// <summary>
    /// Excel adapter (spec Section 3, Layer 3): Workbook, Worksheet, Selection, Values, Formulas,
    /// Named Ranges, Charts (as image for Vision). All calls assume the Office UI thread.
    ///
    /// Important performance invariant: context reads are bounded BEFORE COM asks Excel for Value2
    /// or Formula arrays. Reading an entire-column / entire-sheet selection into memory and trimming
    /// it afterward can allocate millions of cells and freeze Office, so all bulk reads first resize
    /// to the configured context budget.
    /// </summary>
    public sealed class ExcelHostAdapter : IHostAdapter, IIndexedHostAdapter, IVisibleOfficeExecutionHost, IAdvancedOfficeCapabilityHost
    {
        private const int DisplayMaxColumns = 8;
        private const int FormulaCellCap = 60;
        private const int CaptureCellCap = 5000;

        private readonly Excel.Application _app;
        private readonly Func<int> _maxCells;
        private readonly Func<int> _maxChars;
        private Office.IRibbonUI _ribbonUi;

        public ExcelHostAdapter(Excel.Application app, Func<int> maxCells, Func<int> maxChars)
        {
            _app = app;
            _maxCells = maxCells;
            _maxChars = maxChars;
        }

        public HostType Host { get { return HostType.Excel; } }
        public string HostDisplayName { get { return "Excel"; } }

        public string CapabilitySummary
        {
            get
            {
                return "Excel broad Object Model capability layer: workbook/worksheet navigation; bounded values, formulas and formats; rows/columns; tables; names; sort/filter; validation; conditional formatting; charts; PivotTable refresh/inspection; comments/notes; hyperlinks; freeze panes/zoom; page setup; typed values; formulas; professional formatting; chart/current-view capture. Use list_office_capabilities for the exact current registry. Security/Trust Center, VBA/macro execution, arbitrary files/processes and unknown COM reflection are not exposed.";
            }
        }

        public void BindRibbon(Office.IRibbonUI ribbonUi) { _ribbonUi = ribbonUi; }

        public void RevealOperation(string toolName, ToolArguments args, OfficeExecutionStage stage)
        {
            ActivateRelevantRibbonTab(toolName);
            var wb = _app.ActiveWorkbook;
            if (wb == null) return;

            try
            {
                if (toolName == ToolNames.ReadDocumentSection)
                {
                    string sheetName = args.Get("sheet", "");
                    if (string.IsNullOrWhiteSpace(sheetName)) return;
                    var ws = wb.Worksheets[sheetName] as Excel.Worksheet;
                    if (ws == null) return;
                    int row = args.Integer("row", 1, 1, 1048576);
                    int col = args.Integer("column", 1, 1, 16384);
                    int rows = args.Integer("rows", 10, 1, 100);
                    int cols = args.Integer("columns", 8, 1, 32);
                    ShowRange(((Excel.Range)ws.Cells[row, col]).Resize[rows, cols]);
                    return;
                }

                if (toolName == ToolNames.WriteToCell || toolName == ToolNames.InsertFormula ||
                    toolName == ToolNames.HighlightRange || toolName == ToolNames.FormatRange)
                {
                    string sheetName = args.Get("sheet", "");
                    var ws = string.IsNullOrWhiteSpace(sheetName)
                        ? _app.ActiveSheet as Excel.Worksheet
                        : wb.Worksheets[sheetName] as Excel.Worksheet;
                    string address = args.Get("address", args.Get("range", ""));
                    if (ws != null && !string.IsNullOrWhiteSpace(address))
                        ShowRange(ws.Range[address]);
                    return;
                }

                if (toolName == ToolNames.ApplyOfficeCapability || toolName == ToolNames.InspectOfficeCapability)
                {
                    string sheetName = args.Get("sheet", "");
                    string address = args.Get("address", args.Get("destination", ""));
                    var ws = string.IsNullOrWhiteSpace(sheetName)
                        ? _app.ActiveSheet as Excel.Worksheet
                        : wb.Worksheets[sheetName] as Excel.Worksheet;
                    if (ws != null) ws.Activate();
                    if (ws != null && !string.IsNullOrWhiteSpace(address))
                    {
                        try { ShowRange(ws.Range[address]); } catch { }
                    }
                    string chartName = args.Get("chart", "");
                    if (ws != null && !string.IsNullOrWhiteSpace(chartName))
                    {
                        try
                        {
                            foreach (Excel.ChartObject chart in (Excel.ChartObjects)ws.ChartObjects())
                                if (string.Equals(chart.Name, chartName, StringComparison.OrdinalIgnoreCase)) { chart.Activate(); break; }
                        }
                        catch { }
                    }
                    return;
                }

                if (toolName == ToolNames.CreateDataTable && stage == OfficeExecutionStage.Verify)
                {
                    string sheetName = args.Get("sheet", "");
                    if (string.IsNullOrWhiteSpace(sheetName)) return;
                    var ws = wb.Worksheets[sheetName] as Excel.Worksheet;
                    if (ws == null) return;
                    ws.Activate();
                    Excel.Range target = null;
                    try
                    {
                        if (ws.ListObjects.Count > 0) target = ws.ListObjects[1].Range;
                    }
                    catch { }
                    if (target == null) target = ws.UsedRange;
                    if (target != null) ShowRange(target);
                    return;
                }

                if (toolName == ToolNames.CaptureChartAsImage)
                {
                    string chartName = args.Get("chart", "");
                    var ws = _app.ActiveSheet as Excel.Worksheet;
                    if (ws == null) return;
                    foreach (Excel.ChartObject chart in (Excel.ChartObjects)ws.ChartObjects())
                    {
                        if (string.IsNullOrWhiteSpace(chartName) ||
                            string.Equals(chart.Name, chartName, StringComparison.OrdinalIgnoreCase))
                        {
                            chart.Activate();
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("ui", "Excel visible execution target reveal failed", ex);
            }
        }

        private void ActivateRelevantRibbonTab(string toolName)
        {
            string tab = null;
            switch (toolName)
            {
                case ToolNames.InsertFormula: tab = "TabFormulas"; break;
                case ToolNames.CreateDataTable: tab = "TabInsert"; break;
                case ToolNames.ReadDocumentMap:
                case ToolNames.ReadDocumentSection: tab = "TabData"; break;
                case ToolNames.CaptureChartAsImage: tab = "TabInsert"; break;
                case ToolNames.ApplyOfficeCapability:
                case ToolNames.InspectOfficeCapability:
                case ToolNames.WriteToCell:
                case ToolNames.HighlightRange:
                case ToolNames.FormatRange:
                case ToolNames.ReadSelection:
                case ToolNames.CaptureCurrentViewAsImage: tab = "TabHome"; break;
            }
            if (tab == null || _ribbonUi == null) return;
            try { _ribbonUi.ActivateTabMso(tab); }
            catch { }
        }

        private void ShowRange(Excel.Range range)
        {
            if (range == null) return;
            try
            {
                var ws = range.Worksheet as Excel.Worksheet;
                if (ws != null) ws.Activate();
                _app.Goto(range, true);
                range.Select();
            }
            catch
            {
                try { range.Select(); } catch { }
            }
        }

        public OfficeContext ReadContext()
        {
            var ctx = new OfficeContext { Host = HostType.Excel };
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return ctx;
                ctx.DocumentName = wb.Name;
                ctx.DocumentPath = wb.FullName;

                var ws = _app.ActiveSheet as Excel.Worksheet;
                if (ws != null) ctx.ContainerName = ws.Name;

                var sel = _app.Selection as Excel.Range;
                if (sel != null)
                {
                    ctx.SelectionAddress = sel.Address[false, false];
                    ctx.SelectionText = TextUtil.Truncate(Convert.ToString(sel.Text), 200);
                    ctx.ValuesPreview = BuildValuesPreview(sel);
                    ctx.FormulasPreview = BuildFormulasPreview(sel);
                }

                ctx.NamedRanges = BuildNamedRanges(wb);
                ctx.ChartsInfo = BuildChartsInfo(ws);
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "ExcelHostAdapter.ReadContext failed", ex);
            }
            return ctx;
        }

        public string ReadSelection()
        {
            try
            {
                var sel = _app.Selection as Excel.Range;
                if (sel == null) return "(no range selected)";
                var sb = new StringBuilder();
                sb.AppendLine("Range: " + sel.Address[false, false] + " on sheet '" + ActiveSheetName() + "'");
                sb.AppendLine("Rows=" + sel.Rows.Count + " Columns=" + sel.Columns.Count);
                sb.AppendLine("Values:");
                sb.Append(BuildValuesTable(sel, _maxCells()));
                sb.AppendLine("Formulas (first cells):");
                sb.Append(BuildFormulasPreview(sel));
                return sb.ToString();
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "ExcelHostAdapter.ReadSelection failed", ex);
                return "(unable to read selection)";
            }
        }

        public string ReadDocument(int maxChars)
        {
            try
            {
                var wb = _app.ActiveWorkbook;
                if (wb == null) return "(no workbook open)";
                var sb = new StringBuilder();
                sb.AppendLine("Workbook: " + wb.Name);
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    var used = ws.UsedRange;
                    sb.AppendLine();
                    sb.AppendLine("--- Sheet '" + ws.Name + "' used range: " + used.Address[false, false] + " ---");
                    sb.Append(BuildValuesTable(used, 200));
                    if (sb.Length >= Math.Min(maxChars, _maxChars())) break;
                }
                return TextUtil.Truncate(sb.ToString(), Math.Min(maxChars, _maxChars()));
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "ExcelHostAdapter.ReadDocument failed", ex);
                return "(unable to read workbook)";
            }
        }

        public string ReadDocumentMap(int offset)
        {
            var wb = _app.ActiveWorkbook;
            if (wb == null) throw new InvalidOperationException("No workbook is open.");
            int total = wb.Worksheets.Count;
            var sb = new StringBuilder();
            int workbookNames = 0;
            try { workbookNames = wb.Names.Count; } catch { }
            sb.AppendLine("Workbook: " + wb.Name + "; worksheets=" + total + "; workbookNames=" + workbookNames + "; offset=" + offset);
            int end = Math.Min(total, offset + 20);
            for (int i = offset + 1; i <= end; i++)
            {
                var ws = (Excel.Worksheet)wb.Worksheets[i];
                int chartCount = 0, shapeCount = 0;
                try { chartCount = ((Excel.ChartObjects)ws.ChartObjects()).Count; } catch { }
                try { shapeCount = ws.Shapes.Count; } catch { }
                sb.AppendLine("Sheet=" + ws.Name + "; used=" + ws.UsedRange.Address[false, false]
                    + "; visibility=" + ws.Visible + "; tables=" + ws.ListObjects.Count
                    + "; charts=" + chartCount + "; shapes=" + shapeCount);
            }
            sb.AppendLine("nextOffset=" + (end < total ? end.ToString() : "none"));
            sb.AppendLine("Map only: no cell contents read. Table/chart/shape counts come directly from the workbook object model. VBA, connections and external-file contents are not executed or imported.");
            return sb.ToString();
        }

        public string ReadDocumentSection(ToolArguments args)
        {
            var wb = _app.ActiveWorkbook;
            if (wb == null) throw new InvalidOperationException("No workbook is open.");
            string name = args.Get("sheet", "");
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Specify a sheet name from read_document_map.");
            var ws = (Excel.Worksheet)wb.Worksheets[name];
            int row = args.Integer("row", 1, 1, 1048576);
            int col = args.Integer("column", 1, 1, 16384);
            int rows = args.Integer("rows", 10, 1, 100);
            int cols = args.Integer("columns", 8, 1, 32);
            if (rows * cols > 256 || row + rows - 1 > 1048576 || col + cols - 1 > 16384)
                throw new ArgumentException("Request at most 256 cells inside worksheet boundaries.");
            // Resize BEFORE asking COM for arrays. No selection or active-sheet mutation.
            var range = ((Excel.Range)ws.Cells[row, col]).Resize[rows, cols];
            object values = range.Value2, formulas = range.Formula, formats = range.NumberFormat;
            var va = values as Array; var fa = formulas as Array; var nfa = formats as Array;
            var sb = new StringBuilder();
            sb.AppendLine("Sheet=" + ws.Name + "; requested=" + range.Address[false, false]);
            int shown = 0;
            for (int y = 1; y <= rows; y++)
            {
                for (int x = 1; x <= cols; x++)
                {
                    string v = Convert.ToString(va == null ? values : va.GetValue(y, x));
                    string f = Convert.ToString(fa == null ? formulas : fa.GetValue(y, x));
                    string nf = Convert.ToString(nfa == null ? formats : nfa.GetValue(y, x));
                    string cell = "row=" + (row+y-1) + ",column=" + (col+x-1)
                        + "; value=" + Newtonsoft.Json.JsonConvert.SerializeObject(TextUtil.Truncate(v, 600))
                        + "; formula=" + Newtonsoft.Json.JsonConvert.SerializeObject(TextUtil.Truncate(f, 600))
                        + "; numberFormat=" + Newtonsoft.Json.JsonConvert.SerializeObject(TextUtil.Truncate(nf, 120));
                    if (sb.Length + cell.Length > 5500)
                    {
                        sb.AppendLine("PARTIAL: next unread row=" + (row+y-1) + ",column=" + (col+x-1) + "; request a smaller region.");
                        return sb.ToString();
                    }
                    sb.AppendLine(cell); shown++;
                }
            }
            sb.AppendLine("Cells returned=" + shown + "; each value/formula capped at 600 characters; other cells were NOT read.");
            return sb.ToString();
        }

        public byte[] CaptureChartAsImage(string chartName)
        {
            try
            {
                var ws = _app.ActiveSheet as Excel.Worksheet;
                if (ws == null) return null;
                Excel.Chart chart = null;
                if (!string.IsNullOrEmpty(chartName))
                {
                    foreach (Excel.ChartObject co in (Excel.ChartObjects)ws.ChartObjects())
                    {
                        if (string.Equals(co.Name, chartName, StringComparison.OrdinalIgnoreCase))
                        {
                            chart = co.Chart;
                            break;
                        }
                    }
                }
                if (chart == null)
                {
                    try { chart = _app.ActiveChart; } catch { }
                }
                if (chart == null) return null;

                return TempImageCapture.FromExporter(path => chart.Export(path, "PNG", true));
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "ExcelHostAdapter.CaptureChartAsImage failed", ex);
                return null;
            }
        }

        public byte[] CaptureSlideAsImage(int slideIndexOneBased)
        {
            return null;
        }

        public byte[] CaptureCurrentViewAsImage()
        {
            try
            {
                var chartImg = CaptureChartAsImage(null);
                if (chartImg != null && chartImg.Length > 0) return chartImg;

                var sel = _app.Selection as Excel.Range;
                if (sel == null) return null;

                Excel.Range captureRange = FirstArea(sel);
                try
                {
                    if (Convert.ToDouble(captureRange.Cells.CountLarge) > CaptureCellCap)
                    {
                        var win = _app.ActiveWindow;
                        var visible = win != null ? win.VisibleRange : null;
                        if (visible != null)
                        {
                            captureRange = FirstArea(visible);
                            Logging.Logger.Gateway("Excel Vision capture: huge selection replaced with bounded current visible range.");
                        }
                    }
                }
                catch { }

                int totalRows, totalCols, shownRows, shownCols, areaCount;
                captureRange = CreateBoundedReadRange(
                    captureRange, CaptureCellCap, 80,
                    out totalRows, out totalCols, out shownRows, out shownCols, out areaCount);

                return TempImageCapture.FromExporter(path =>
                {
                    captureRange.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);
                    var wb = _app.ActiveWorkbook;
                    var originalSheet = _app.ActiveSheet;
                    bool events = _app.EnableEvents;
                    bool alerts = _app.DisplayAlerts;
                    Excel.Chart tempChart = null;
                    try
                    {
                        _app.EnableEvents = false;
                        tempChart = (Excel.Chart)wb.Charts.Add();
                        tempChart.Paste();
                        tempChart.Export(path, "PNG", true);
                    }
                    finally
                    {
                        try
                        {
                            // Only our own temporary chart is deleted; never a user worksheet.
                            _app.DisplayAlerts = false;
                            if (tempChart != null) tempChart.Delete();
                        }
                        finally
                        {
                            try {
                                var sheet = originalSheet as Excel.Worksheet;
                                if (sheet != null) sheet.Activate();
                                var chart = originalSheet as Excel.Chart;
                                if (chart != null) chart.Activate();
                                sel.Select();
                            }
                            finally { _app.DisplayAlerts = alerts; _app.EnableEvents = events; }
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "ExcelHostAdapter.CaptureCurrentViewAsImage failed", ex);
                return null;
            }
        }

        public WritePreview PrepareWrite(string toolName, string argumentsJson)
        {
            switch (toolName)
            {
                case ToolNames.CreateDataTable:
                    return ExcelTableBuilder.Prepare(_app, argumentsJson);
                case ToolNames.WriteToCell:
                case ToolNames.InsertFormula:
                case ToolNames.HighlightRange:
                case ToolNames.FormatRange:
                    return ExcelWrite.Prepare(this, toolName, argumentsJson);
                case ToolNames.ApplyOfficeCapability:
                    return PrepareCapabilityWrite(argumentsJson);
                default:
                    throw new OmnixException(ErrorCode.CORE_ERROR,
                        "Tool '" + toolName + "' is not supported by Excel.",
                        "ExcelHostAdapter.PrepareWrite", "Use a supported tool.");
            }
        }

        public void ApplyWrite(string toolName, string argumentsJson)
        {
            if (toolName == ToolNames.CreateDataTable) ExcelTableBuilder.Apply(_app, argumentsJson);
            else if (toolName == ToolNames.ApplyOfficeCapability) ApplyCapabilityWrite(argumentsJson);
            else ExcelWrite.ApplyWrite(this, toolName, argumentsJson);
        }

        public string InspectCapability(ToolArguments arguments)
        {
            return ExcelAdvancedCapabilities.Inspect(this, arguments);
        }

        public WritePreview PrepareCapabilityWrite(string argumentsJson)
        {
            return ExcelAdvancedCapabilities.Prepare(this, argumentsJson);
        }

        public void ApplyCapabilityWrite(string argumentsJson)
        {
            ExcelAdvancedCapabilities.Apply(this, argumentsJson);
        }

        internal Excel.Application App { get { return _app; } }
        internal int MaxCells { get { return _maxCells(); } }

        internal string ActiveSheetName()
        {
            var ws = _app.ActiveSheet as Excel.Worksheet;
            return ws != null ? ws.Name : "?";
        }

        private string BuildValuesPreview(Excel.Range sel)
        {
            int cap = Math.Max(1, _maxCells());
            var table = BuildValuesTable(sel, cap);
            return TextUtil.Truncate(table, _maxChars());
        }

        private string BuildFormulasPreview(Excel.Range sel)
        {
            var sb = new StringBuilder();
            int cap = Math.Max(1, Math.Min(FormulaCellCap, _maxCells()));
            try
            {
                int totalRows, totalCols, shownRows, shownCols, areaCount;
                Excel.Range part = CreateBoundedReadRange(
                    sel, cap, DisplayMaxColumns,
                    out totalRows, out totalCols, out shownRows, out shownCols, out areaCount);

                AppendFormulas(part, shownRows, shownCols, sb);
                if (areaCount > 1)
                    sb.AppendLine("…[multi-area selection: showing first area only]");
                if (shownRows < totalRows || shownCols < totalCols)
                    sb.AppendLine("…[formulas truncated: shown " + shownRows + "×" + shownCols +
                                  " of " + totalRows + "×" + totalCols + "]");
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "BuildFormulasPreview failed", ex);
            }
            return TextUtil.Truncate(sb.ToString(), 1200);
        }

        private static void AppendFormulas(Excel.Range range, int rows, int cols, StringBuilder sb)
        {
            object raw = range.Formula;
            if (rows == 1 && cols == 1)
            {
                sb.AppendLine(raw == null ? "" : Convert.ToString(raw));
                return;
            }

            var formulas = raw as object[,];
            if (formulas == null) return;
            for (int r = 1; r <= rows; r++)
            {
                var line = new List<string>();
                for (int c = 1; c <= cols; c++)
                {
                    object v = formulas[r, c];
                    line.Add(v == null ? "" : Convert.ToString(v));
                }
                sb.AppendLine(string.Join(" | ", line));
            }
        }

        internal string BuildValuesTable(Excel.Range range, int cellCap)
        {
            var sb = new StringBuilder();
            try
            {
                int totalRows, totalCols, shownRows, shownCols, areaCount;
                Excel.Range part = CreateBoundedReadRange(
                    range, Math.Max(1, cellCap), DisplayMaxColumns,
                    out totalRows, out totalCols, out shownRows, out shownCols, out areaCount);

                object raw = part.Value2;
                if (shownRows == 1 && shownCols == 1)
                {
                    sb.AppendLine(part.Address[false, false] + " = " + SafeText(raw));
                }
                else
                {
                    var vals = raw as object[,];
                    if (vals != null)
                    {
                        for (int r = 1; r <= shownRows; r++)
                        {
                            var line = new List<string>();
                            for (int c = 1; c <= shownCols; c++)
                                line.Add(SafeText(vals[r, c]));
                            sb.AppendLine(string.Join(" | ", line));
                        }
                    }
                }

                if (areaCount > 1)
                    sb.AppendLine("…[multi-area selection: showing first area only]");
                if (shownRows < totalRows || shownCols < totalCols)
                    sb.AppendLine("…[" + totalRows + " rows × " + totalCols +
                                  " cols in first area — shown " + shownRows + "×" + shownCols + "]");
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "BuildValuesTable failed", ex);
            }
            return sb.ToString();
        }

        private static Excel.Range CreateBoundedReadRange(
            Excel.Range range,
            int cellCap,
            int maxColumns,
            out int totalRows,
            out int totalCols,
            out int shownRows,
            out int shownCols,
            out int areaCount)
        {
            if (range == null) throw new ArgumentNullException("range");

            Excel.Range source = FirstArea(range);
            areaCount = 1;
            try { areaCount = Math.Max(1, range.Areas.Count); } catch { }

            totalRows = Math.Max(1, source.Rows.Count);
            totalCols = Math.Max(1, source.Columns.Count);
            int safeCap = Math.Max(1, cellCap);
            int safeMaxColumns = Math.Max(1, maxColumns);

            shownCols = Math.Min(totalCols, Math.Min(safeMaxColumns, safeCap));
            shownRows = Math.Min(totalRows, Math.Max(1, safeCap / Math.Max(1, shownCols)));

            if (shownRows == totalRows && shownCols == totalCols)
                return source;
            return source.Resize[shownRows, shownCols];
        }

        private static Excel.Range FirstArea(Excel.Range range)
        {
            if (range == null) return null;
            try
            {
                if (range.Areas != null && range.Areas.Count > 1)
                    return range.Areas[1];
            }
            catch { }
            return range;
        }

        private static string SafeText(object v)
        {
            if (v == null) return "";
            return Convert.ToString(v) ?? "";
        }

        private string BuildNamedRanges(Excel.Workbook wb)
        {
            var sb = new StringBuilder();
            try
            {
                int i = 0;
                foreach (Excel.Name n in wb.Names)
                {
                    if (i++ >= 50) { sb.AppendLine("…"); break; }
                    sb.AppendLine(n.Name + " = " + n.RefersTo);
                }
            }
            catch { }
            return sb.Length == 0 ? "(none)" : sb.ToString();
        }

        private string BuildChartsInfo(Excel.Worksheet ws)
        {
            if (ws == null) return "(none)";
            var names = new List<string>();
            try
            {
                foreach (Excel.ChartObject co in (Excel.ChartObjects)ws.ChartObjects())
                    names.Add(co.Name);
            }
            catch { }
            return names.Count == 0 ? "(no charts)" : string.Join(", ", names);
        }
    }

    /// <summary>
    /// Excel write-tool plumbing. Every mutation is bounded before the preview is shown and is
    /// validated again immediately before ApplyWrite, so a model cannot turn a single-cell request
    /// into a whole-sheet mutation by changing arguments between stages.
    /// </summary>
    internal static class ExcelWrite
    {
        private const int ExcelCellTextLimit = 32767;
        private const int ExcelFormulaLengthLimit = 8192;
        private const int MaxHighlightCells = 10000;
        private const int MaxFormatCells = 10000;
        private const int MaxAddressChars = 128;

        public static WritePreview Prepare(ExcelHostAdapter adapter, string toolName, string argumentsJson)
        {
            var args = ToolArguments.Parse(argumentsJson);
            string address = NormalizeAddress(args.Get("address", args.Get("range", "")));
            var ws = RequireWorksheet(adapter, toolName, args.Get("sheet", ""));
            Excel.Range target = ResolveAndValidateTarget(adapter, ws, toolName, address, args);

            string before;
            if (toolName == ToolNames.HighlightRange)
            {
                string old;
                try { old = Convert.ToString(target.Interior.ColorIndex); }
                catch { old = "(mixed/unknown)"; }
                before = "Target " + target.Address[false, false] + " current Interior.ColorIndex: " + old;
            }
            else if (toolName == ToolNames.FormatRange)
            {
                string numberFormat = "(mixed/unknown)";
                string fontName = "(mixed/unknown)";
                try { numberFormat = Convert.ToString(target.NumberFormat); } catch { }
                try { fontName = Convert.ToString(target.Font.Name); } catch { }
                before = "Target " + target.Address[false, false] + " (" + GetCellCount(target) +
                         " cells); current font=" + fontName + "; numberFormat=" + numberFormat;
            }
            else
            {
                object current = toolName == ToolNames.InsertFormula ? target.Formula : target.Value2;
                before = target.Address[false, false] + " currently = " +
                         (current == null ? "(empty)" : TextUtil.Truncate(Convert.ToString(current), 1000));
            }

            string after;
            if (toolName == ToolNames.WriteToCell)
            {
                var valueToken = args.Token("value");
                string previewValue = valueToken == null ? "" : valueToken.ToString(Newtonsoft.Json.Formatting.None);
                after = target.Address[false, false] + " will contain: " + TextUtil.Truncate(previewValue, 2000);
            }
            else if (toolName == ToolNames.InsertFormula)
                after = target.Address[false, false] + " formula will be: " + TextUtil.Truncate(args.Get("formula", args.Get("value", "")), 2000);
            else if (toolName == ToolNames.FormatRange)
                after = target.Address[false, false] + " (" + GetCellCount(target) + " cells) formatting: " + DescribeFormatting(args);
            else
                after = target.Address[false, false] + " (" + GetCellCount(target) + " cells) will be highlighted yellow.";

            return new WritePreview
            {
                ToolName = toolName,
                Title = "Excel — " + ws.Name + " — " + toolName,
                Before = before,
                After = after,
                ArgumentsJson = argumentsJson
            };
        }

        public static void ApplyWrite(ExcelHostAdapter adapter, string toolName, string argumentsJson)
        {
            var args = ToolArguments.Parse(argumentsJson);
            string address = NormalizeAddress(args.Get("address", args.Get("range", "")));
            var ws = RequireWorksheet(adapter, toolName, args.Get("sheet", ""));
            Excel.Range target = ResolveAndValidateTarget(adapter, ws, toolName, address, args);

            switch (toolName)
            {
                case ToolNames.WriteToCell:
                    ApplyTypedCellValue(target, args.Token("value"));
                    break;
                case ToolNames.InsertFormula:
                    target.NumberFormat = "General";
                    target.Formula = args.Get("formula", args.Get("value", ""));
                    if (!Convert.ToBoolean(target.HasFormula))
                        throw new InvalidOperationException("Excel did not accept this cell as a formula. Inspect the target cell before retrying.");
                    break;
                case ToolNames.HighlightRange:
                    target.Interior.Color = 0x3BEBFF;
                    break;
                case ToolNames.FormatRange:
                    ApplyFormatting(target, args);
                    break;
                default:
                    throw new OmnixException(ErrorCode.CORE_ERROR, "Unknown Excel write tool: " + toolName, "", "");
            }
            Logging.Logger.Install("Excel write tool applied: " + toolName + " -> " + target.Address[false, false]);
        }

        private static Excel.Worksheet RequireWorksheet(ExcelHostAdapter adapter, string toolName, string sheet)
        {
            var ws = adapter.App.ActiveSheet as Excel.Worksheet;
            if (!string.IsNullOrWhiteSpace(sheet))
            {
                var book = adapter.App.ActiveWorkbook;
                if (book == null) throw new InvalidOperationException("No workbook is active.");
                ws = book.Worksheets[sheet] as Excel.Worksheet;
            }
            if (ws == null)
                throw new OmnixException(ErrorCode.CORE_ERROR, "No worksheet is active.", toolName, "Open a worksheet.");
            return ws;
        }

        private static string NormalizeAddress(string address)
        {
            address = (address ?? "").Trim();
            if (address.Length == 0)
                throw new OmnixException(ErrorCode.CORE_ERROR, "Missing target address.",
                    "Excel write requires 'address'.", "Provide an A1-style address like A1 or A1:C5.");
            if (address.Length > MaxAddressChars)
                throw new OmnixException(ErrorCode.CORE_ERROR, "Target address is too long.",
                    "Excel write address exceeded " + MaxAddressChars + " characters.", "Use a bounded A1-style range on the active sheet.");
            return address;
        }

        private static Excel.Range ResolveAndValidateTarget(
            ExcelHostAdapter adapter,
            Excel.Worksheet ws,
            string toolName,
            string address,
            ToolArguments args)
        {
            Excel.Range target;
            try { target = ws.Range[address]; }
            catch (Exception ex)
            {
                throw new OmnixException(ErrorCode.CORE_ERROR, "Invalid Excel target address.",
                    toolName + " address='" + address + "'. " + ex.Message,
                    "Use an A1-style address on the active worksheet.");
            }

            if (target == null)
                throw new OmnixException(ErrorCode.CORE_ERROR, "Excel target could not be resolved.", toolName, "Use a valid A1-style address.");

            int areaCount = 1;
            try { areaCount = target.Areas.Count; } catch { }
            if (areaCount != 1)
                throw new OmnixException(ErrorCode.CORE_ERROR, "Multi-area Excel writes are not allowed.",
                    toolName + " resolved to " + areaCount + " areas.", "Use one contiguous range.");

            long cells = GetCellCount(target);
            if (toolName == ToolNames.WriteToCell || toolName == ToolNames.InsertFormula)
            {
                if (cells != 1)
                    throw new OmnixException(ErrorCode.CORE_ERROR,
                        "This write tool requires exactly one target cell.",
                        toolName + " resolved to " + cells + " cells.",
                        "Use write_to_cell/insert_formula for one cell at a time, or ask for a smaller explicit change.");
            }
            else if (toolName == ToolNames.HighlightRange || toolName == ToolNames.FormatRange)
            {
                int configuredCap = Math.Max(1, adapter.MaxCells);
                long hardCap = toolName == ToolNames.FormatRange ? MaxFormatCells : MaxHighlightCells;
                long cap = Math.Min(hardCap, configuredCap);
                if (cells > cap)
                    throw new OmnixException(ErrorCode.CORE_ERROR,
                        (toolName == ToolNames.FormatRange ? "Format" : "Highlight") + " range is too large for one AI-approved mutation.",
                        "Requested " + cells + " cells; limit=" + cap + ".",
                        "Use a smaller contiguous range and approve it separately.");
                if (toolName == ToolNames.FormatRange) ValidateFormattingArgs(args);
            }

            if (toolName == ToolNames.WriteToCell)
            {
                var valueToken = args.Token("value");
                if (valueToken != null && valueToken.Type == Newtonsoft.Json.Linq.JTokenType.String)
                {
                    string value = valueToken.ToString();
                    if (value.Length > ExcelCellTextLimit)
                        throw new OmnixException(ErrorCode.CORE_ERROR,
                            "Cell value is too large for Excel.",
                            "write_to_cell length=" + value.Length + "; Excel limit=" + ExcelCellTextLimit + ".",
                            "Shorten the value or split it across cells deliberately.");
                }
                else if (valueToken != null &&
                         valueToken.Type != Newtonsoft.Json.Linq.JTokenType.Integer &&
                         valueToken.Type != Newtonsoft.Json.Linq.JTokenType.Float &&
                         valueToken.Type != Newtonsoft.Json.Linq.JTokenType.Boolean &&
                         valueToken.Type != Newtonsoft.Json.Linq.JTokenType.Null)
                {
                    throw new OmnixException(ErrorCode.CORE_ERROR,
                        "write_to_cell accepts only text, number, boolean or null.",
                        "Unsupported JSON value type=" + valueToken.Type,
                        "Use create_data_table for structured multi-cell data.");
                }
            }
            else if (toolName == ToolNames.InsertFormula)
            {
                string formula = args.Get("formula", args.Get("value", "")) ?? "";
                if (string.IsNullOrWhiteSpace(formula) || !formula.TrimStart().StartsWith("=", StringComparison.Ordinal))
                    throw new OmnixException(ErrorCode.CORE_ERROR,
                        "Formula must begin with '='.", toolName, "Provide a valid Excel formula such as =SUM(A1:A10).");
                if (formula.Length > ExcelFormulaLengthLimit)
                    throw new OmnixException(ErrorCode.CORE_ERROR,
                        "Formula is too long for a bounded AI write.",
                        "insert_formula length=" + formula.Length + "; limit=" + ExcelFormulaLengthLimit + ".",
                        "Simplify the formula or break the task into smaller explicit steps.");
            }

            return target;
        }

        private static void ApplyTypedCellValue(Excel.Range target, Newtonsoft.Json.Linq.JToken token)
        {
            if (token == null || token.Type == Newtonsoft.Json.Linq.JTokenType.Null)
            {
                target.Value2 = null;
                return;
            }

            if (token.Type == Newtonsoft.Json.Linq.JTokenType.String)
            {
                target.NumberFormat = "@";
                target.Value2 = token.ToString();
                return;
            }

            if (token.Type == Newtonsoft.Json.Linq.JTokenType.Boolean)
            {
                target.Value2 = (bool)token;
                return;
            }

            if (token.Type == Newtonsoft.Json.Linq.JTokenType.Integer ||
                token.Type == Newtonsoft.Json.Linq.JTokenType.Float)
            {
                double value = Convert.ToDouble(token, System.Globalization.CultureInfo.InvariantCulture);
                if (double.IsNaN(value) || double.IsInfinity(value))
                    throw new InvalidOperationException("Numeric cell values must be finite.");
                target.Value2 = value;
                return;
            }

            throw new InvalidOperationException("Unsupported cell value type.");
        }

        private static string DescribeFormatting(ToolArguments args)
        {
            var items = new List<string>();
            AddIf(items, "font", args.Get("fontName", ""));
            AddIf(items, "fontSize", args.Get("fontSize", ""));
            AddIf(items, "bold", args.Get("bold", ""));
            AddIf(items, "italic", args.Get("italic", ""));
            AddIf(items, "underline", args.Get("underline", ""));
            AddIf(items, "fontColor", args.Get("fontColor", ""));
            AddIf(items, "fillColor", args.Get("fillColor", ""));
            AddIf(items, "horizontal", args.Get("horizontalAlignment", ""));
            AddIf(items, "vertical", args.Get("verticalAlignment", ""));
            AddIf(items, "numberFormat", args.Get("numberFormat", ""));
            AddIf(items, "wrapText", args.Get("wrapText", ""));
            AddIf(items, "border", args.Get("border", ""));
            AddIf(items, "autofitColumns", args.Get("autofitColumns", ""));
            AddIf(items, "autofitRows", args.Get("autofitRows", ""));
            return items.Count == 0 ? "(no formatting properties supplied)" : string.Join(", ", items);
        }

        private static void AddIf(List<string> items, string label, string value)
        {
            if (!string.IsNullOrWhiteSpace(value)) items.Add(label + "=" + value);
        }

        private static void ValidateFormattingArgs(ToolArguments args)
        {
            string fontName = args.Get("fontName", "");
            if (fontName.Length > 80)
                throw new OmnixException(ErrorCode.CORE_ERROR, "Font name is too long.", "format_range fontName", "Use a normal installed font name.");

            string fontSize = args.Get("fontSize", "");
            if (!string.IsNullOrWhiteSpace(fontSize))
            {
                double size;
                if (!double.TryParse(fontSize, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out size) || size < 6 || size > 72)
                    throw new OmnixException(ErrorCode.CORE_ERROR, "Font size must be between 6 and 72.", "format_range fontSize", "Use a normal Office font size.");
            }

            ValidateOptionalBool(args, "bold");
            ValidateOptionalBool(args, "italic");
            ValidateOptionalBool(args, "underline");
            ValidateOptionalBool(args, "wrapText");
            ValidateOptionalBool(args, "autofitColumns");
            ValidateOptionalBool(args, "autofitRows");

            string numberFormat = args.Get("numberFormat", "");
            if (numberFormat.Length > 100)
                throw new OmnixException(ErrorCode.CORE_ERROR, "Number format is too long.", "format_range numberFormat", "Use a concise Excel number format.");

            ValidateAlignment(args.Get("horizontalAlignment", ""), true);
            ValidateAlignment(args.Get("verticalAlignment", ""), false);
            ValidateColor(args.Get("fontColor", ""));
            ValidateColor(args.Get("fillColor", ""));

            string border = args.Get("border", "");
            if (!string.IsNullOrWhiteSpace(border) &&
                !string.Equals(border, "none", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(border, "thin", StringComparison.OrdinalIgnoreCase))
                throw new OmnixException(ErrorCode.CORE_ERROR, "Border must be 'none' or 'thin'.", "format_range border", "Use a supported professional border style.");
        }

        private static void ValidateOptionalBool(ToolArguments args, string key)
        {
            string raw = args.Get(key, "");
            bool parsed;
            if (!string.IsNullOrWhiteSpace(raw) && !bool.TryParse(raw, out parsed))
                throw new OmnixException(ErrorCode.CORE_ERROR, key + " must be true or false.", "format_range " + key, "Use a JSON boolean.");
        }

        private static void ValidateAlignment(string raw, bool horizontal)
        {
            if (string.IsNullOrWhiteSpace(raw)) return;
            string value = raw.Trim().ToLowerInvariant();
            string[] allowed = horizontal
                ? new[] { "general", "left", "center", "right" }
                : new[] { "top", "center", "bottom" };
            if (!allowed.Contains(value))
                throw new OmnixException(ErrorCode.CORE_ERROR,
                    (horizontal ? "Horizontal" : "Vertical") + " alignment is unsupported.",
                    "format_range alignment=" + raw,
                    "Use " + string.Join(", ", allowed) + ".");
        }

        private static void ValidateColor(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return;
            string value = raw.Trim();
            if (value.Length != 7 || value[0] != '#' ||
                !value.Substring(1).All(ch => Uri.IsHexDigit(ch)))
                throw new OmnixException(ErrorCode.CORE_ERROR, "Colors must use #RRGGBB.", "format_range color=" + raw, "Use a six-digit RGB color.");
        }

        private static void ApplyFormatting(Excel.Range target, ToolArguments args)
        {
            ValidateFormattingArgs(args);

            string fontName = args.Get("fontName", "");
            if (!string.IsNullOrWhiteSpace(fontName)) target.Font.Name = fontName;

            string fontSize = args.Get("fontSize", "");
            if (!string.IsNullOrWhiteSpace(fontSize))
                target.Font.Size = double.Parse(fontSize, System.Globalization.CultureInfo.InvariantCulture);

            bool value;
            if (bool.TryParse(args.Get("bold", ""), out value)) target.Font.Bold = value;
            if (bool.TryParse(args.Get("italic", ""), out value)) target.Font.Italic = value;
            if (bool.TryParse(args.Get("underline", ""), out value))
                target.Font.Underline = value ? Excel.XlUnderlineStyle.xlUnderlineStyleSingle : Excel.XlUnderlineStyle.xlUnderlineStyleNone;
            if (bool.TryParse(args.Get("wrapText", ""), out value)) target.WrapText = value;

            string fontColor = args.Get("fontColor", "");
            if (!string.IsNullOrWhiteSpace(fontColor))
                target.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml(fontColor));
            string fillColor = args.Get("fillColor", "");
            if (!string.IsNullOrWhiteSpace(fillColor))
                target.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml(fillColor));

            string numberFormat = args.Get("numberFormat", "");
            if (!string.IsNullOrWhiteSpace(numberFormat)) target.NumberFormat = numberFormat;

            string h = args.Get("horizontalAlignment", "").Trim().ToLowerInvariant();
            if (h == "general") target.HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral;
            else if (h == "left") target.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            else if (h == "center") target.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            else if (h == "right") target.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            string v = args.Get("verticalAlignment", "").Trim().ToLowerInvariant();
            if (v == "top") target.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            else if (v == "center") target.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            else if (v == "bottom") target.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;

            string border = args.Get("border", "").Trim().ToLowerInvariant();
            if (border == "none")
                target.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            else if (border == "thin")
            {
                target.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                target.Borders.Weight = Excel.XlBorderWeight.xlThin;
            }

            if (bool.TryParse(args.Get("autofitColumns", ""), out value) && value) target.Columns.AutoFit();
            if (bool.TryParse(args.Get("autofitRows", ""), out value) && value) target.Rows.AutoFit();
        }

        private static long GetCellCount(Excel.Range target)
        {
            try { return Convert.ToInt64(target.Cells.CountLarge); }
            catch { return Convert.ToInt64(target.Cells.Count); }
        }
    }
}
