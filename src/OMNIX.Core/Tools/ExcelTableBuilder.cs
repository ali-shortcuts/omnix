using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace OMNIX.Core.Tools
{
    /// <summary>Small additive table operation; never overwrites an existing sheet or runs model code.</summary>
    public static class ExcelTableBuilder
    {
        public static JObject ValidatePlan(string json)
        {
            if (string.IsNullOrEmpty(json) || json.Length > 32000)
                throw new ArgumentException("Table plan must be nonempty and at most 32000 characters.");
            var plan = JObject.Parse(json);
            string name = (string)plan["sheet"];
            if (string.IsNullOrWhiteSpace(name) || name.Length > 31 || name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0
                || name.StartsWith("'") || name.EndsWith("'") || name.Any(char.IsControl))
                throw new ArgumentException("Use a valid new worksheet name (1–31 characters).");
            var headers = plan["headers"] as JArray;
            var rows = plan["rows"] as JArray;
            if (headers == null || headers.Count < 1 || headers.Count > 24 || rows == null || rows.Count > 50)
                throw new ArgumentException("Provide 1–24 headers and 0–50 rows.");
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in headers)
            {
                string text = h.Type == JTokenType.String ? (string)h : null;
                if (string.IsNullOrWhiteSpace(text) || text.Length > 100 || !names.Add(text.Trim()))
                    throw new ArgumentException("Headers must be distinct nonempty text, at most 100 characters.");
            }
            foreach (var token in rows)
            {
                var row = token as JArray;
                if (row == null || row.Count != headers.Count) throw new ArgumentException("Every row must match the headers.");
                foreach (var value in row)
                {
                    if (value.Type != JTokenType.Null && value.Type != JTokenType.String && value.Type != JTokenType.Integer
                        && value.Type != JTokenType.Float && value.Type != JTokenType.Boolean)
                        throw new ArgumentException("Cells must be text, numbers, booleans or null; no nested objects.");
                    if (value.ToString().Length > 500) throw new ArgumentException("Cell text exceeds 500 characters.");
                    if (value.Type == JTokenType.Integer && Math.Abs((double)value) > 999999999999999d)
                        throw new ArgumentException("Use text for identifiers or integers longer than 15 digits to preserve Excel precision.");
                    if (value.Type == JTokenType.Float && (double.IsNaN((double)value) || double.IsInfinity((double)value)))
                        throw new ArgumentException("Numbers must be finite.");
                }
            }
            return plan;
        }

        private static Excel.Workbook CheckTarget(Excel.Application app, JObject plan)
        {
            var wb = app.ActiveWorkbook;
            if (wb == null || wb.ReadOnly || wb.ProtectStructure)
                throw new InvalidOperationException("An editable workbook with unprotected structure is required.");
            foreach (object item in wb.Sheets)
            {
                // Includes chart sheets, whose names share the worksheet namespace.
                string name = item is Excel.Worksheet ? ((Excel.Worksheet)item).Name : ((Excel.Chart)item).Name;
                if (string.Equals(name, (string)plan["sheet"], StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("That sheet already exists. Choose a new name; existing data will not be overwritten.");
            }
            return wb;
        }

        public static WritePreview Prepare(Excel.Application app, string json)
        {
            var plan = ValidatePlan(json);
            CheckTarget(app, plan);
            return new WritePreview { ToolName = ToolNames.CreateDataTable,
                Title = "Create new worksheet and table: " + (string)plan["sheet"],
                Before = "Existing sheets are preserved. This adds one worksheet. Native Ctrl+Z is not guaranteed; delete the new sheet to reverse.",
                After = plan.ToString(), ArgumentsJson = json };
        }

        public static void Apply(Excel.Application app, string json)
        {
            var plan = ValidatePlan(json);
            var wb = CheckTarget(app, plan);
            var headers = (JArray)plan["headers"]; var rows = (JArray)plan["rows"];
            Excel.Worksheet created = null;
            bool events = app.EnableEvents;
            object previousSheet = app.ActiveSheet;
            try
            {
                // Avoid firing other add-ins' change handlers for each new cell.
                app.EnableEvents = false;
                created = (Excel.Worksheet)wb.Worksheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                created.Name = (string)plan["sheet"];
                var area = created.Range["A1"].Resize[Math.Max(2, rows.Count + 1), headers.Count];
                // Write each string as literal text, never as a formula supplied in data.
                for (int c = 0; c < headers.Count; c++)
                {
                    var cell = (Excel.Range)created.Cells[1, c+1];
                    cell.NumberFormat = "@"; cell.Value2 = (string)headers[c];
                }
                for (int y = 0; y < rows.Count; y++) for (int c = 0; c < headers.Count; c++)
                {
                    var cell = (Excel.Range)created.Cells[y+2, c+1]; var value = rows[y][c];
                    if (value.Type == JTokenType.String) { cell.NumberFormat = "@"; cell.Value2 = (string)value; }
                    else if (value.Type == JTokenType.Boolean) cell.Value2 = (bool)value;
                    else if (value.Type != JTokenType.Null) cell.Value2 = (double)value;
                }
                var table = created.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, area,
                    Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);
                table.TableStyle = "TableStyleMedium2";
                area.Columns.ColumnWidth = 18;
                if (table.ListColumns.Count != headers.Count || table.ListRows.Count != Math.Max(1, rows.Count))
                    throw new InvalidOperationException("Table dimensions did not match the approved plan.");
                for (int c = 0; c < headers.Count; c++)
                    if (Convert.ToString(((Excel.Range)created.Cells[1,c+1]).Value2) != (string)headers[c])
                        throw new InvalidOperationException("Table header verification failed.");
            }
            catch (Exception failure)
            {
                if (created != null)
                {
                    bool alerts = app.DisplayAlerts;
                    try { app.DisplayAlerts = false; created.Delete(); }
                    catch (Exception cleanup) { throw new InvalidOperationException("Table creation failed and its new worksheet could not be removed. Inspect the new sheet before retrying.", new AggregateException(failure, cleanup)); }
                    finally { app.DisplayAlerts = alerts; }
                }
                throw;
            }
            finally
            {
                try {
                    var ws = previousSheet as Excel.Worksheet; var chart = previousSheet as Excel.Chart;
                    if (ws != null) ws.Activate(); else if (chart != null) chart.Activate();
                } finally { app.EnableEvents = events; }
            }
        }
    }
}
