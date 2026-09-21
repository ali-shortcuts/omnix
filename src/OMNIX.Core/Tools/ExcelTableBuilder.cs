using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace OMNIX.Core.Tools
{
    /// <summary>
    /// Additive Excel worksheet/table builder.
    /// Existing sheets are never overwritten. Primitive strings are always literal text.
    /// Formula/date cells require explicit typed objects, so a model cannot accidentally turn
    /// ordinary imported text into executable Excel formulas.
    /// </summary>
    public static class ExcelTableBuilder
    {
        private const int MaxPlanChars = 32000;
        private const int MaxHeaders = 24;
        private const int MaxRows = 50;
        private const int MaxCells = 512;
        private const int MaxFormulaChars = 8192;
        private const int MaxCellTextChars = 500;

        public static JObject ValidatePlan(string json)
        {
            if (string.IsNullOrEmpty(json) || json.Length > MaxPlanChars)
                throw new ArgumentException("Table plan must be nonempty and at most " + MaxPlanChars + " characters.");

            JObject plan;
            using (var reader = new JsonTextReader(new System.IO.StringReader(json)))
            {
                reader.DateParseHandling = DateParseHandling.None;
                plan = JObject.Load(reader);
            }

            ValidateSheetName((string)plan["sheet"]);

            var headers = plan["headers"] as JArray;
            var rows = plan["rows"] as JArray;
            if (headers == null || headers.Count < 1 || headers.Count > MaxHeaders || rows == null || rows.Count > MaxRows)
                throw new ArgumentException("Provide 1–" + MaxHeaders + " headers and 0–" + MaxRows + " rows.");

            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in headers)
            {
                string text = h.Type == JTokenType.String ? (string)h : null;
                if (string.IsNullOrWhiteSpace(text) || text.Length > 100 || !names.Add(text.Trim()))
                    throw new ArgumentException("Headers must be distinct nonempty text, at most 100 characters.");
            }

            if ((rows.Count + 1) * headers.Count > MaxCells)
                throw new ArgumentException("Create at most " + MaxCells + " cells per table operation; split larger imports into stages.");

            foreach (var token in rows)
            {
                var row = token as JArray;
                if (row == null || row.Count != headers.Count)
                    throw new ArgumentException("Every row must match the headers.");
                foreach (var value in row)
                    ValidateCell(value);
            }

            return plan;
        }

        private static void ValidateCell(JToken value)
        {
            if (value == null || value.Type == JTokenType.Null) return;

            if (value.Type == JTokenType.String)
            {
                if (value.ToString().Length > MaxCellTextChars)
                    throw new ArgumentException("Cell text exceeds " + MaxCellTextChars + " characters.");
                return;
            }

            if (value.Type == JTokenType.Integer)
            {
                if (Math.Abs(Convert.ToDouble(value, CultureInfo.InvariantCulture)) > 999999999999999d)
                    throw new ArgumentException("Use text for identifiers or integers longer than 15 digits to preserve Excel precision.");
                return;
            }

            if (value.Type == JTokenType.Float)
            {
                double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (double.IsNaN(number) || double.IsInfinity(number))
                    throw new ArgumentException("Numbers must be finite.");
                return;
            }

            if (value.Type == JTokenType.Boolean) return;

            var obj = value as JObject;
            if (obj == null)
                throw new ArgumentException("Cells must be text, numbers, booleans, null, or a typed formula/date object.");

            bool hasFormula = obj["formula"] != null;
            bool hasDate = obj["date"] != null;
            if (hasFormula == hasDate)
                throw new ArgumentException("Typed cells must contain exactly one of 'formula' or 'date'.");

            if (hasFormula)
            {
                string formula = (string)obj["formula"];
                if (string.IsNullOrWhiteSpace(formula) || !formula.TrimStart().StartsWith("=", StringComparison.Ordinal))
                    throw new ArgumentException("Formula cells must begin with '='.");
                if (formula.Length > MaxFormulaChars)
                    throw new ArgumentException("Formula exceeds the Excel formula safety limit.");
            }
            else
            {
                string date = (string)obj["date"];
                DateTime parsed;
                if (!DateTime.TryParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out parsed))
                    throw new ArgumentException("Date cells must use ISO yyyy-MM-dd.");
            }

            var format = obj["numberFormat"];
            if (format != null && (format.Type != JTokenType.String || format.ToString().Length > 80))
                throw new ArgumentException("numberFormat must be short text.");
        }

        private static void ValidateSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name) || name.Length > 31 ||
                name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0 ||
                name.StartsWith("'") || name.EndsWith("'") || name.Any(char.IsControl))
                throw new ArgumentException("Use a valid new worksheet name (1–31 characters).");
        }

        private static Excel.Workbook RequireEditableWorkbook(Excel.Application app)
        {
            var wb = app.ActiveWorkbook;
            if (wb == null || wb.ReadOnly || wb.ProtectStructure)
                throw new InvalidOperationException("An editable workbook with unprotected structure is required.");
            return wb;
        }

        private static bool SheetExists(Excel.Workbook wb, string name)
        {
            foreach (object item in wb.Sheets)
            {
                string existing = item is Excel.Worksheet
                    ? ((Excel.Worksheet)item).Name
                    : ((Excel.Chart)item).Name;
                if (string.Equals(existing, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static string ResolveUniqueSheetName(Excel.Workbook wb, string preferred)
        {
            if (!SheetExists(wb, preferred)) return preferred;
            for (int n = 2; n <= 999; n++)
            {
                string suffix = " (" + n + ")";
                int baseLength = Math.Max(1, 31 - suffix.Length);
                string candidate = (preferred.Length > baseLength ? preferred.Substring(0, baseLength) : preferred) + suffix;
                if (!SheetExists(wb, candidate)) return candidate;
            }
            throw new InvalidOperationException("Could not resolve a unique worksheet name.");
        }

        public static WritePreview Prepare(Excel.Application app, string json)
        {
            var plan = ValidatePlan(json);
            var wb = RequireEditableWorkbook(app);
            string requested = (string)plan["sheet"];
            bool uniqueName = plan["uniqueName"] != null && (bool)plan["uniqueName"];
            string resolved = requested;

            if (SheetExists(wb, requested))
            {
                if (!uniqueName)
                    throw new InvalidOperationException("That sheet already exists. Existing data will not be overwritten.");
                resolved = ResolveUniqueSheetName(wb, requested);
            }

            // Bind the exact previewed name into the approved arguments. Apply must not silently
            // choose a different destination after the user has reviewed the preview.
            plan["sheet"] = resolved;
            plan["uniqueName"] = false;

            var headers = (JArray)plan["headers"];
            var rows = (JArray)plan["rows"];
            int formulas = rows.SelectMany(r => (JArray)r).Count(v => v is JObject && ((JObject)v)["formula"] != null);
            int dates = rows.SelectMany(r => (JArray)r).Count(v => v is JObject && ((JObject)v)["date"] != null);

            return new WritePreview
            {
                ToolName = ToolNames.CreateDataTable,
                Title = "Create Excel sheet: " + resolved,
                Before = "All existing worksheets remain unchanged. A new sheet will be added.",
                After = "Sheet '" + resolved + "': " + headers.Count + " columns, " + rows.Count +
                        " data rows, " + formulas + " real Excel formulas, " + dates +
                        " typed dates. A styled table and fitted columns will be created.",
                ArgumentsJson = plan.ToString(Formatting.None)
            };
        }

        public static void Apply(Excel.Application app, string json)
        {
            var plan = ValidatePlan(json);
            var wb = RequireEditableWorkbook(app);
            string sheetName = (string)plan["sheet"];
            if (SheetExists(wb, sheetName))
                throw new InvalidOperationException("The approved target sheet now exists. Nothing was overwritten; inspect the workbook and retry.");

            var headers = (JArray)plan["headers"];
            var rows = (JArray)plan["rows"];
            Excel.Worksheet created = null;
            bool completed = false;
            bool events = app.EnableEvents;
            object previousSheet = app.ActiveSheet;

            try
            {
                app.EnableEvents = false;
                created = (Excel.Worksheet)wb.Worksheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                created.Name = sheetName;

                int dataRowCount = Math.Max(1, rows.Count);
                var area = created.Range["A1"].Resize[dataRowCount + 1, headers.Count];

                for (int c = 0; c < headers.Count; c++)
                {
                    var cell = (Excel.Range)created.Cells[1, c + 1];
                    cell.NumberFormat = "@";
                    cell.Value2 = (string)headers[c];
                }

                for (int y = 0; y < rows.Count; y++)
                {
                    for (int c = 0; c < headers.Count; c++)
                    {
                        var cell = (Excel.Range)created.Cells[y + 2, c + 1];
                        WriteCell(cell, rows[y][c]);
                    }
                }

                var table = created.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    area,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing);
                table.TableStyle = "TableStyleMedium2";

                created.Range["A1"].Resize[1, headers.Count].WrapText = true;
                area.Columns.AutoFit();
                for (int c = 1; c <= headers.Count; c++)
                {
                    var column = (Excel.Range)created.Columns[c];
                    if (Convert.ToDouble(column.ColumnWidth) > 36d) column.ColumnWidth = 36d;
                    else if (Convert.ToDouble(column.ColumnWidth) < 10d) column.ColumnWidth = 10d;
                }
                created.Range["A1"].Resize[1, headers.Count].EntireRow.AutoFit();

                for (int y = 0; y < rows.Count; y++)
                {
                    for (int c = 0; c < headers.Count; c++)
                    {
                        var cell = (Excel.Range)created.Cells[y + 2, c + 1];
                        VerifyCell(cell, rows[y][c]);
                    }
                }

                if (table.ListColumns.Count != headers.Count ||
                    table.ListRows.Count != dataRowCount)
                    throw new InvalidOperationException("Table dimensions did not match the approved plan.");

                for (int c = 0; c < headers.Count; c++)
                {
                    if (Convert.ToString(((Excel.Range)created.Cells[1, c + 1]).Value2) != (string)headers[c])
                        throw new InvalidOperationException("Table header verification failed.");
                }

                completed = true;
            }
            catch (Exception failure)
            {
                if (created != null)
                {
                    bool alerts = app.DisplayAlerts;
                    try
                    {
                        app.DisplayAlerts = false;
                        created.Delete();
                    }
                    catch (Exception cleanup)
                    {
                        throw new InvalidOperationException(
                            "Sheet creation failed and its new worksheet could not be removed. Inspect the new sheet before retrying.",
                            new AggregateException(failure, cleanup));
                    }
                    finally
                    {
                        app.DisplayAlerts = alerts;
                    }
                }
                throw;
            }
            finally
            {
                try
                {
                    var ws = previousSheet as Excel.Worksheet;
                    var chart = previousSheet as Excel.Chart;
                    if (completed && created != null)
                    {
                        created.Activate();
                        created.Range["A1"].Select();
                    }
                    else if (ws != null) ws.Activate();
                    else if (chart != null) chart.Activate();
                }
                finally
                {
                    app.EnableEvents = events;
                }
            }
        }

        private static void WriteCell(Excel.Range cell, JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                cell.Value2 = null;
                return;
            }

            if (token.Type == JTokenType.String)
            {
                cell.NumberFormat = "@";
                cell.Value2 = (string)token;
                return;
            }

            if (token.Type == JTokenType.Boolean)
            {
                cell.Value2 = (bool)token;
                return;
            }

            if (token.Type == JTokenType.Integer || token.Type == JTokenType.Float)
            {
                cell.Value2 = Convert.ToDouble(token, CultureInfo.InvariantCulture);
                return;
            }

            var obj = (JObject)token;
            string numberFormat = (string)obj["numberFormat"];

            if (obj["formula"] != null)
            {
                cell.NumberFormat = string.IsNullOrWhiteSpace(numberFormat) ? "General" : numberFormat;
                cell.Formula = (string)obj["formula"];
                if (!Convert.ToBoolean(cell.HasFormula))
                    throw new InvalidOperationException("Excel did not accept an approved formula.");
                return;
            }

            DateTime date = DateTime.ParseExact((string)obj["date"], "yyyy-MM-dd", CultureInfo.InvariantCulture);
            cell.NumberFormat = string.IsNullOrWhiteSpace(numberFormat) ? "yyyy-mm-dd" : numberFormat;
            cell.Value2 = date.ToOADate();
        }

        private static void VerifyCell(Excel.Range cell, JToken expected)
        {
            object actual = cell.Value2;

            if (expected == null || expected.Type == JTokenType.Null)
            {
                if (actual != null || Convert.ToBoolean(cell.HasFormula))
                    throw new InvalidOperationException("Blank-cell verification failed.");
                return;
            }

            if (expected.Type == JTokenType.String)
            {
                if (Convert.ToString(actual) != (string)expected || Convert.ToBoolean(cell.HasFormula))
                    throw new InvalidOperationException("Text-cell verification failed.");
                return;
            }

            if (expected.Type == JTokenType.Boolean)
            {
                if (!(actual is bool) || (bool)actual != (bool)expected || Convert.ToBoolean(cell.HasFormula))
                    throw new InvalidOperationException("Boolean-cell verification failed.");
                return;
            }

            if (expected.Type == JTokenType.Integer || expected.Type == JTokenType.Float)
            {
                if (actual == null ||
                    Convert.ToDouble(actual, CultureInfo.InvariantCulture) != Convert.ToDouble(expected, CultureInfo.InvariantCulture) ||
                    Convert.ToBoolean(cell.HasFormula))
                    throw new InvalidOperationException("Numeric-cell verification failed.");
                return;
            }

            var obj = (JObject)expected;
            if (obj["formula"] != null)
            {
                if (!Convert.ToBoolean(cell.HasFormula))
                    throw new InvalidOperationException("Formula verification failed.");
                string actualFormula = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture);
                if (!string.Equals(actualFormula, (string)obj["formula"], StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Excel formula read-back did not match the approved formula.");
                return;
            }

            DateTime date = DateTime.ParseExact((string)obj["date"], "yyyy-MM-dd", CultureInfo.InvariantCulture);
            if (actual == null || Math.Abs(Convert.ToDouble(actual, CultureInfo.InvariantCulture) - date.ToOADate()) > 0.000001d)
                throw new InvalidOperationException("Date-cell verification failed.");
        }
    }
}
