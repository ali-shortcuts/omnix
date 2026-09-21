using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace OMNIX.Core.Reference
{
    // Names and official links, not a replacement calculation engine or a capability claim.
    public static class OfficeReference
    {
        private static readonly Lazy<Dictionary<string,string>> Excel = new Lazy<Dictionary<string,string>>(() => {
            using (var stream = typeof(OfficeReference).Assembly.GetManifestResourceStream("OMNIX.ExcelFunctions.json"))
            using (var reader = new StreamReader(stream))
                return JsonConvert.DeserializeObject<Dictionary<string,string>>(reader.ReadToEnd());
        });
        public static string Search(string host, string query, int offset = 0)
        {
            query = (query ?? "").Trim();
            if (string.Equals(host, "Word", StringComparison.OrdinalIgnoreCase))
                return "Word table formulas are fields, not an Excel workbook. Microsoft lists 18 functions: ABS, AND, AVERAGE, COUNT, DEFINED, FALSE, IF, INT, MAX, MIN, MOD, NOT, OR, PRODUCT, ROUND, SIGN, SUM, TRUE. Example: =SUM(ABOVE). Update fields after editing data. Supported functions and examples: https://support.microsoft.com/en-us/word/use-a-formula-in-a-word-table\nOMNIX currently edits selected text; this reference does not enable every Word command.";
            if (string.Equals(host, "PowerPoint", StringComparison.OrdinalIgnoreCase))
                return "PowerPoint: slides, shapes, notes and visual layout. There is no Excel-style worksheet function catalog for slide text. OMNIX currently inserts slides and writes speaker notes; chart data and arbitrary ribbon commands are not covered.";
            var matches = Excel.Value.Where(p => p.Key.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(p => p.Key, StringComparer.Ordinal).ToList();
            offset = Math.Max(0, offset);
            var page = matches.Skip(offset).Take(40).Select(p => p.Key + " — " + p.Value);
            return "Excel reference: " + Excel.Value.Count + " names indexed from Microsoft on 2026-09-21. Availability depends on Excel version; this is not an installed-function count.\n" +
                "Matches: " + matches.Count + "; offset: " + offset + "; nextOffset: " + (offset + 40 < matches.Count ? (offset + 40).ToString() : "none") + "\n" +
                "Examples: =MAX(D2:D10); =SUM(G2:G10); =DSUM(A1:D20,\"Amount\",F1:F2). DSUM requires matching criteria headers.\n" + string.Join("\n", page);
        }
    }
}
