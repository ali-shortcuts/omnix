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
        private static readonly Lazy<Dictionary<string, string>> Excel =
            new Lazy<Dictionary<string, string>>(() =>
            {
                using (var stream = typeof(OfficeReference).Assembly.GetManifestResourceStream("OMNIX.ExcelFunctions.json"))
                using (var reader = new StreamReader(stream))
                    return JsonConvert.DeserializeObject<Dictionary<string, string>>(reader.ReadToEnd());
            });

        public static string Search(string host, string query, int offset = 0, string language = "en")
        {
            query = (query ?? "").Trim();
            bool fa = string.Equals(language, "fa", StringComparison.OrdinalIgnoreCase);

            if (string.Equals(host, "Word", StringComparison.OrdinalIgnoreCase))
            {
                if (fa)
                    return "راهنمای Word\n" +
                           "فرمول‌های جدول در Word به‌صورت Field هستند و موتور فرمول Excel نیستند. " +
                           "Microsoft این ۱۸ تابع را برای فرمول جدول Word مستند کرده است: " +
                           "ABS, AND, AVERAGE, COUNT, DEFINED, FALSE, IF, INT, MAX, MIN, MOD, NOT, OR, PRODUCT, ROUND, SIGN, SUM, TRUE.\n" +
                           "نمونه: =SUM(ABOVE). بعد از تغییر داده‌ها، Fieldها باید Update شوند.\n" +
                           "منبع رسمی: https://support.microsoft.com/en-us/word/use-a-formula-in-a-word-table\n" +
                           "این راهنما به‌تنهایی دسترسی اجرایی تازه‌ای به Word ایجاد نمی‌کند؛ قابلیت‌های واقعی فقط از ابزارهای تأییدشده OMNIX می‌آیند.";

                return "Word table formulas are fields, not an Excel workbook. Microsoft lists 18 functions: " +
                       "ABS, AND, AVERAGE, COUNT, DEFINED, FALSE, IF, INT, MAX, MIN, MOD, NOT, OR, PRODUCT, ROUND, SIGN, SUM, TRUE. " +
                       "Example: =SUM(ABOVE). Update fields after editing data. Supported functions and examples: " +
                       "https://support.microsoft.com/en-us/word/use-a-formula-in-a-word-table\n" +
                       "OMNIX currently edits selected text; this reference does not enable every Word command.";
            }

            if (string.Equals(host, "PowerPoint", StringComparison.OrdinalIgnoreCase))
            {
                if (fa)
                    return "راهنمای PowerPoint\n" +
                           "PowerPoint شامل اسلایدها، Shapeها، Notes و چیدمان بصری است و مانند Excel کاتالوگ تابع Worksheet ندارد. " +
                           "OMNIX می‌تواند ساختار اسلاید، متن Shapeها و Notes را بخواند و از ابزارهای نوشتنی تأییدشده استفاده کند. " +
                           "این صفحهٔ راهنما به‌تنهایی دسترسی به همه فرمان‌های Ribbon نمی‌دهد.";

                return "PowerPoint: slides, shapes, notes and visual layout. There is no Excel-style worksheet function catalog for slide text. " +
                       "OMNIX currently inserts slides and writes speaker notes; chart data and arbitrary ribbon commands are not covered.";
            }

            var matches = Excel.Value
                .Where(p => p.Key.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(p => p.Key, StringComparer.Ordinal)
                .ToList();

            offset = Math.Max(0, offset);
            var page = matches.Skip(offset).Take(40).Select(p => p.Key + " — " + p.Value);
            string next = offset + 40 < matches.Count ? (offset + 40).ToString() : (fa ? "پایان" : "none");

            if (fa)
            {
                return "مرجع Excel\n" +
                       Excel.Value.Count + " نام تابع از فهرست رسمی Microsoft نمایه شده است. " +
                       "در دسترس بودن هر تابع به نسخهٔ نصب‌شدهٔ Excel بستگی دارد؛ این عدد به معنی پشتیبانی قطعی همهٔ توابع روی سیستم شما نیست.\n" +
                       "نتایج: " + matches.Count + "؛ شروع صفحه: " + offset + "؛ صفحهٔ بعد: " + next + "\n" +
                       "نمونه‌ها: =MAX(D2:D10) ؛ =SUM(G2:G10) ؛ =DSUM(A1:D20,\"Amount\",F1:F2)\n" +
                       "نام توابع همان نام رسمی انگلیسی Excel باقی می‌ماند تا فرمول‌ها دقیق باشند.\n\n" +
                       string.Join("\n", page);
            }

            return "Excel reference: " + Excel.Value.Count +
                   " names indexed from Microsoft on 2026-09-21. Availability depends on Excel version; this is not an installed-function count.\n" +
                   "Matches: " + matches.Count + "; offset: " + offset + "; nextOffset: " + next + "\n" +
                   "Examples: =MAX(D2:D10); =SUM(G2:G10); =DSUM(A1:D20,\"Amount\",F1:F2). DSUM requires matching criteria headers.\n" +
                   string.Join("\n", page);
        }
    }
}
