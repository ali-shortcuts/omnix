using System;
using OMNIX.Core.Context;

namespace OMNIX.Core.AiGateway
{
    /// <summary>
    /// Conservative multilingual detector used only to decide whether a text-only final answer
    /// should be accepted for a request that appears to require a real Office mutation.
    /// It does not authorize writes; ToolExecutor confirmation remains mandatory.
    /// </summary>
    public static class MutationIntentDetector
    {
        private static readonly string[] StrongMutationTerms =
        {
            "build", "create", "make", "insert", "write", "update", "edit", "delete", "remove",
            "format", "rename", "add", "fill", "populate", "generate", "replace", "set", "apply",
            "database", "worksheet", "sheet", "table", "formula", "slide", "document",
            "بساز", "ساخته", "ایجاد", "اضافه", "وارد کن", "بنویس", "بنويس", "ویرایش", "ويرايش",
            "تغییر", "تغيير", "حذف", "پاک کن", "فرمت", "قالب", "جدول", "دیتابیس", "ديتابيس",
            "شیت", "شيت", "فرمول", "اسلاید", "اسلايد", "در همین فایل", "در همين فايل"
        };

        private static readonly string[] AdviceOnlyTerms =
        {
            "how to", "how do i", "explain", "guide me", "tutorial", "what is",
            "چگونه", "چطور", "راهنما", "توضیح بده", "توضيح بده", "روش ساخت", "آموزش"
        };

        private static readonly string[] ExplicitExecutionTerms =
        {
            "in this file", "in this workbook", "in this document", "in this presentation",
            "using the real tools", "use the tools", "actually create", "actually build",
            "در همین فایل", "در همين فايل", "داخل اکسل", "داخل ورد", "داخل پاورپوینت",
            "با ابزار واقعی", "با ابزارها", "واقعاً بساز", "واقعا بساز", "در فایل ایجاد کن"
        };

        public static bool LikelyMutation(string text, IHostAdapter host)
        {
            if (host == null || string.IsNullOrWhiteSpace(text)) return false;
            string n = Normalize(text);

            bool strong = ContainsAny(n, StrongMutationTerms);
            if (!strong) return false;

            bool explicitExecution = ContainsAny(n, ExplicitExecutionTerms);
            bool adviceOnly = ContainsAny(n, AdviceOnlyTerms);
            if (adviceOnly && !explicitExecution) return false;

            // Host-specific nouns raise confidence for terse commands such as "add a slide".
            if (host.Host == HostType.Excel &&
                (n.Contains("excel") || n.Contains("اکسل") || n.Contains("sheet") || n.Contains("شیت") ||
                 n.Contains("table") || n.Contains("جدول") || n.Contains("formula") || n.Contains("فرمول")))
                return true;

            if (host.Host == HostType.Word &&
                (n.Contains("word") || n.Contains("ورد") || n.Contains("document") || n.Contains("سند") ||
                 n.Contains("paragraph") || n.Contains("پاراگراف") || n.Contains("table") || n.Contains("جدول")))
                return true;

            if (host.Host == HostType.PowerPoint &&
                (n.Contains("powerpoint") || n.Contains("پاورپوینت") || n.Contains("slide") || n.Contains("اسلاید") ||
                 n.Contains("presentation") || n.Contains("پرزنت")))
                return true;

            return explicitExecution || strong;
        }

        private static bool ContainsAny(string text, string[] terms)
        {
            foreach (string term in terms)
                if (!string.IsNullOrWhiteSpace(term) &&
                    text.IndexOf(Normalize(term), StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            return false;
        }

        private static string Normalize(string value)
        {
            return (value ?? "").Replace('ي', 'ی').Replace('ك', 'ک').ToLowerInvariant();
        }
    }
}
