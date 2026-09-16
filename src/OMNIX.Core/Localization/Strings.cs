using System;
using System.IO;
using System.Windows;
using OMNIX.Core.Logging;

namespace OMNIX.Core.Localization
{
    /// <summary>
    /// Spec 10.7: UI strings live in a central Resource Dictionary (Localization/Strings.xaml),
    /// never hard-coded in XAML or C#. English is the default; the architecture is ready for a
    /// future Persian (RTL) dictionary swap without touching any view.
    /// </summary>
    public static class Strings
    {
        // WPF dictionaries belong to the thread that loads them. Background diagnostics must
        // not publish an empty/foreign-thread dictionary to the Office UI.
        [ThreadStatic]
        private static ResourceDictionary _dict;

        public static ResourceDictionary Dictionary
        {
            get
            {
                if (_dict != null) return _dict;
                // Relative component URI lets WPF initialize its pack URI support itself,
                // including first use from a non-WPF Office/background entry point.
                var loaded = new ResourceDictionary
                {
                    Source = new Uri("/OMNIX.Core;component/Localization/Strings.xaml", UriKind.Relative)
                };
                _dict = loaded; // Cache only after successful load; a failure can be retried.
                return loaded;
            }
        }

        /// <summary>Translate a key; falls back to the key itself when missing (honest diagnostics).</summary>
        public static string T(string key)
        {
            try
            {
                if (Dictionary.Contains(key)) return Dictionary[key] as string;
            }
            catch { }
            return key;
        }
    }
}
