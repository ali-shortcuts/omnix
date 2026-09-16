using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Xml;

namespace OMNIX.Core.Localization
{
    /// <summary>One embedded string source, usable before WPF starts and from worker threads.</summary>
    public static class Strings
    {
        private static readonly Lazy<Dictionary<string, string>> Values =
            new Lazy<Dictionary<string, string>>(LoadValues, true);
        [ThreadStatic]
        private static ResourceDictionary _dict;

        private static Dictionary<string, string> LoadValues()
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            using (var stream = typeof(Strings).Assembly.GetManifestResourceStream("OMNIX.Core.Localization.Strings.source.xaml"))
            {
                if (stream == null) throw new InvalidOperationException("Embedded localization source is missing.");
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                using (var reader = XmlReader.Create(stream, settings))
                {
                    var document = new XmlDocument { XmlResolver = null };
                    document.Load(reader);
                    foreach (XmlNode node in document.DocumentElement.ChildNodes)
                    {
                        var element = node as XmlElement;
                        if (element == null || element.LocalName != "String") continue;
                        var key = element.GetAttribute("Key", "http://schemas.microsoft.com/winfx/2006/xaml");
                        if (!string.IsNullOrEmpty(key)) result.Add(key, element.InnerText);
                    }
                }
            }
            return result;
        }

        public static ResourceDictionary Dictionary
        {
            get
            {
                if (_dict != null) return _dict;
                var dictionary = new ResourceDictionary();
                foreach (var item in Values.Value) dictionary.Add(item.Key, item.Value);
                _dict = dictionary;
                return dictionary;
            }
        }

        public static string T(string key)
        {
            string value;
            return Values.Value.TryGetValue(key, out value) ? value : key;
        }
    }
}
