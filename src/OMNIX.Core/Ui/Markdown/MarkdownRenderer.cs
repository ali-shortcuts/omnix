using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Linq;

namespace OMNIX.Core.Ui.Markdown
{
    /// <summary>
    /// Lightweight WPF Markdown renderer tuned for a 360px docked panel (spec Section 5):
    /// headings, bold/italic, inline code, fenced code blocks with simple syntax highlighting,
    /// real tables (not raw text), bullet/numbered lists, links, blockquotes and horizontal rules.
    /// Excel formulas (lines starting with "=" or ```excel blocks) render in monospace.
    /// Zero external dependencies — deterministic behavior inside Office.
    /// </summary>
    public static class MarkdownRenderer
    {
        [ThreadStatic] private static FlowDocument _renderDocument;
        private static readonly Regex InlineCode = new Regex("`([^`]+)`", RegexOptions.Compiled);
        private static readonly Regex Bold = new Regex(@"\*\*([^*]+)\*\*", RegexOptions.Compiled);
        private static readonly Regex Italic = new Regex(@"(?<!\*)\*([^*\n]+)\*(?!\*)", RegexOptions.Compiled);
        private static readonly Regex Link = new Regex(@"\[([^\]]+)\]\(([^)]+)\)", RegexOptions.Compiled);

        public static void Render(FlowDocument doc, string markdown)
        {
            _renderDocument = doc;
            doc.Blocks.Clear();
            if (string.IsNullOrEmpty(markdown)) return;

            string[] lines = markdown.Replace("\r\n", "\n").Split('\n');
            int i = 0;
            while (i < lines.Length)
            {
                string line = lines[i];

                if (line.StartsWith("```", StringComparison.Ordinal))
                {
                    string lang = line.Substring(3).Trim();
                    var codeLines = new List<string>();
                    i++;
                    while (i < lines.Length && !lines[i].StartsWith("```", StringComparison.Ordinal))
                    {
                        codeLines.Add(lines[i]);
                        i++;
                    }
                    i++; // closing fence
                    doc.Blocks.Add(BuildCodeBlock(string.Join(Environment.NewLine, codeLines), lang));
                    continue;
                }

                if (line.TrimStart().StartsWith("|") && i + 1 < lines.Length &&
                    lines[i + 1].TrimStart().StartsWith("|") && lines[i + 1].Contains("--"))
                {
                    // Markdown table with separator row
                    var header = SplitRow(line);
                    i += 2;
                    var rows = new List<string[]>();
                    while (i < lines.Length && lines[i].TrimStart().StartsWith("|"))
                    {
                        rows.Add(SplitRow(lines[i]));
                        i++;
                    }
                    doc.Blocks.Add(BuildTable(header, rows));
                    continue;
                }

                string trimmed = line.TrimStart();
                if (trimmed.StartsWith("#"))
                {
                    int level = 0;
                    while (level < trimmed.Length && trimmed[level] == '#') level++;
                    var p = new Paragraph();
                    p.FontSize = level <= 1 ? 16 : (level == 2 ? 14.5 : 13.5);
                    p.FontWeight = FontWeights.Bold;
                    p.Margin = new Thickness(0, 8, 0, 4);
                    AddInlineRuns(p, trimmed.Substring(level).Trim());
                    doc.Blocks.Add(p);
                    i++;
                    continue;
                }

                if (trimmed.StartsWith("---") || trimmed.StartsWith("***"))
                {
                    var hr = new Paragraph(new Run(new string('─', 34)))
                    {
                        Foreground = TryFindBrush("B.Border"),
                        Margin = new Thickness(0, 4, 0, 4)
                    };
                    doc.Blocks.Add(hr);
                    i++;
                    continue;
                }

                if (trimmed.StartsWith("> "))
                {
                    var p = new Paragraph
                    {
                        Margin = new Thickness(10, 2, 0, 2),
                        Padding = new Thickness(6, 2, 0, 2),
                        BorderBrush = TryFindBrush("B.Accent"),
                        BorderThickness = new Thickness(2, 0, 0, 0)
                    };
                    AddInlineRuns(p, trimmed.Substring(2));
                    doc.Blocks.Add(p);
                    i++;
                    continue;
                }

                if (trimmed.StartsWith("- ") || trimmed.StartsWith("* ") || trimmed.StartsWith("• "))
                {
                    var p = new Paragraph { Margin = new Thickness(12, 1, 0, 1) };
                    p.Inlines.Add(new Run("•  ") { Foreground = TryFindBrush("B.Accent") });
                    AddInlineRuns(p, trimmed.Substring(2));
                    doc.Blocks.Add(p);
                    i++;
                    continue;
                }

                if (trimmed.Length > 2 && char.IsDigit(trimmed[0]) && trimmed[1] == '.' ||
                    trimmed.Length > 3 && char.IsDigit(trimmed[0]) && char.IsDigit(trimmed[1]) && trimmed[2] == '.')
                {
                    int dot = trimmed.IndexOf('.');
                    var p = new Paragraph { Margin = new Thickness(12, 1, 0, 1) };
                    p.Inlines.Add(new Run(trimmed.Substring(0, dot + 1) + "  ") { Foreground = TryFindBrush("B.Accent") });
                    AddInlineRuns(p, trimmed.Substring(dot + 1).TrimStart());
                    doc.Blocks.Add(p);
                    i++;
                    continue;
                }

                if (trimmed.Length == 0)
                {
                    i++;
                    continue;
                }

                // Excel formula line (monospace styling)
                if (trimmed.StartsWith("=") || trimmed.StartsWith("+SUM") || trimmed.StartsWith("=SUM"))
                {
                    doc.Blocks.Add(BuildCodeBlock(trimmed, "excel"));
                    i++;
                    continue;
                }

                // Paragraph (merge consecutive lines)
                var paraLines = new List<string> { line };
                i++;
                while (i < lines.Length)
                {
                    string nx = lines[i];
                    if (nx.TrimStart().Length == 0 || IsBlockStart(nx)) break;
                    paraLines.Add(nx);
                    i++;
                }
                var bodyPara = new Paragraph { Margin = new Thickness(0, 2, 0, 2) };
                AddInlineRuns(bodyPara, string.Join(" ", paraLines));
                doc.Blocks.Add(bodyPara);
            }
        }

        private static bool IsBlockStart(string line)
        {
            string t = line.TrimStart();
            return t.StartsWith("#") || t.StartsWith("```") || t.StartsWith("> ") ||
                   t.StartsWith("- ") || t.StartsWith("* ") || t.StartsWith("|") ||
                   t.StartsWith("---");
        }

        private static string[] SplitRow(string row)
        {
            string t = row.Trim();
            if (t.StartsWith("|")) t = t.Substring(1);
            if (t.EndsWith("|")) t = t.Substring(0, t.Length - 1);
            return t.Split('|');
        }

        private static Table BuildTable(string[] header, List<string[]> rows)
        {
            var table = new Table { CellSpacing = 0, Margin = new Thickness(0, 4, 0, 6) };
            int cols = header.Length;
            AddColumns(table, cols);

            var headerRow = new TableRow { Background = TryFindBrush("B.Surface") };
            foreach (string h in header)
            {
                var cell = new TableCell(new Paragraph(new Run(h.Trim())
                {
                    FontWeight = FontWeights.Bold,
                    FontSize = 11
                }))
                { Padding = new Thickness(4) };
                headerRow.Cells.Add(cell);
            }
            var rg = new TableRowGroup();
            rg.Rows.Add(headerRow);
            table.RowGroups.Add(rg);

            var bodyGroup = new TableRowGroup();
            foreach (var row in rows.Take(40))
            {
                var tr = new TableRow();
                for (int c = 0; c < cols; c++)
                {
                    string val = c < row.Length ? row[c].Trim() : "";
                    var cell = new TableCell(new Paragraph(new Run(val) { FontSize = 11 }))
                    {
                        Padding = new Thickness(4),
                        BorderBrush = TryFindBrush("B.Border"),
                        BorderThickness = new Thickness(0, 0, 0, 0.5)
                    };
                    tr.Cells.Add(cell);
                }
                bodyGroup.Rows.Add(tr);
            }
            if (rows.Count > 40)
            {
                var noteRow = new TableRow();
                var noteCell = new TableCell(new Paragraph(new Run("… " + (rows.Count - 40) + " more rows") { FontStyle = FontStyles.Italic, FontSize = 10 }))
                { Padding = new Thickness(4) };
                noteRow.Cells.Add(noteCell);
                bodyGroup.Rows.Add(noteRow);
            }
            table.RowGroups.Add(bodyGroup);
            return table;
        }

        private static TableColumn[] EnumerableRange(int n)
        {
            var list = new TableColumn[n];
            for (int i = 0; i < n; i++) list[i] = new TableColumn();
            return list;
        }

        private static void AddColumns(Table table, int count)
        {
            foreach (var col in EnumerableRange(count))
                table.Columns.Add(col);
        }

        private static Block BuildCodeBlock(string code, string lang)
        {
            var table = new Table { CellSpacing = 0, Margin = new Thickness(0, 4, 0, 6) };
            AddColumns(table, 1);
            var row = new TableRow();
            var cell = new TableCell { Padding = new Thickness(6), Background = TryFindBrush("B.CodeBackground") };

            bool isExcel = string.Equals(lang, "excel", StringComparison.OrdinalIgnoreCase) ||
                           string.Equals(lang, "formula", StringComparison.OrdinalIgnoreCase);
            var para = new Paragraph { Margin = new Thickness(0) };

            if (string.IsNullOrEmpty(lang))
            {
                para.Inlines.Add(new Run(code)
                {
                    FontFamily = new FontFamily("Consolas, Courier New"),
                    FontSize = 11
                });
            }
            else
            {
                CodeHighlighter.Highlight(para, code, lang, isExcel);
            }
            cell.Blocks.Add(para);
            row.Cells.Add(cell);
            var group = new TableRowGroup();
            group.Rows.Add(row);
            table.RowGroups.Add(group);
            return table;
        }

        internal static void AddInlineRuns(Paragraph p, string text)
        {
            int pos = 0;
            var matches = new List<MatchToken>();

            CollectMatches(matches, InlineCode.Matches(text), InlineMatchType.Code);
            CollectMatches(matches, Link.Matches(text), InlineMatchType.Link);
            CollectMatches(matches, Bold.Matches(text), InlineMatchType.Bold);
            CollectMatches(matches, Italic.Matches(text), InlineMatchType.Italic);
            matches.Sort((a, b) => a.Index.CompareTo(b.Index));

            // resolve overlaps (keep earliest, drop nested)
            var accepted = new List<MatchToken>();
            int lastEnd = -1;
            foreach (var m in matches)
            {
                if (m.Index >= lastEnd)
                {
                    accepted.Add(m);
                    lastEnd = m.Index + m.Length;
                }
            }

            foreach (var m in accepted)
            {
                if (m.Index > pos)
                    p.Inlines.Add(new Run(text.Substring(pos, m.Index - pos)));

                string inner = text.Substring(m.Index, m.Length);
                switch (m.Type)
                {
                    case InlineMatchType.Code:
                        p.Inlines.Add(new Run(Strip(inner, 1))
                        {
                            FontFamily = new FontFamily("Consolas, Courier New"),
                            FontSize = 11,
                            Background = TryFindBrush("B.CodeBackground"),
                            Foreground = TryFindBrush("B.CodeKeyword")
                        });
                        break;
                    case InlineMatchType.Bold:
                        p.Inlines.Add(new Run(Strip(inner, 2)) { FontWeight = FontWeights.Bold });
                        break;
                    case InlineMatchType.Italic:
                        p.Inlines.Add(new Run(Strip(inner, 1)) { FontStyle = FontStyles.Italic });
                        break;
                    case InlineMatchType.Link:
                        var lm = Link.Match(inner);
                        var run = new Run(lm.Groups[1].Value)
                        {
                            Foreground = TryFindBrush("B.Link"),
                            TextDecorations = TextDecorations.Underline
                        };
                        string url = lm.Groups[2].Value;
                        run.ToolTip = url;
                        var href = new System.Windows.Documents.Hyperlink(run)
                        {
                            NavigateUri = new Uri(url, UriKind.RelativeOrAbsolute)
                        };
                        href.RequestNavigate += delegate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
                        {
                            Util.ProcessLauncher.Open(e.Uri.OriginalString);
                            e.Handled = true;
                        };
                        p.Inlines.Add(href);
                        break;
                }
                pos = m.Index + m.Length;
            }
            if (pos < text.Length)
                p.Inlines.Add(new Run(text.Substring(pos)));
        }

        private static string Strip(string s, int n) { return s.Length > 2 * n ? s.Substring(n, s.Length - 2 * n) : s; }

        private static void CollectMatches(List<MatchToken> list, MatchCollection ms, InlineMatchType type)
        {
            foreach (Match m in ms)
                list.Add(new MatchToken { Index = m.Index, Length = m.Length, Type = type });
        }

        internal static Brush TryFindBrush(string key)
        {
            try
            {
                object v = _renderDocument != null ? _renderDocument.TryFindResource(key) : null;
                if (v is Brush) return (Brush)v;
            }
            catch { }
            return key.Contains("Border") ? Brushes.LightGray : Brushes.Black;
        }

        private enum InlineMatchType { Code, Bold, Italic, Link }

        private sealed class MatchToken
        {
            public int Index;
            public int Length;
            public InlineMatchType Type;
        }
    }

    /// <summary>Very small, deterministic syntax highlighter for the code blocks.</summary>
    public static class CodeHighlighter
    {
        private static readonly string[] CsKeywords =
            { "public", "private", "internal", "protected", "static", "void", "class", "new", "return", "if", "else", "for", "foreach", "while", "var", "int", "string", "bool", "double", "null", "true", "false", "async", "await", "using", "namespace", "this" };
        private static readonly string[] PyKeywords =
            { "def", "class", "if", "elif", "else", "for", "while", "import", "from", "return", "True", "False", "None", "and", "or", "not", "in", "is" };
        private static readonly string[] JsKeywords =
            { "const", "let", "var", "function", "return", "if", "else", "for", "while", "new", "class", "async", "await", "true", "false", "null", "=>", "import", "export" };
        private static readonly string[] SqlKeywords =
            { "SELECT", "FROM", "WHERE", "JOIN", "LEFT", "INNER", "GROUP", "BY", "ORDER", "INSERT", "INTO", "VALUES", "UPDATE", "SET", "DELETE", "CREATE", "TABLE", "AS", "ON", "AND", "OR", "NOT", "NULL" };

        public static void Highlight(Paragraph p, string code, string lang, bool isExcel)
        {
            string[] keywords;
            switch ((lang ?? "").ToLowerInvariant())
            {
                case "cs":
                case "csharp":
                case "c#": keywords = CsKeywords; break;
                case "py":
                case "python": keywords = PyKeywords; break;
                case "js":
                case "javascript":
                case "ts":
                case "typescript": keywords = JsKeywords; break;
                case "sql": keywords = SqlKeywords; break;
                default: keywords = null; break;
            }

            if (isExcel || keywords == null)
            {
                // Monospace render; for Excel formulas highlight the leading = and function names.
                foreach (string line in code.Split('\n'))
                {
                    AppendFormulaOrPlain(p, line.TrimEnd('\r'), isExcel);
                }
                return;
            }

            string commentToken = (lang == "py") ? "#" : "//";
            foreach (string rawLine in code.Split('\n'))
            {
                string line = rawLine.TrimEnd('\r');
                int commentIdx = line.IndexOf(commentToken, StringComparison.Ordinal);
                string codePart = commentIdx >= 0 ? line.Substring(0, commentIdx) : line;
                AppendKeywords(p, codePart, keywords);
                if (commentIdx >= 0)
                    p.Inlines.Add(new Run(line.Substring(commentIdx)) { Foreground = MarkdownRenderer.TryFindBrush("B.CodeComment"), FontFamily = Mono() });
                p.Inlines.Add(new Run(Environment.NewLine) { FontFamily = Mono() });
            }
        }

        private static FontFamily Mono() { return new FontFamily("Consolas, Courier New"); }

        private static void AppendKeywords(Paragraph p, string line, string[] keywords)
        {
            int pos = 0;
            var tokens = new List<KeyValuePair<int, int>>();
            foreach (string kw in keywords)
            {
                int start = 0;
                while (true)
                {
                    int idx = line.IndexOf(kw, start, StringComparison.Ordinal);
                    if (idx < 0) break;
                    bool boundaryLeft = idx == 0 || !char.IsLetterOrDigit(line[idx - 1]);
                    bool boundaryRight = idx + kw.Length >= line.Length || !char.IsLetterOrDigit(line[idx + kw.Length]);
                    if (boundaryLeft && boundaryRight) tokens.Add(new KeyValuePair<int, int>(idx, kw.Length));
                    start = idx + kw.Length;
                }
            }
            tokens.Sort((a, b) => a.Key.CompareTo(b.Key));
            int lastEnd = -1;
            var accepted = new List<KeyValuePair<int, int>>();
            foreach (var t in tokens)
            {
                if (t.Key >= lastEnd) { accepted.Add(t); lastEnd = t.Key + t.Value; }
            }

            foreach (var t in accepted)
            {
                if (t.Key > pos) AppendPlain(p, line.Substring(pos, t.Key - pos));
                p.Inlines.Add(new Run(line.Substring(t.Key, t.Value))
                {
                    Foreground = MarkdownRenderer.TryFindBrush("B.CodeKeyword"),
                    FontFamily = Mono()
                });
                pos = t.Key + t.Value;
            }
            if (pos < line.Length) AppendPlain(p, line.Substring(pos));
        }

        private static void AppendPlain(Paragraph p, string s)
        {
            if (s.Length == 0) return;
            p.Inlines.Add(new Run(s) { FontFamily = Mono(), Foreground = MarkdownRenderer.TryFindBrush("B.Foreground") });
        }

        private static void AppendFormulaOrPlain(Paragraph p, string line, bool isExcel)
        {
            if (isExcel && line.TrimStart().StartsWith("="))
            {
                int paren = line.IndexOf('(');
                if (paren > 1)
                {
                    p.Inlines.Add(new Run(line.Substring(0, paren)) { Foreground = MarkdownRenderer.TryFindBrush("B.CodeKeyword"), FontFamily = Mono() });
                    p.Inlines.Add(new Run(line.Substring(paren)) { FontFamily = Mono(), Foreground = MarkdownRenderer.TryFindBrush("B.Foreground") });
                }
                else
                {
                    p.Inlines.Add(new Run(line) { FontFamily = Mono(), Foreground = MarkdownRenderer.TryFindBrush("B.CodeKeyword") });
                }
            }
            else
            {
                p.Inlines.Add(new Run(line) { FontFamily = Mono(), Foreground = MarkdownRenderer.TryFindBrush("B.Foreground") });
            }
            p.Inlines.Add(new Run(Environment.NewLine) { FontFamily = Mono() });
        }
    }
}
