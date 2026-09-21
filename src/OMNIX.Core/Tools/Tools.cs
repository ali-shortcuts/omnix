using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OMNIX.Core.Tools
{
    /// <summary>
    /// Layer 7 — approved OMNIX v3 tool whitelist. Read tools expose bounded Office context only;
    /// write tools require explicit user preview/confirmation. No shell, registry, arbitrary file,
    /// process, or unrestricted system tool exists.
    /// </summary>
    public static class ToolNames
    {
        // read-only
        public const string ReadOfficeAccess = "read_office_access";
        public const string SearchConversation = "search_conversation";
        public const string SearchOfficeReference = "search_office_reference";
        public const string ReadDocumentMap = "read_document_map";
        public const string ReadDocumentSection = "read_document_section";
        public const string ReadSelection = "read_selection";
        public const string ReadDocument = "read_document";
        public const string ReadPresentation = "read_presentation";
        public const string CaptureChartAsImage = "capture_chart_as_image";
        public const string CaptureSlideAsImage = "capture_slide_as_image";
        public const string CaptureCurrentViewAsImage = "capture_current_view_as_image";
        public const string ListOfficeCapabilities = "list_office_capabilities";

        // write (user confirmation + native Office undo)
        public const string CreateDataTable = "create_data_table";
        public const string WriteToCell = "write_to_cell";
        public const string InsertFormula = "insert_formula";
        public const string RewriteSelectedText = "rewrite_selected_text";
        public const string InsertSlide = "insert_slide";
        public const string AddSpeakerNotes = "add_speaker_notes";
        public const string HighlightRange = "highlight_range";
        public const string FormatRange = "format_range";
        public const string ExecuteOfficeCapability = "execute_office_capability";

        private static readonly HashSet<string> Whitelist = new HashSet<string>(StringComparer.Ordinal)
        {
            ReadOfficeAccess, SearchConversation, SearchOfficeReference, ReadDocumentMap, ReadDocumentSection, ReadSelection, ReadDocument, ReadPresentation,
            CaptureChartAsImage, CaptureSlideAsImage, CaptureCurrentViewAsImage, ListOfficeCapabilities,
            CreateDataTable, WriteToCell, InsertFormula, RewriteSelectedText, InsertSlide, AddSpeakerNotes, HighlightRange, FormatRange, ExecuteOfficeCapability
        };

        private static readonly HashSet<string> WriteTools = new HashSet<string>(StringComparer.Ordinal)
        {
            CreateDataTable, WriteToCell, InsertFormula, RewriteSelectedText, InsertSlide, AddSpeakerNotes, HighlightRange, FormatRange, ExecuteOfficeCapability
        };

        public static string Normalize(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return string.Empty;
            string value = name.Trim().Replace('-', '_');
            foreach (string prefix in new[] { "omnix.", "omnix_tool.", "tools.", "functions.", "function." })
                if (value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    value = value.Substring(prefix.Length);
                    break;
                }
            return value.ToLowerInvariant();
        }

        public static bool IsWhitelisted(string name) { return Whitelist.Contains(Normalize(name)); }
        public static bool IsWriteTool(string name) { return WriteTools.Contains(Normalize(name)); }
    }

    public sealed class ToolCall
    {
        public string Name { get; set; }
        public string ArgumentsJson { get; set; }
    }

    public sealed class ToolResult
    {
        public bool Success { get; set; }
        public string ContentForModel { get; set; }
        public string UiNote { get; set; }
        public byte[] CapturedPng { get; set; }

        public static ToolResult Ok(string contentForModel, string uiNote = null)
        {
            return new ToolResult { Success = true, ContentForModel = contentForModel, UiNote = uiNote };
        }

        public static ToolResult Fail(string contentForModel)
        {
            return new ToolResult { Success = false, ContentForModel = contentForModel };
        }
    }

    public sealed class WritePreview
    {
        public string ToolName { get; set; }
        public string Title { get; set; }
        public string Before { get; set; }
        public string After { get; set; }
        public string ArgumentsJson { get; set; }
    }

    /// <summary>Minimal JSON argument accessor with defaults.</summary>
    public sealed class ToolArguments
    {
        private readonly JObject _obj;

        private ToolArguments(JObject obj) { _obj = obj ?? new JObject(); }

        public static ToolArguments Parse(string json)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(json)) return new ToolArguments(new JObject());
                return new ToolArguments(JObject.Parse(json));
            }
            catch
            {
                return new ToolArguments(new JObject());
            }
        }

        public int Integer(string key, int fallback, int minimum, int maximum)
        {
            int value;
            if (!int.TryParse(Get(key, fallback.ToString(System.Globalization.CultureInfo.InvariantCulture)), out value)
                || value < minimum || value > maximum)
                throw new ArgumentException(key + " must be between " + minimum + " and " + maximum + ".");
            return value;
        }

        public string Get(string key, string fallback)
        {
            var token = _obj[key];
            if (token == null) return fallback;
            string s = token.ToString();
            return s ?? fallback;
        }

        public JToken Token(string key)
        {
            return _obj[key];
        }
    }

    /// <summary>
    /// Parses the model's ```omnix_tool {"tool":..., "args":{...}}``` block.
    /// Returns null when the reply contains no tool call.
    /// </summary>
    public static class ToolCallParser
    {
        private const string FencedMarker = "```omnix_tool";
        private const string XmlMarker = "<tool_call>";
        private const string NativeMarker = "<|tool_call_start|>";
        private const string NativeEndMarker = "<|tool_call_end|>";

        public static ToolCall Parse(string reply)
        {
            if (string.IsNullOrEmpty(reply)) return null;

            int fenceIndex = reply.IndexOf(FencedMarker, StringComparison.OrdinalIgnoreCase);
            int xmlIndex = reply.IndexOf(XmlMarker, StringComparison.OrdinalIgnoreCase);
            int nativeIndex = reply.IndexOf(NativeMarker, StringComparison.OrdinalIgnoreCase);
            int idx = FirstNonNegative(fenceIndex, xmlIndex, nativeIndex);

            // Some OpenAI-compatible providers return their native textual function-call syntax
            // without wrapping it in OMNIX's fenced protocol. Accept it only when the ENTIRE reply
            // is one whitelisted call; ordinary prose can never become a tool request accidentally.
            if (idx < 0)
            {
                ToolCall direct = ParseFunctionStyle(reply.Trim());
                if (direct == null) return null;
                direct.Name = ToolNames.Normalize(direct.Name);
                return ToolNames.IsWhitelisted(direct.Name) ? direct : null;
            }

            string body;
            int closeEnd;
            if (idx == nativeIndex)
            {
                int start = idx + NativeMarker.Length;
                int end = reply.IndexOf(NativeEndMarker, start, StringComparison.OrdinalIgnoreCase);
                if (end < 0) return Invalid("Incomplete native tool call");
                body = reply.Substring(start, end - start).Trim();
                closeEnd = end + NativeEndMarker.Length;
            }
            else if (idx == xmlIndex)
            {
                int start = idx + XmlMarker.Length;
                const string closing = "</tool_call>";
                int end = reply.IndexOf(closing, start, StringComparison.OrdinalIgnoreCase);
                if (end < 0) return Invalid("Incomplete tool call");
                body = reply.Substring(start, end - start).Trim();
                if (body.StartsWith("omnix_tool", StringComparison.OrdinalIgnoreCase))
                    body = body.Substring("omnix_tool".Length).Trim();
                closeEnd = end + closing.Length;
            }
            else
            {
                int start = idx + FencedMarker.Length;
                const string closing = "```";
                int end = reply.IndexOf(closing, start, StringComparison.OrdinalIgnoreCase);
                if (end < 0) return Invalid("Incomplete tool call");
                body = reply.Substring(start, end - start).Trim();
                closeEnd = end + closing.Length;
            }

            string remainder = reply.Substring(closeEnd);
            if (ContainsToolMarker(remainder))
                return Invalid("Only one tool call per response is supported");

            ToolCall jsonCall = ParseJsonStyle(body);
            if (jsonCall != null) return jsonCall;

            ToolCall nativeCall = ParseFunctionStyle(body);
            return nativeCall ?? Invalid(body);
        }

        private static ToolCall ParseJsonStyle(string body)
        {
            try
            {
                var obj = JObject.Parse(body);
                string tool = (string)obj["tool"];
                if (string.IsNullOrWhiteSpace(tool)) return Invalid("Missing tool name");
                string args = obj["args"] != null ? obj["args"].ToString(Formatting.None) : "{}";
                return new ToolCall { Name = ToolNames.Normalize(tool), ArgumentsJson = args };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Parses provider-native text such as:
        /// [write_to_cell(sheet='Sheet1', address='B2', value='کد محصول')]
        /// The parser is deliberately data-only: no reflection, dynamic invocation or eval.
        /// It also accepts nested Python-style lists/dictionaries so table plans from compatible
        /// providers can still flow through the normal OMNIX whitelist and confirmation boundary.
        /// </summary>
        private static ToolCall ParseFunctionStyle(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return null;
            body = body.Trim();
            if (body.StartsWith("[", StringComparison.Ordinal) && body.EndsWith("]", StringComparison.Ordinal))
                body = body.Substring(1, body.Length - 2).Trim();

            var match = System.Text.RegularExpressions.Regex.Match(
                body,
                @"^([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)\s*$",
                System.Text.RegularExpressions.RegexOptions.Singleline);
            if (!match.Success) return null;

            string name = match.Groups[1].Value;
            try
            {
                JObject args = ParseNativeArguments(match.Groups[2].Value);
                return new ToolCall { Name = ToolNames.Normalize(name), ArgumentsJson = args.ToString(Formatting.None) };
            }
            catch
            {
                return Invalid("Malformed native arguments");
            }
        }

        private static JObject ParseNativeArguments(string text)
        {
            var obj = new JObject();
            if (string.IsNullOrWhiteSpace(text)) return obj;
            foreach (string part in SplitTopLevel(text, ','))
            {
                if (string.IsNullOrWhiteSpace(part)) continue;
                int eq = IndexOfTopLevel(part, '=');
                if (eq <= 0) throw new FormatException("Expected named argument.");
                string key = part.Substring(0, eq).Trim();
                if (!System.Text.RegularExpressions.Regex.IsMatch(key, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                    throw new FormatException("Invalid argument name.");
                obj[key] = ParseNativeValue(part.Substring(eq + 1));
            }
            return obj;
        }

        private static JToken ParseNativeValue(string raw)
        {
            string value = (raw ?? "").Trim();
            if (value.Length == 0) return JValue.CreateString("");

            if ((value[0] == '\'' && value[value.Length - 1] == '\'') ||
                (value[0] == '"' && value[value.Length - 1] == '"'))
                return new JValue(Unquote(value));

            if (value[0] == '[' && value[value.Length - 1] == ']')
            {
                var array = new JArray();
                string inner = value.Substring(1, value.Length - 2);
                foreach (string item in SplitTopLevel(inner, ','))
                    if (!string.IsNullOrWhiteSpace(item)) array.Add(ParseNativeValue(item));
                return array;
            }

            if (value[0] == '{' && value[value.Length - 1] == '}')
            {
                var result = new JObject();
                string inner = value.Substring(1, value.Length - 2);
                foreach (string item in SplitTopLevel(inner, ','))
                {
                    if (string.IsNullOrWhiteSpace(item)) continue;
                    int colon = IndexOfTopLevel(item, ':');
                    if (colon <= 0) throw new FormatException("Invalid object entry.");
                    string keyRaw = item.Substring(0, colon).Trim();
                    string key = (keyRaw.StartsWith("'") || keyRaw.StartsWith("\""))
                        ? Unquote(keyRaw) : keyRaw;
                    if (string.IsNullOrWhiteSpace(key)) throw new FormatException("Invalid object key.");
                    result[key] = ParseNativeValue(item.Substring(colon + 1));
                }
                return result;
            }

            if (string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)) return new JValue(true);
            if (string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)) return new JValue(false);
            if (string.Equals(value, "null", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
                return JValue.CreateNull();

            long integer;
            if (long.TryParse(value, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out integer))
                return new JValue(integer);

            double number;
            if (double.TryParse(value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out number) &&
                !double.IsNaN(number) && !double.IsInfinity(number))
                return new JValue(number);

            throw new FormatException("Unsupported native literal.");
        }

        private static string Unquote(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length < 2) return value ?? "";
            char quote = value[0];
            if (value[value.Length - 1] != quote) throw new FormatException("Unterminated string.");
            string inner = value.Substring(1, value.Length - 2);
            var sb = new System.Text.StringBuilder();
            bool escape = false;
            for (int i = 0; i < inner.Length; i++)
            {
                char ch = inner[i];
                if (!escape)
                {
                    if (ch == '\\') { escape = true; continue; }
                    sb.Append(ch);
                    continue;
                }
                switch (ch)
                {
                    case 'n': sb.Append('\n'); break;
                    case 'r': sb.Append('\r'); break;
                    case 't': sb.Append('\t'); break;
                    case '\\': sb.Append('\\'); break;
                    case '\'': sb.Append('\''); break;
                    case '"': sb.Append('"'); break;
                    default: sb.Append(ch); break;
                }
                escape = false;
            }
            if (escape) sb.Append('\\');
            return sb.ToString();
        }

        private static List<string> SplitTopLevel(string text, char separator)
        {
            var result = new List<string>();
            if (text == null) return result;
            int depth = 0, start = 0;
            char quote = '\0';
            bool escape = false;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (quote != '\0')
                {
                    if (escape) { escape = false; continue; }
                    if (ch == '\\') { escape = true; continue; }
                    if (ch == quote) quote = '\0';
                    continue;
                }
                if (ch == '\'' || ch == '"') { quote = ch; continue; }
                if (ch == '[' || ch == '{' || ch == '(') depth++;
                else if (ch == ']' || ch == '}' || ch == ')') depth--;
                else if (ch == separator && depth == 0)
                {
                    result.Add(text.Substring(start, i - start));
                    start = i + 1;
                }
                if (depth < 0) throw new FormatException("Unbalanced delimiters.");
            }
            if (quote != '\0' || depth != 0) throw new FormatException("Unbalanced native syntax.");
            result.Add(text.Substring(start));
            return result;
        }

        private static int IndexOfTopLevel(string text, char target)
        {
            if (text == null) return -1;
            int depth = 0;
            char quote = '\0';
            bool escape = false;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (quote != '\0')
                {
                    if (escape) { escape = false; continue; }
                    if (ch == '\\') { escape = true; continue; }
                    if (ch == quote) quote = '\0';
                    continue;
                }
                if (ch == '\'' || ch == '"') { quote = ch; continue; }
                if (ch == '[' || ch == '{' || ch == '(') depth++;
                else if (ch == ']' || ch == '}' || ch == ')') depth--;
                else if (ch == target && depth == 0) return i;
            }
            return -1;
        }

        private static bool ContainsToolMarker(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            return text.IndexOf(FencedMarker, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   text.IndexOf(XmlMarker, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   text.IndexOf(NativeMarker, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static int FirstNonNegative(params int[] values)
        {
            int best = -1;
            foreach (int value in values)
                if (value >= 0 && (best < 0 || value < best)) best = value;
            return best;
        }

        private static ToolCall Invalid(string detail)
        {
            return new ToolCall { Name = "", ArgumentsJson = detail ?? "" };
        }
    }
}
