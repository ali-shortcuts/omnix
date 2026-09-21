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

        // write (user confirmation + native Office undo)
        public const string CreateDataTable = "create_data_table";
        public const string WriteToCell = "write_to_cell";
        public const string InsertFormula = "insert_formula";
        public const string RewriteSelectedText = "rewrite_selected_text";
        public const string InsertSlide = "insert_slide";
        public const string AddSpeakerNotes = "add_speaker_notes";
        public const string HighlightRange = "highlight_range";

        private static readonly HashSet<string> Whitelist = new HashSet<string>(StringComparer.Ordinal)
        {
            SearchConversation, SearchOfficeReference, ReadDocumentMap, ReadDocumentSection, ReadSelection, ReadDocument, ReadPresentation,
            CaptureChartAsImage, CaptureSlideAsImage, CaptureCurrentViewAsImage,
            CreateDataTable, WriteToCell, InsertFormula, RewriteSelectedText, InsertSlide, AddSpeakerNotes, HighlightRange
        };

        private static readonly HashSet<string> WriteTools = new HashSet<string>(StringComparer.Ordinal)
        {
            CreateDataTable, WriteToCell, InsertFormula, RewriteSelectedText, InsertSlide, AddSpeakerNotes, HighlightRange
        };

        public static bool IsWhitelisted(string name) { return !string.IsNullOrEmpty(name) && Whitelist.Contains(name); }
        public static bool IsWriteTool(string name) { return !string.IsNullOrEmpty(name) && WriteTools.Contains(name); }
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
    }

    /// <summary>
    /// Parses the model's ```omnix_tool {"tool":..., "args":{...}}``` block.
    /// Returns null when the reply contains no tool call.
    /// </summary>
    public static class ToolCallParser
    {
        public static ToolCall Parse(string reply)
        {
            if (string.IsNullOrEmpty(reply)) return null;
            const string fenced = "```omnix_tool";
            const string xml = "<tool_call>";
            int fenceIndex = reply.IndexOf(fenced, StringComparison.OrdinalIgnoreCase);
            int xmlIndex = reply.IndexOf(xml, StringComparison.OrdinalIgnoreCase);
            if (fenceIndex < 0 && xmlIndex < 0) return null;
            bool useXml = xmlIndex >= 0 && (fenceIndex < 0 || xmlIndex < fenceIndex);
            string marker = useXml ? xml : fenced;
            int idx = useXml ? xmlIndex : fenceIndex;
            int start = idx + marker.Length;
            string closing = useXml ? "</tool_call>" : "```";
            int end = reply.IndexOf(closing, start, StringComparison.OrdinalIgnoreCase);
            // Malformed or ambiguous calls must produce a tool error, never a success answer.
            if (end < 0) return new ToolCall { Name = "", ArgumentsJson = "Incomplete tool call" };
            string remainder = reply.Substring(end + closing.Length);
            if (remainder.IndexOf(fenced, StringComparison.OrdinalIgnoreCase) >= 0 ||
                remainder.IndexOf(xml, StringComparison.OrdinalIgnoreCase) >= 0)
                return new ToolCall { Name = "", ArgumentsJson = "Only one tool call per response is supported" };
            string body = reply.Substring(start, end - start).Trim();
            if (useXml && body.StartsWith("omnix_tool", StringComparison.OrdinalIgnoreCase))
                body = body.Substring("omnix_tool".Length).Trim();

            try
            {
                var obj = JObject.Parse(body);
                string tool = (string)obj["tool"];
                if (string.IsNullOrWhiteSpace(tool)) return new ToolCall { Name = "", ArgumentsJson = "Missing tool name" };
                string args = obj["args"] != null ? obj["args"].ToString(Formatting.None) : "{}";
                return new ToolCall { Name = tool.Trim(), ArgumentsJson = args };
            }
            catch
            {
                return new ToolCall { Name = "", ArgumentsJson = body };
            }
        }
    }
}
