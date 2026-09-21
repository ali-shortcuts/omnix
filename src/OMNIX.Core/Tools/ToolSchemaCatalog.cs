using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OMNIX.Core.AiGateway;
using OMNIX.Core.Context;

namespace OMNIX.Core.Tools
{
    /// <summary>
    /// Provider-facing native function schemas. These are execution contracts, not new powers:
    /// every name still passes ToolNames whitelist and every write still crosses ToolExecutor
    /// preview/confirmation/scope checks.
    /// </summary>
    public static class ToolSchemaCatalog
    {
        public static List<ToolDefinition> ForHost(IHostAdapter host)
        {
            var tools = new List<ToolDefinition>
            {
                Def(ToolNames.ReadOfficeAccess,
                    "Inspect measured Office editability, protection/read-only state, and whether OMNIX write confirmation is available.",
                    Obj()),
                Def(ToolNames.SearchConversation,
                    "Search earlier saved conversation excerpts for this Office document.",
                    Obj(P("query", "string", "Search text"), Req("query"))),
                Def(ToolNames.SearchOfficeReference,
                    "Search OMNIX educational Office reference by host and topic.",
                    Obj(P("host", "string", "Excel, Word, or PowerPoint"), P("query", "string", "Reference query"), P("offset", "integer", "Pagination offset"))),
                Def(ToolNames.ReadSelection,
                    "Read the current Office selection.",
                    Obj()),
                Def(ToolNames.ReadDocument,
                    "Read a bounded textual/structural representation of the active Office document.",
                    Obj()),
                Def(ToolNames.CaptureCurrentViewAsImage,
                    "Capture the current Office view/selection as an image when pixel-level inspection is needed.",
                    Obj()),
                Def(ToolNames.ListOfficeCapabilities,
                    "Search the implemented advanced capability catalog for the active Office host.",
                    Obj(P("query", "string", "Capability/category search"), P("offset", "integer", "Pagination offset"))),
                Def(ToolNames.ExecuteOfficeCapability,
                    "Execute one exact capability returned by list_office_capabilities. Mutations require OMNIX confirmation.",
                    Obj(P("capability", "string", "Exact capability id"), PObject("args", "Capability-specific arguments"), Req("capability")))
            };

            if (host is IIndexedHostAdapter)
            {
                tools.Add(Def(ToolNames.ReadDocumentMap,
                    "List the active document's bounded structure/containers with pagination.",
                    Obj(P("offset", "integer", "Pagination offset"))));
                tools.Add(Def(ToolNames.ReadDocumentSection,
                    "Read one bounded section of the active document. Arguments depend on the active Office host.",
                    Obj(P("sheet", "string", "Excel worksheet"), P("row", "integer", "Excel one-based row"), P("column", "integer", "Excel one-based column"),
                        P("rows", "integer", "Excel row count"), P("columns", "integer", "Excel column count"),
                        P("story", "string", "Word story name"), P("start", "integer", "Word/PowerPoint text start"),
                        P("count", "integer", "Text count"), P("slide", "integer", "PowerPoint slide"),
                        P("shape", "integer", "PowerPoint shape"), P("part", "string", "PowerPoint part such as notes"))));
            }

            if (host != null && host.Host == HostType.Excel)
            {
                tools.Add(Def(ToolNames.CreateDataTable,
                    "Create one new styled Excel worksheet/table without overwriting an existing sheet.",
                    Obj(P("sheet", "string", "Preferred worksheet name"), P("uniqueName", "boolean", "Resolve a unique suffix if name exists"),
                        PArrayStrings("headers", "Column headers"), PArrayArrays("rows", "Rows of values/formulas/dates"),
                        Req("sheet", "headers", "rows"))));
                tools.Add(Def(ToolNames.WriteToCell,
                    "Write one literal Excel cell value. JSON numbers remain numeric; strings remain literal text.",
                    Obj(P("sheet", "string", "Worksheet; defaults to active"), P("address", "string", "Cell address"), PAny("value", "Text/number/boolean/null"), Req("address", "value"))));
                tools.Add(Def(ToolNames.InsertFormula,
                    "Write one real Excel formula.",
                    Obj(P("sheet", "string", "Worksheet; defaults to active"), P("address", "string", "Cell address"), P("formula", "string", "Formula beginning with ="), Req("address", "formula"))));
                tools.Add(Def(ToolNames.HighlightRange,
                    "Highlight one bounded Excel range.",
                    Obj(P("sheet", "string", "Worksheet; defaults to active"), P("address", "string", "Range address"), Req("address"))));
                tools.Add(Def(ToolNames.FormatRange,
                    "Professionally format one bounded Excel range using the real Excel Object Model.",
                    Obj(P("sheet", "string", "Worksheet; defaults to active"), P("address", "string", "Range address"),
                        P("fontName", "string", "Font name"), P("fontSize", "number", "Font size"), P("bold", "boolean", "Bold"),
                        P("italic", "boolean", "Italic"), P("underline", "boolean", "Underline"), P("fontColor", "string", "#RRGGBB"),
                        P("fillColor", "string", "#RRGGBB"), P("horizontalAlignment", "string", "general|left|center|right"),
                        P("verticalAlignment", "string", "top|center|bottom"), P("numberFormat", "string", "Excel number format"),
                        P("wrapText", "boolean", "Wrap text"), P("border", "string", "none|thin"),
                        P("autofitColumns", "boolean", "AutoFit columns"), P("autofitRows", "boolean", "AutoFit rows"), Req("address"))));
                tools.Add(Def(ToolNames.CaptureChartAsImage,
                    "Capture an Excel chart as an image.",
                    Obj(P("chart", "string", "Chart name; optional"))));
            }
            else if (host != null && host.Host == HostType.Word)
            {
                tools.Add(Def(ToolNames.RewriteSelectedText,
                    "Replace the exact current Word selection after confirmation.",
                    Obj(P("text", "string", "Replacement text"), Req("text"))));
            }
            else if (host != null && host.Host == HostType.PowerPoint)
            {
                tools.Add(Def(ToolNames.ReadPresentation,
                    "Read a bounded textual representation of the presentation.",
                    Obj()));
                tools.Add(Def(ToolNames.InsertSlide,
                    "Insert a PowerPoint slide at a one-based position.",
                    Obj(P("index", "integer", "Slide position"), P("title", "string", "Slide title"), P("body", "string", "Slide body"))));
                tools.Add(Def(ToolNames.AddSpeakerNotes,
                    "Set speaker notes for a PowerPoint slide.",
                    Obj(P("slide", "integer", "Slide number"), P("notes", "string", "Speaker notes"), Req("slide", "notes"))));
                tools.Add(Def(ToolNames.CaptureSlideAsImage,
                    "Capture one PowerPoint slide as an image.",
                    Obj(P("slide", "integer", "Slide number"))));
            }

            return tools;
        }

        private static ToolDefinition Def(string name, string description, JObject schema)
        {
            return new ToolDefinition
            {
                Name = name,
                Description = description,
                ParametersJson = schema.ToString(Formatting.None)
            };
        }

        private static JObject Obj(params JProperty[] properties)
        {
            var props = new JObject();
            var required = new JArray();
            foreach (var p in properties)
            {
                if (p == null) continue;
                if (p.Name == "$required")
                {
                    foreach (var x in (JArray)p.Value) required.Add(x);
                    continue;
                }
                props.Add(p);
            }

            var o = new JObject
            {
                ["type"] = "object",
                ["properties"] = props,
                ["additionalProperties"] = false
            };
            if (required.Count > 0) o["required"] = required;
            return o;
        }

        private static JProperty Req(params string[] names)
        {
            return new JProperty("$required", new JArray(names ?? new string[0]));
        }

        private static JProperty P(string name, string type, string description)
        {
            return new JProperty(name, new JObject { ["type"] = type, ["description"] = description });
        }

        private static JProperty PObject(string name, string description)
        {
            return new JProperty(name, new JObject
            {
                ["type"] = "object",
                ["description"] = description,
                ["additionalProperties"] = true
            });
        }

        private static JProperty PAny(string name, string description)
        {
            return new JProperty(name, new JObject { ["description"] = description });
        }

        private static JProperty PArrayStrings(string name, string description)
        {
            return new JProperty(name, new JObject
            {
                ["type"] = "array",
                ["description"] = description,
                ["items"] = new JObject { ["type"] = "string" }
            });
        }

        private static JProperty PArrayArrays(string name, string description)
        {
            return new JProperty(name, new JObject
            {
                ["type"] = "array",
                ["description"] = description,
                ["items"] = new JObject
                {
                    ["type"] = "array",
                    ["items"] = new JObject()
                }
            });
        }
    }
}
