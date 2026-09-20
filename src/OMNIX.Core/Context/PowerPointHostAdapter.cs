using System;
using System.Text;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using OMNIX.Core.Errors;
using OMNIX.Core.Tools;
using OMNIX.Core.Util;

namespace OMNIX.Core.Context
{
    /// <summary>
    /// PowerPoint adapter (spec Section 3, Layer 3): Presentation, current slide as image for
    /// Vision, speaker notes, shapes/text. Write tools: insert_slide, add_speaker_notes.
    /// </summary>
    public sealed class PowerPointHostAdapter : IHostAdapter, IIndexedHostAdapter
    {
        private const int MaxSlideTitleChars = 500;
        private const int MaxSlideBodyChars = 20000;
        private const int MaxSpeakerNotesChars = 20000;
        private const int PreviewChars = 4000;

        private readonly Ppt.Application _app;
        private readonly Func<int> _maxChars;

        public PowerPointHostAdapter(Ppt.Application app, Func<int> maxChars)
        {
            _app = app;
            _maxChars = maxChars;
        }

        public HostType Host { get { return HostType.PowerPoint; } }
        public string HostDisplayName { get { return "PowerPoint"; } }

        public OfficeContext ReadContext()
        {
            var ctx = new OfficeContext { Host = HostType.PowerPoint };
            try
            {
                var pres = _app.ActivePresentation;
                if (pres == null) return ctx;
                ctx.DocumentName = pres.Name;
                ctx.DocumentPath = pres.FullName;
                ctx.SlideCount = pres.Slides.Count;

                Ppt.Slide slide = ResolveActiveSlide();
                if (slide != null)
                {
                    ctx.CurrentSlideIndex = slide.SlideIndex;
                    ctx.SlideTitle = GetSlideTitle(slide);
                    ctx.NotesPreview = GetNotes(slide, Math.Min(1000, Math.Max(1, _maxChars())));
                    ctx.SelectionAddress = "Slide " + slide.SlideIndex;
                }
                else
                {
                    ctx.CurrentSlideIndex = 0;
                }
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "PowerPointHostAdapter.ReadContext failed", ex);
            }
            return ctx;
        }

        public string ReadSelection()
        {
            try
            {
                var pres = _app.ActivePresentation;
                if (pres == null) return "(no presentation open)";
                Ppt.Slide slide = ResolveActiveSlide();
                if (slide == null) return "(no active slide)";

                int cap = Math.Max(1, _maxChars());
                var sb = new StringBuilder();
                sb.AppendLine("Slide " + slide.SlideIndex + " of " + pres.Slides.Count);
                sb.AppendLine("Title: " + GetSlideTitle(slide));
                sb.AppendLine("Shapes text:");
                AppendShapesText(slide, sb, Math.Max(1, cap - sb.Length - 256));
                if (sb.Length < cap)
                {
                    int remaining = Math.Max(1, cap - sb.Length - 32);
                    sb.AppendLine("Speaker notes: " + GetNotes(slide, remaining));
                }
                return TextUtil.Truncate(sb.ToString(), cap);
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "PowerPointHostAdapter.ReadSelection failed", ex);
                return "(unable to read slide)";
            }
        }

        public string ReadDocument(int maxChars)
        {
            try
            {
                var pres = _app.ActivePresentation;
                if (pres == null) return "(no presentation open)";

                int cap = Math.Max(1, Math.Min(maxChars, _maxChars()));
                var sb = new StringBuilder();
                sb.AppendLine("Presentation: " + pres.Name + " (" + pres.Slides.Count + " slides)");
                for (int i = 1; i <= pres.Slides.Count && sb.Length < cap; i++)
                {
                    var slide = pres.Slides[i];
                    sb.AppendLine();
                    sb.AppendLine("--- Slide " + i + " ---");
                    sb.AppendLine("Title: " + GetSlideTitle(slide));
                    int remaining = Math.Max(1, cap - sb.Length);
                    AppendShapesText(slide, sb, remaining);
                }
                return TextUtil.Truncate(sb.ToString(), cap);
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "PowerPointHostAdapter.ReadDocument failed", ex);
                return "(unable to read presentation)";
            }
        }

        public string ReadDocumentMap(int offset)
        {
            var pres = _app.ActivePresentation;
            int total = pres.Slides.Count, end = Math.Min(total, offset + 20);
            var sb = new StringBuilder();
            sb.AppendLine("Slides=" + total + "; offset=" + offset);
            for (int i = offset + 1; i <= end; i++)
            {
                var slide = pres.Slides[i];
                sb.AppendLine("slide=" + i + "; shapes=" + slide.Shapes.Count
                    + "; title=" + TextUtil.Truncate(GetSlideTitle(slide), 120));
            }
            sb.AppendLine("nextOffset=" + (end < total ? end.ToString() : "none"));
            sb.AppendLine("Read each shape by slide/shape/start/count. Shape text is not a full visual inspection; tables, groups and embedded objects may require slide capture.");
            return sb.ToString();
        }

        public string ReadDocumentSection(ToolArguments args)
        {
            var pres = _app.ActivePresentation;
            int slideIndex = args.Integer("slide", 1, 1, pres.Slides.Count);
            var slide = pres.Slides[slideIndex];
            int shapeIndex = args.Integer("shape", 1, 1, slide.Shapes.Count);
            var shape = slide.Shapes[shapeIndex];
            if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                return "This shape has no text frame. Use capture_slide_as_image for its visible content.";
            var range = shape.TextFrame.TextRange;
            int start = args.Integer("start", 0, 0, range.Length);
            int count = args.Integer("count", 3000, 1, 4000);
            int length = Math.Min(count, range.Length - start);
            string text = length == 0 ? "" : range.Characters(start + 1, length).Text;
            return "Slide=" + slideIndex + "; shape=" + shapeIndex + "; text [" + start + "," + (start+length)
                + "); nextStart=" + (start+length < range.Length ? (start+length).ToString() : "none") + "\n" + text;
        }

        public byte[] CaptureChartAsImage(string chartName) { return null; }

        public byte[] CaptureSlideAsImage(int slideIndexOneBased)
        {
            try
            {
                var pres = _app.ActivePresentation;
                if (pres == null) return null;

                Ppt.Slide slide = null;
                if (slideIndexOneBased >= 1 && slideIndexOneBased <= pres.Slides.Count)
                    slide = pres.Slides[slideIndexOneBased];
                if (slide == null) slide = ResolveActiveSlide();
                if (slide == null) return null;

                return TempImageCapture.FromExporter(path => slide.Export(path, "PNG", 1280, 720));
            }
            catch (Exception ex)
            {
                Logging.Logger.Error("startup-debug", "PowerPointHostAdapter.CaptureSlideAsImage failed", ex);
                return null;
            }
        }

        public byte[] CaptureCurrentViewAsImage()
        {
            return CaptureSlideAsImage(0);
        }

        public WritePreview PrepareWrite(string toolName, string argumentsJson)
        {
            var args = ToolArguments.Parse(argumentsJson);
            switch (toolName)
            {
                case ToolNames.InsertSlide:
                {
                    string title = RequireBounded(args.Get("title", "") ?? "", MaxSlideTitleChars, "slide title", toolName);
                    string body = RequireBounded(args.Get("body", "") ?? "", MaxSlideBodyChars, "slide body", toolName);
                    int requested = ParseInt(args.Get("index", "0"), 0);
                    int resolved = NormalizeInsertIndex(requested);
                    return new WritePreview
                    {
                        ToolName = toolName,
                        Title = "PowerPoint — insert slide",
                        Before = "Presentation currently has " + ActiveSlideCount() + " slides.",
                        After = "A new slide will be added at position " + resolved +
                                " with title: " + TextUtil.Truncate(string.IsNullOrEmpty(title) ? "(empty)" : title, 500) +
                                (string.IsNullOrEmpty(body) ? "" : Environment.NewLine +
                                 "Body characters: " + body.Length + Environment.NewLine + TextUtil.Truncate(body, PreviewChars)),
                        ArgumentsJson = argumentsJson
                    };
                }
                case ToolNames.AddSpeakerNotes:
                {
                    int idx = ParseInt(args.Get("slide", "0"), 0);
                    Ppt.Slide slide = GetSlide(idx);
                    if (slide == null)
                        throw new OmnixException(ErrorCode.CORE_ERROR, "Cannot resolve target slide.", toolName, "Open or select the target slide.");
                    string notes = RequireBounded(args.Get("notes", "") ?? "", MaxSpeakerNotesChars, "speaker notes", toolName);
                    string before = GetNotes(slide, PreviewChars);
                    return new WritePreview
                    {
                        ToolName = toolName,
                        Title = "PowerPoint — add speaker notes",
                        Before = "Current notes of slide " + slide.SlideIndex + ": " + TextUtil.Truncate(before, PreviewChars),
                        After = "New notes characters: " + notes.Length + Environment.NewLine + TextUtil.Truncate(notes, PreviewChars) +
                                (notes.Length > PreviewChars ? Environment.NewLine + "…[notes preview truncated]" : ""),
                        ArgumentsJson = argumentsJson
                    };
                }
                default:
                    throw new OmnixException(ErrorCode.CORE_ERROR,
                        "Tool '" + toolName + "' is not supported by PowerPoint.",
                        "PowerPointHostAdapter.PrepareWrite", "Use insert_slide or add_speaker_notes.");
            }
        }

        public void ApplyWrite(string toolName, string argumentsJson)
        {
            var args = ToolArguments.Parse(argumentsJson);
            var pres = _app.ActivePresentation;
            if (pres == null)
                throw new OmnixException(ErrorCode.CORE_ERROR, "No presentation open.", toolName, "Open a presentation.");

            try { _app.StartNewUndoEntry(); } catch { }

            switch (toolName)
            {
                case ToolNames.InsertSlide:
                {
                    string title = RequireBounded(args.Get("title", "") ?? "", MaxSlideTitleChars, "slide title", toolName);
                    string body = RequireBounded(args.Get("body", "") ?? "", MaxSlideBodyChars, "slide body", toolName);
                    int requested = ParseInt(args.Get("index", "0"), 0);
                    int index = NormalizeInsertIndex(requested);
                    var slide = pres.Slides.Add(index, Ppt.PpSlideLayout.ppLayoutText);
                    if (!string.IsNullOrEmpty(title) && slide.Shapes.Placeholders.Count >= 1)
                        slide.Shapes.Placeholders[1].TextFrame.TextRange.Text = title;
                    if (!string.IsNullOrEmpty(body) && slide.Shapes.Placeholders.Count >= 2)
                        slide.Shapes.Placeholders[2].TextFrame.TextRange.Text = body;
                    break;
                }
                case ToolNames.AddSpeakerNotes:
                {
                    int idx = ParseInt(args.Get("slide", "0"), 0);
                    Ppt.Slide slide = GetSlide(idx);
                    if (slide == null)
                        throw new OmnixException(ErrorCode.CORE_ERROR, "Cannot resolve target slide.", toolName, "Open or select the target slide.");
                    string notes = RequireBounded(args.Get("notes", "") ?? "", MaxSpeakerNotesChars, "speaker notes", toolName);
                    var noteShape = GetNotesBodyShape(slide);
                    if (noteShape == null || noteShape.TextFrame == null)
                        throw new OmnixException(ErrorCode.CORE_ERROR,
                            "Speaker-notes placeholder is unavailable on this slide.", toolName,
                            "Use a presentation/slide layout with a standard notes body placeholder.");
                    noteShape.TextFrame.TextRange.Text = notes;
                    break;
                }
                default:
                    throw new OmnixException(ErrorCode.CORE_ERROR, "Unknown PowerPoint write tool: " + toolName, "", "");
            }
            Logging.Logger.Install("PowerPoint write tool applied: " + toolName);
        }

        private Ppt.Slide ResolveActiveSlide()
        {
            var win = _app.ActiveWindow;
            if (win == null) return null;
            try { return win.View.Slide as Ppt.Slide; }
            catch { return null; }
        }

        private Ppt.Slide GetSlide(int indexOneBased)
        {
            var pres = _app.ActivePresentation;
            if (pres == null) return null;
            if (indexOneBased >= 1 && indexOneBased <= pres.Slides.Count) return pres.Slides[indexOneBased];
            return ResolveActiveSlide();
        }

        private int NormalizeInsertIndex(int requested)
        {
            int count = ActiveSlideCount();
            if (requested <= 0) return count + 1;
            if (requested < 1) return 1;
            if (requested > count + 1) return count + 1;
            return requested;
        }

        private int ActiveSlideCount()
        {
            var pres = _app.ActivePresentation;
            return pres != null ? pres.Slides.Count : 0;
        }

        private static int ParseInt(string s, int fallback)
        {
            int v;
            return int.TryParse(s, out v) ? v : fallback;
        }

        private static string RequireBounded(string value, int limit, string label, string toolName)
        {
            value = value ?? "";
            if (value.Length > limit)
                throw new OmnixException(ErrorCode.CORE_ERROR,
                    "PowerPoint " + label + " is too large for one AI-approved mutation.",
                    toolName + " " + label + " chars=" + value.Length + "; limit=" + limit + ".",
                    "Split the change into smaller explicit steps.");
            return value;
        }

        private static string GetSlideTitle(Ppt.Slide slide)
        {
            try
            {
                if (slide != null && slide.Shapes.Title != null)
                {
                    var range = slide.Shapes.Title.TextFrame.TextRange;
                    int length = 0;
                    try { length = range.Length; } catch { }
                    if (length > 0)
                        return TextUtil.Truncate(range.Characters(1, Math.Min(length, MaxSlideTitleChars)).Text, MaxSlideTitleChars);
                    return TextUtil.Truncate(range.Text, MaxSlideTitleChars);
                }
            }
            catch { }
            return "(no title)";
        }

        private static Ppt.Shape GetNotesBodyShape(Ppt.Slide slide)
        {
            if (slide == null) return null;
            try
            {
                var placeholders = slide.NotesPage.Shapes.Placeholders;
                for (int i = 1; i <= placeholders.Count; i++)
                {
                    var shape = placeholders[i];
                    try
                    {
                        if (shape.PlaceholderFormat.Type == Ppt.PpPlaceholderType.ppPlaceholderBody)
                            return shape;
                    }
                    catch { }
                }
                if (placeholders.Count >= 2) return placeholders[2];
            }
            catch { }
            return null;
        }

        private static string GetNotes(Ppt.Slide slide, int maxChars)
        {
            try
            {
                var shape = GetNotesBodyShape(slide);
                if (shape == null || shape.TextFrame == null) return "";
                var range = shape.TextFrame.TextRange;
                int cap = Math.Max(1, maxChars);
                int length = 0;
                try { length = range.Length; } catch { }
                if (length > 0)
                    return range.Characters(1, Math.Min(length, cap)).Text ?? "";
                return TextUtil.Truncate(range.Text ?? "", cap);
            }
            catch { return ""; }
        }

        private static void AppendShapesText(Ppt.Slide slide, StringBuilder sb, int maxChars)
        {
            if (slide == null || sb == null || maxChars <= 0) return;
            int startLength = sb.Length;
            foreach (Ppt.Shape shape in slide.Shapes)
            {
                if (sb.Length - startLength >= maxChars) break;
                try
                {
                    if (shape.HasTextFrame != Office.MsoTriState.msoTrue ||
                        shape.TextFrame.HasText != Office.MsoTriState.msoTrue)
                        continue;

                    int remaining = Math.Max(1, maxChars - (sb.Length - startLength));
                    var range = shape.TextFrame.TextRange;
                    int length = 0;
                    try { length = range.Length; } catch { }
                    string text;
                    if (length > 0)
                        text = range.Characters(1, Math.Min(length, remaining)).Text ?? "";
                    else
                        text = TextUtil.Truncate(range.Text ?? "", remaining);
                    sb.AppendLine("• [" + shape.Name + "] " + text);
                }
                catch { }
            }
            if (sb.Length - startLength > maxChars)
                sb.Length = startLength + maxChars;
        }
    }
}
