using System;
using System.Collections.Generic;
using OMNIX.Core.Tools;
using Office = Microsoft.Office.Core;

namespace OMNIX.Core.Context
{
    /// <summary>
    /// One adapter per Office host (spec Section 3, Layer 3). Implemented per host inside
    /// OMNIX.Core so the thin host projects only wire the Application object in.
    /// </summary>
    /// <summary>Optional bounded navigation; every call stays in the request's document scope.</summary>
    public interface IIndexedHostAdapter
    {
        string ReadDocumentMap(int offset);
        string ReadDocumentSection(ToolArguments arguments);
    }


    /// <summary>
    /// Visible execution is deliberately separate from chat status text. Implementations move the
    /// real Office UI to the object being inspected/changed and may activate the relevant native
    /// Ribbon tab. They must never simulate a click for a command that was not actually invoked.
    /// </summary>
    public enum OfficeExecutionStage
    {
        Inspect,
        Preview,
        Apply,
        Verify
    }

    public interface IVisibleOfficeExecutionHost
    {
        /// <summary>Called by the host Ribbon once Office has supplied its real IRibbonUI instance.</summary>
        void BindRibbon(Office.IRibbonUI ribbonUi);

        /// <summary>
        /// Bring the actual target into view: workbook/sheet/range, Word range, or slide/shape.
        /// This is presentation of the real operation, not an animation in the OMNIX chat pane.
        /// </summary>
        void RevealOperation(string toolName, ToolArguments arguments, OfficeExecutionStage stage);

        /// <summary>Truthful summary of executable host capabilities exposed to the current model.</summary>
        string CapabilitySummary { get; }
    }

    public interface IHostAdapter
    {
        HostType Host { get; }
        string HostDisplayName { get; }

        /// <summary>Current document state (called on the Office UI thread).</summary>
        OfficeContext ReadContext();

        /// <summary>Raw text of the current selection (read tool: read_selection).</summary>
        string ReadSelection();

        /// <summary>Whole document/workbook/presentation text, capped (read tool).</summary>
        string ReadDocument(int maxChars);

        /// <summary>Excel only: PNG bytes of a chart (by name, or the active chart).</summary>
        byte[] CaptureChartAsImage(string chartName);

        /// <summary>PowerPoint only: PNG bytes of a slide (by index, or the current one).</summary>
        byte[] CaptureSlideAsImage(int slideIndexOneBased);

        /// <summary>Current slide/document as an image when the selection itself is visual (attach-from-document).</summary>
        byte[] CaptureCurrentViewAsImage();

        // ---- write tools (spec Layer 7 whitelist; always previewed + user-confirmed) ----
        WritePreview PrepareWrite(string toolName, string argumentsJson);
        void ApplyWrite(string toolName, string argumentsJson);
    }
}
