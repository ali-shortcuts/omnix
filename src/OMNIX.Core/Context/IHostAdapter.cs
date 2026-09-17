using System;
using System.Collections.Generic;
using OMNIX.Core.Tools;

namespace OMNIX.Core.Context
{
    /// <summary>
    /// One adapter per Office host (spec Section 3, Layer 3). Implemented per host inside
    /// OMNIX.Core so the thin host projects only wire the Application object in.
    /// </summary>
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

    /// <summary>
    /// Optional targeted retrieval capability. Search-aware decorators implement this interface
    /// without changing the existing host adapter mutation/read contracts. Implementations must use
    /// bounded/native host search where available and must not materialize an entire Office file just
    /// to answer one query.
    /// </summary>
    public interface IDocumentSearchProvider
    {
        string SearchDocument(string query, int maxResults, int maxChars);
    }
}
