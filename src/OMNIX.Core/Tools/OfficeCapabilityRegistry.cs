using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OMNIX.Core.Context;

namespace OMNIX.Core.Tools
{
    public sealed class OfficeCapabilityDescriptor
    {
        public HostType Host { get; set; }
        public string Operation { get; set; }
        public string Category { get; set; }
        public bool Write { get; set; }
        public string Summary { get; set; }
        public string Arguments { get; set; }
    }

    /// <summary>
    /// Truthful catalog of the broad Office surface OMNIX exposes to models.
    /// This is intentionally explicit rather than reflective: unknown operations fail closed.
    /// Security/Trust Center, VBA/macro execution, arbitrary file/process access and modal
    /// dialog automation are not part of this registry.
    /// </summary>
    public static class OfficeCapabilityRegistry
    {
        private static readonly OfficeCapabilityDescriptor[] Items = new[]
        {
            // Excel inspection
            R(HostType.Excel,"workbook.summary","Workbook","Workbook/sheet/name/table/chart/pivot/connection counts and protection state.",""),
            R(HostType.Excel,"sheet.inspect","Worksheet","Inspect one worksheet: used range, visibility, dimensions, tables, charts, pivots, shapes and frozen panes.","sheet"),
            R(HostType.Excel,"range.inspect","Range","Inspect values, formulas, number formats, validation, comments/notes and formatting for a bounded range.","sheet,address"),
            R(HostType.Excel,"table.inspect","Table","Inspect ListObject columns, rows, totals, style, filters and source range.","sheet,table"),
            R(HostType.Excel,"name.inspect","Names","Inspect workbook/worksheet defined names and formulas.","query"),
            R(HostType.Excel,"chart.inspect","Charts","Inspect chart type, title, series, source and placement metadata.","sheet,chart"),
            R(HostType.Excel,"pivot.inspect","Pivot","Inspect PivotTable fields, filters, row/column/data fields and cache metadata.","sheet,pivot"),
            R(HostType.Excel,"validation.inspect","Validation","Inspect data validation on a bounded range.","sheet,address"),
            R(HostType.Excel,"conditional_format.inspect","Conditional formatting","Inspect conditional-format rules on a bounded range.","sheet,address"),
            R(HostType.Excel,"filter.inspect","Sort/Filter","Inspect AutoFilter and visible filter state for a worksheet/table.","sheet,table?"),
            R(HostType.Excel,"page_setup.inspect","Page layout","Inspect page orientation, margins, print area, titles and scaling.","sheet"),
            R(HostType.Excel,"view.inspect","View","Inspect active sheet/window selection, zoom, freeze panes and split state.",""),
            // Excel write
            W(HostType.Excel,"sheet.create","Worksheet","Create a new worksheet with a unique or explicit name.","name,position?,uniqueName?"),
            W(HostType.Excel,"sheet.rename","Worksheet","Rename a worksheet without overwriting another sheet.","sheet,newName"),
            W(HostType.Excel,"sheet.delete","Worksheet","Delete one explicitly named worksheet after confirmation.","sheet"),
            W(HostType.Excel,"sheet.copy","Worksheet","Copy a worksheet inside the active workbook.","sheet,newName?"),
            W(HostType.Excel,"sheet.move","Worksheet","Move a worksheet to a specific workbook position.","sheet,index"),
            W(HostType.Excel,"sheet.visibility","Worksheet","Show, hide or very-hide a worksheet.","sheet,state=visible|hidden|veryHidden"),
            W(HostType.Excel,"range.clear","Range","Clear contents, formats, comments/notes or all from a bounded range.","sheet,address,what=contents|formats|comments|all"),
            W(HostType.Excel,"range.insert","Range","Insert cells/rows/columns with an explicit shift direction.","sheet,address,mode=cellsDown|cellsRight|rows|columns"),
            W(HostType.Excel,"range.delete","Range","Delete bounded cells/rows/columns with an explicit shift direction.","sheet,address,mode=cellsUp|cellsLeft|rows|columns"),
            W(HostType.Excel,"range.merge","Range","Merge or unmerge a bounded range.","sheet,address,merge=true|false"),
            W(HostType.Excel,"range.autofill","Range","Fill a destination from a source pattern/formula using Excel AutoFill.","sheet,source,destination,type?"),
            W(HostType.Excel,"row.height","Rows/Columns","Set row height or AutoFit bounded rows.","sheet,address,height|autofit"),
            W(HostType.Excel,"column.width","Rows/Columns","Set column width or AutoFit bounded columns.","sheet,address,width|autofit"),
            W(HostType.Excel,"row.visibility","Rows/Columns","Hide/show bounded rows.","sheet,address,hidden"),
            W(HostType.Excel,"column.visibility","Rows/Columns","Hide/show bounded columns.","sheet,address,hidden"),
            W(HostType.Excel,"table.create","Table","Convert a bounded range into a named Excel table.","sheet,address,name?,style?"),
            W(HostType.Excel,"table.resize","Table","Resize an existing Excel table to an explicit range.","sheet,table,address"),
            W(HostType.Excel,"table.rename","Table","Rename an Excel table.","sheet,table,newName"),
            W(HostType.Excel,"table.style","Table","Apply a built-in table style and totals/header options.","sheet,table,style?,showTotals?,showHeaders?,bandedRows?"),
            W(HostType.Excel,"table.delete","Table","Remove a table; optionally keep its cells as a normal range.","sheet,table,keepData=true|false"),
            W(HostType.Excel,"sort.apply","Sort/Filter","Sort a bounded range or table by one explicit key.","sheet,address,key,order=asc|desc,header=true|false"),
            W(HostType.Excel,"filter.apply","Sort/Filter","Apply one AutoFilter criterion to a bounded range/table.","sheet,address,field,criteria"),
            W(HostType.Excel,"filter.clear","Sort/Filter","Clear worksheet AutoFilter filters.","sheet"),
            W(HostType.Excel,"validation.add","Validation","Add list, whole-number, decimal, date or text-length validation to a bounded range.","sheet,address,type,operator?,formula1,formula2?,allowBlank?"),
            W(HostType.Excel,"validation.delete","Validation","Remove data validation from a bounded range.","sheet,address"),
            W(HostType.Excel,"conditional_format.add","Conditional formatting","Add cell-value or formula conditional formatting with an explicit fill/font color.","sheet,address,type=cellValue|formula,operator?,formula1,formula2?,fillColor?,fontColor?"),
            W(HostType.Excel,"conditional_format.clear","Conditional formatting","Clear conditional-format rules from a bounded range.","sheet,address"),
            W(HostType.Excel,"name.add","Names","Create/update a workbook defined name referring to an explicit formula/range.","name,refersTo"),
            W(HostType.Excel,"name.delete","Names","Delete one workbook defined name.","name"),
            W(HostType.Excel,"chart.create","Charts","Create a chart from an explicit source range and place it on a worksheet.","sheet,address,chartType,title?,left?,top?,width?,height?"),
            W(HostType.Excel,"chart.delete","Charts","Delete one named chart object.","sheet,chart"),
            W(HostType.Excel,"chart.title","Charts","Set or clear chart title text.","sheet,chart,title"),
            W(HostType.Excel,"pivot.refresh","Pivot","Refresh one PivotTable and its cache.","sheet,pivot"),
            W(HostType.Excel,"comment.set","Comments/Notes","Create or replace a legacy cell note/comment on one cell.","sheet,address,text"),
            W(HostType.Excel,"comment.delete","Comments/Notes","Delete a cell note/comment from one cell.","sheet,address"),
            W(HostType.Excel,"hyperlink.add","Links","Add a hyperlink to a single cell.","sheet,address,url,text?"),
            W(HostType.Excel,"view.freeze","View","Freeze panes at the top-left boundary of an explicit cell, or unfreeze.","sheet,address?,freeze"),
            W(HostType.Excel,"view.zoom","View","Set active Excel window zoom.","zoom"),
            W(HostType.Excel,"page_setup.set","Page layout","Set selected page-layout properties without opening dialogs.","sheet,orientation?,fitToPagesWide?,fitToPagesTall?,printArea?"),

            // Word inspection
            R(HostType.Word,"document.summary","Document","Inspect sections, pages, paragraphs, tables, styles, fields, bookmarks, comments, revisions, controls and shapes.",""),
            R(HostType.Word,"selection.inspect","Selection","Inspect exact current selection text, style, paragraph and table context.",""),
            R(HostType.Word,"paragraph.inspect","Paragraph","Inspect a bounded paragraph by index.","index"),
            R(HostType.Word,"table.inspect","Tables","Inspect table dimensions and bounded cell text.","table"),
            R(HostType.Word,"styles.inspect","Styles","Search document styles and their type/in-use state.","query?"),
            R(HostType.Word,"bookmarks.inspect","Bookmarks","List bookmarks and their ranges.","query?"),
            R(HostType.Word,"fields.inspect","Fields","List bounded fields and field codes/results.","offset?"),
            R(HostType.Word,"comments.inspect","Review","List bounded comments and referenced ranges.","offset?"),
            R(HostType.Word,"revisions.inspect","Review","Inspect Track Changes state and bounded revisions.","offset?"),
            R(HostType.Word,"content_controls.inspect","Content controls","Inspect content controls, titles, tags and types.","offset?"),
            R(HostType.Word,"sections.inspect","Sections/Page setup","Inspect section page setup, headers/footers and orientation.","offset?"),
            // Word write
            W(HostType.Word,"text.insert","Text","Insert text before/after current selection or at a character position.","text,where=beforeSelection|afterSelection|position,position?"),
            W(HostType.Word,"text.replace_range","Text","Replace an explicit main-story character range.","start,end,text"),
            W(HostType.Word,"find_replace","Text","Find/replace bounded document text using Word Find options.","find,replace,matchCase?,wholeWord?,replaceAll?"),
            W(HostType.Word,"paragraph.format","Paragraph","Format current selection paragraphs: alignment, spacing, indents and keep options.","alignment?,spaceBefore?,spaceAfter?,lineSpacing?,leftIndent?,rightIndent?,firstLineIndent?,keepWithNext?"),
            W(HostType.Word,"font.format","Font","Format current selection font: name, size, bold, italic, underline, color.","fontName?,fontSize?,bold?,italic?,underline?,fontColor?"),
            W(HostType.Word,"style.apply","Styles","Apply an existing Word style to current selection.","style"),
            W(HostType.Word,"table.create","Tables","Create a Word table at current selection.","rows,columns,style?"),
            W(HostType.Word,"table.add_row","Tables","Add a row to a table.","table,position?"),
            W(HostType.Word,"table.add_column","Tables","Add a column to a table.","table"),
            W(HostType.Word,"table.delete_row","Tables","Delete one row from a table.","table,row"),
            W(HostType.Word,"table.delete_column","Tables","Delete one column from a table.","table,column"),
            W(HostType.Word,"table.style","Tables","Apply a built-in table style and AutoFit behavior.","table,style?,autofit=content|window|fixed?"),
            W(HostType.Word,"bookmark.add","Bookmarks","Create/replace a bookmark around the current selection.","name"),
            W(HostType.Word,"bookmark.delete","Bookmarks","Delete one named bookmark.","name"),
            W(HostType.Word,"comment.add","Review","Add a comment to the current selection.","text"),
            W(HostType.Word,"comment.delete","Review","Delete one comment by index.","comment"),
            W(HostType.Word,"track_changes","Review","Turn Track Changes on or off.","enabled"),
            W(HostType.Word,"revisions.accept","Review","Accept revisions in current selection or document.","scope=selection|document"),
            W(HostType.Word,"revisions.reject","Review","Reject revisions in current selection or document.","scope=selection|document"),
            W(HostType.Word,"header_footer.set","Headers/Footers","Set primary header or footer text in one section.","section,type=header|footer,text"),
            W(HostType.Word,"page_setup.set","Sections/Page setup","Set orientation and margins for one section.","section,orientation?,topMargin?,bottomMargin?,leftMargin?,rightMargin?"),
            W(HostType.Word,"field.insert","Fields","Insert a safe Word field at the selection from an allowlisted field type.","type=page|numPages|date|time|fileName|author"),
            W(HostType.Word,"hyperlink.add","Links","Add a hyperlink around selection or insert linked display text.","url,text?"),
            W(HostType.Word,"content_control.add","Content controls","Add a plain/rich text/check-box content control at selection.","type=plainText|richText|checkBox,title?,tag?"),

            // PowerPoint inspection
            R(HostType.PowerPoint,"presentation.summary","Presentation","Inspect slides, sections, masters, layouts, themes and media/shape counts.",""),
            R(HostType.PowerPoint,"slide.inspect","Slides","Inspect one slide, layout, transition and shape inventory.","slide"),
            R(HostType.PowerPoint,"shape.inspect","Shapes","Inspect one shape: type, bounds, text, fill/line, table/chart/group metadata.","slide,shape"),
            R(HostType.PowerPoint,"table.inspect","Tables","Inspect bounded PowerPoint table cell text.","slide,shape"),
            R(HostType.PowerPoint,"notes.inspect","Notes","Inspect speaker notes for one slide.","slide"),
            R(HostType.PowerPoint,"animations.inspect","Animations","Inspect main animation sequence effects on one slide.","slide"),
            R(HostType.PowerPoint,"sections.inspect","Sections","Inspect presentation sections and slide membership.",""),
            // PowerPoint write
            W(HostType.PowerPoint,"slide.delete","Slides","Delete one slide.","slide"),
            W(HostType.PowerPoint,"slide.duplicate","Slides","Duplicate one slide and optionally move the copy.","slide,index?"),
            W(HostType.PowerPoint,"slide.move","Slides","Move one slide to a 1-based position.","slide,index"),
            W(HostType.PowerPoint,"slide.layout","Slides","Apply an existing custom layout by index.","slide,layout"),
            W(HostType.PowerPoint,"shape.add_textbox","Shapes","Add a text box to a slide with explicit bounds/text.","slide,left,top,width,height,text"),
            W(HostType.PowerPoint,"shape.add_shape","Shapes","Add a basic rectangle/roundedRectangle/ellipse/line with explicit bounds.","slide,type,left,top,width,height"),
            W(HostType.PowerPoint,"shape.delete","Shapes","Delete one shape.","slide,shape"),
            W(HostType.PowerPoint,"shape.move_resize","Shapes","Set one shape's position and size.","slide,shape,left?,top?,width?,height?"),
            W(HostType.PowerPoint,"shape.rotate","Shapes","Set one shape rotation in degrees.","slide,shape,rotation"),
            W(HostType.PowerPoint,"shape.fill","Shapes","Set solid fill color/transparency or no fill.","slide,shape,color?,transparency?,none?"),
            W(HostType.PowerPoint,"shape.line","Shapes","Set line color/weight or hide line.","slide,shape,color?,weight?,visible?"),
            W(HostType.PowerPoint,"text.set","Text","Set text of one text-frame shape.","slide,shape,text"),
            W(HostType.PowerPoint,"text.format","Text","Format shape text font/alignment.","slide,shape,fontName?,fontSize?,bold?,italic?,fontColor?,alignment?"),
            W(HostType.PowerPoint,"table.create","Tables","Create a table shape with explicit bounds.","slide,rows,columns,left,top,width,height"),
            W(HostType.PowerPoint,"table.set_cell","Tables","Set text in one table cell.","slide,shape,row,column,text"),
            W(HostType.PowerPoint,"notes.set","Notes","Replace speaker notes for one slide.","slide,text"),
            W(HostType.PowerPoint,"hyperlink.add","Links","Add a URL action to one shape.","slide,shape,url"),
            W(HostType.PowerPoint,"transition.set","Transitions","Set a safe transition entry effect and advance options.","slide,type=none|fade|push|wipe,advanceOnClick?,advanceAfterSeconds?"),
            W(HostType.PowerPoint,"animation.add","Animations","Add a basic entrance animation to one shape.","slide,shape,type=appear|fade|fly"),
            W(HostType.PowerPoint,"animation.clear","Animations","Remove main-sequence animation effects from one slide.","slide"),
            W(HostType.PowerPoint,"view.goto_slide","View","Navigate to one slide without changing content.","slide")
        };

        public static string Search(HostType host, string query, int offset)
        {
            query = (query ?? "").Trim();
            var matches = Items.Where(x => x.Host == host &&
                (query.Length == 0 ||
                 x.Operation.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                 x.Category.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                 x.Summary.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0))
                .OrderBy(x => x.Category, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.Operation, StringComparer.OrdinalIgnoreCase)
                .ToList();

            offset = Math.Max(0, offset);
            const int pageSize = 40;
            var page = matches.Skip(offset).Take(pageSize).ToList();
            var sb = new StringBuilder();
            sb.AppendLine(host + " capability registry: " + matches.Count + " matching operations; offset=" + offset +
                          "; nextOffset=" + (offset + pageSize < matches.Count ? (offset + pageSize).ToString() : "none"));
            sb.AppendLine("Use inspect_office_capability for read operations and apply_office_capability for write operations. Unknown operations fail closed.");
            foreach (var item in page)
            {
                sb.AppendLine((item.Write ? "WRITE" : "READ") + " | " + item.Operation + " | " + item.Category +
                              " | " + item.Summary + (string.IsNullOrEmpty(item.Arguments) ? "" : " | args: " + item.Arguments));
            }
            return sb.ToString();
        }

        public static bool IsKnown(HostType host, string operation, bool write)
        {
            return Items.Any(x => x.Host == host && x.Write == write &&
                string.Equals(x.Operation, operation, StringComparison.OrdinalIgnoreCase));
        }

        private static OfficeCapabilityDescriptor R(HostType host,string op,string category,string summary,string args)
        { return new OfficeCapabilityDescriptor { Host=host,Operation=op,Category=category,Summary=summary,Arguments=args,Write=false }; }
        private static OfficeCapabilityDescriptor W(HostType host,string op,string category,string summary,string args)
        { return new OfficeCapabilityDescriptor { Host=host,Operation=op,Category=category,Summary=summary,Arguments=args,Write=true }; }
    }
}
