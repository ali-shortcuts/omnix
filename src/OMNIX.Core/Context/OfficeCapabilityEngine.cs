using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using OMNIX.Core.Errors;
using OMNIX.Core.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OMNIX.Core.Context
{
    public sealed class OfficeCapabilityDescriptor
    {
        public HostType Host { get; set; }
        public string Id { get; set; }
        public string Category { get; set; }
        public string Description { get; set; }
        public string Arguments { get; set; }
    }

    /// <summary>
    /// The model sees a truthful, queryable capability catalog instead of an "execute anything"
    /// escape hatch. Every listed capability is implemented through the Office Object Model and
    /// every mutation still crosses the normal OMNIX preview/confirmation boundary.
    /// </summary>
    public static class OfficeCapabilityRegistry
    {
        private static readonly List<OfficeCapabilityDescriptor> Items = Build();

        public static IReadOnlyList<OfficeCapabilityDescriptor> ForHost(HostType host)
        {
            return Items.Where(x => x.Host == host).ToList();
        }

        public static bool Exists(HostType host, string id)
        {
            return Items.Any(x => x.Host == host && string.Equals(x.Id, id, StringComparison.OrdinalIgnoreCase));
        }

        public static string Search(HostType host, string query, int offset)
        {
            query = (query ?? "").Trim();
            var matches = Items.Where(x => x.Host == host &&
                (query.Length == 0 ||
                 x.Id.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                 x.Category.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                 x.Description.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0))
                .OrderBy(x => x.Category, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.Id, StringComparer.OrdinalIgnoreCase)
                .ToList();

            offset = Math.Max(0, offset);
            int take = Math.Min(50, Math.Max(0, matches.Count - offset));
            var sb = new StringBuilder();
            sb.AppendLine("Host=" + host + "; implementedCapabilities=" + Items.Count(x => x.Host == host) +
                          "; matches=" + matches.Count + "; offset=" + offset);
            foreach (var item in matches.Skip(offset).Take(take))
                sb.AppendLine(item.Id + " | " + item.Category + " | " + item.Description +
                              " | args: " + item.Arguments);
            sb.AppendLine("nextOffset=" + (offset + take < matches.Count ? (offset + take).ToString() : "none"));
            sb.AppendLine("Only capabilities listed here are executable. Trust Center, macro/VBA execution, arbitrary shell/files, and unrestricted ExecuteMso are intentionally not exposed.");
            return sb.ToString();
        }

        private static OfficeCapabilityDescriptor C(HostType host, string id, string category, string description, string args)
        {
            return new OfficeCapabilityDescriptor { Host = host, Id = id, Category = category, Description = description, Arguments = args };
        }

        private static List<OfficeCapabilityDescriptor> Build()
        {
            var x = new List<OfficeCapabilityDescriptor>();

            // Excel — workbook/worksheet/range/data/layout/analysis/presentation surface.
            x.Add(C(HostType.Excel,"worksheet.add","Worksheet","Add a new worksheet without overwriting an existing one.","name"));
            x.Add(C(HostType.Excel,"worksheet.rename","Worksheet","Rename an existing worksheet.","sheet,newName"));
            x.Add(C(HostType.Excel,"worksheet.move","Worksheet","Move a worksheet to a one-based position.","sheet,index"));
            x.Add(C(HostType.Excel,"worksheet.visibility","Worksheet","Show, hide, or very-hide a worksheet.","sheet,state=visible|hidden|veryHidden"));
            x.Add(C(HostType.Excel,"worksheet.tab_color","Worksheet","Set worksheet tab color.","sheet,color=#RRGGBB"));
            x.Add(C(HostType.Excel,"range.clear_contents","Range","Clear values/formulas but keep formatting.","sheet,address"));
            x.Add(C(HostType.Excel,"range.clear_formats","Range","Clear formatting but keep values/formulas.","sheet,address"));
            x.Add(C(HostType.Excel,"range.insert_rows","Range","Insert entire rows at the target.","sheet,address"));
            x.Add(C(HostType.Excel,"range.delete_rows","Range","Delete entire rows at the target.","sheet,address"));
            x.Add(C(HostType.Excel,"range.insert_columns","Range","Insert entire columns at the target.","sheet,address"));
            x.Add(C(HostType.Excel,"range.delete_columns","Range","Delete entire columns at the target.","sheet,address"));
            x.Add(C(HostType.Excel,"range.merge","Range","Merge a bounded contiguous range.","sheet,address"));
            x.Add(C(HostType.Excel,"range.unmerge","Range","Unmerge a bounded range.","sheet,address"));
            x.Add(C(HostType.Excel,"range.autofit","Layout","AutoFit rows and/or columns.","sheet,address,rows=true|false,columns=true|false"));
            x.Add(C(HostType.Excel,"range.column_width","Layout","Set column width for target columns.","sheet,address,width"));
            x.Add(C(HostType.Excel,"range.row_height","Layout","Set row height for target rows.","sheet,address,height"));
            x.Add(C(HostType.Excel,"range.hide_rows","Layout","Hide or unhide target rows.","sheet,address,hidden=true|false"));
            x.Add(C(HostType.Excel,"range.hide_columns","Layout","Hide or unhide target columns.","sheet,address,hidden=true|false"));
            x.Add(C(HostType.Excel,"range.find_replace","Editing","Find/replace within a bounded range.","sheet,address,find,replace,matchCase=false"));
            x.Add(C(HostType.Excel,"range.remove_duplicates","Data","Remove duplicate rows using one-based column indices relative to the range.","sheet,address,columns=[1,2],hasHeader=true"));
            x.Add(C(HostType.Excel,"validation.list","Data Validation","Apply list validation to a range.","sheet,address,source"));
            x.Add(C(HostType.Excel,"validation.clear","Data Validation","Remove data validation from a range.","sheet,address"));
            x.Add(C(HostType.Excel,"filter.apply","Filter","Apply AutoFilter criteria to one field of a range/table.","sheet,address,field,criteria"));
            x.Add(C(HostType.Excel,"filter.clear","Filter","Clear active worksheet filters.","sheet"));
            x.Add(C(HostType.Excel,"sort.range","Sort","Sort a rectangular range by a key column.","sheet,address,keyColumn,descending=false,hasHeader=true"));
            x.Add(C(HostType.Excel,"window.freeze_panes","View","Freeze panes above/left of the target cell.","sheet,address"));
            x.Add(C(HostType.Excel,"window.unfreeze_panes","View","Remove frozen panes.","sheet"));
            x.Add(C(HostType.Excel,"name.define","Names","Create/update a workbook-level named range.","name,sheet,address"));
            x.Add(C(HostType.Excel,"name.delete","Names","Delete a workbook-level defined name.","name"));
            x.Add(C(HostType.Excel,"hyperlink.add","Links","Add a hyperlink to a cell/range.","sheet,address,url,text"));
            x.Add(C(HostType.Excel,"note.add","Notes","Add or replace a legacy cell note/comment.","sheet,address,text"));
            x.Add(C(HostType.Excel,"note.delete","Notes","Delete a legacy cell note/comment.","sheet,address"));
            x.Add(C(HostType.Excel,"table.create","Table","Create an Excel ListObject from an existing range.","sheet,address,name,style=TableStyleMedium2"));
            x.Add(C(HostType.Excel,"table.delete","Table","Unlist or delete a named table.","sheet,name,deleteData=false"));
            x.Add(C(HostType.Excel,"chart.create","Chart","Create a chart from a source range.","sheet,address,name,type=column|bar|line|pie,title"));
            x.Add(C(HostType.Excel,"chart.delete","Chart","Delete a named embedded chart.","sheet,name"));
            x.Add(C(HostType.Excel,"chart.title","Chart","Set chart title text.","sheet,name,title"));
            x.Add(C(HostType.Excel,"page.orientation","Page Layout","Set portrait or landscape orientation.","sheet,orientation=portrait|landscape"));
            x.Add(C(HostType.Excel,"page.fit","Page Layout","Fit worksheet printout to pages wide/tall.","sheet,pagesWide,pagesTall"));
            x.Add(C(HostType.Excel,"print_area.set","Page Layout","Set worksheet print area.","sheet,address"));
            x.Add(C(HostType.Excel,"print_area.clear","Page Layout","Clear worksheet print area.","sheet"));

            // Word — content, formatting, document structure, review and layout.
            x.Add(C(HostType.Word,"text.insert_before","Editing","Insert text before current selection.","text"));
            x.Add(C(HostType.Word,"text.insert_after","Editing","Insert text after current selection.","text"));
            x.Add(C(HostType.Word,"text.find_replace","Editing","Find/replace within the main document story.","find,replace,matchCase=false,wholeWord=false"));
            x.Add(C(HostType.Word,"selection.font","Formatting","Format current selection font.","name,size,bold,italic,underline,color=#RRGGBB"));
            x.Add(C(HostType.Word,"selection.paragraph","Formatting","Format current selection paragraphs.","alignment=left|center|right|justify,spaceBefore,spaceAfter,lineSpacing"));
            x.Add(C(HostType.Word,"selection.style","Styles","Apply a named Word style to current selection.","style"));
            x.Add(C(HostType.Word,"selection.bullets","Lists","Apply/remove bullets on selected paragraphs.","enabled=true|false"));
            x.Add(C(HostType.Word,"selection.numbering","Lists","Apply/remove numbering on selected paragraphs.","enabled=true|false"));
            x.Add(C(HostType.Word,"table.insert","Tables","Insert a table at current selection.","rows,columns"));
            x.Add(C(HostType.Word,"table.add_row","Tables","Add a row to a selected/containing table.","position=after|before"));
            x.Add(C(HostType.Word,"table.add_column","Tables","Add a column to a selected/containing table.","position=after|before"));
            x.Add(C(HostType.Word,"table.autofit","Tables","AutoFit selected/containing table.","mode=content|window|fixed"));
            x.Add(C(HostType.Word,"break.page","Document Structure","Insert a page break at current selection.",""));
            x.Add(C(HostType.Word,"break.section","Document Structure","Insert a section break.","type=nextPage|continuous|evenPage|oddPage"));
            x.Add(C(HostType.Word,"header.set","Headers/Footers","Set primary header text for current section.","text"));
            x.Add(C(HostType.Word,"footer.set","Headers/Footers","Set primary footer text for current section.","text"));
            x.Add(C(HostType.Word,"bookmark.add","Bookmarks","Create/replace a bookmark at current selection.","name"));
            x.Add(C(HostType.Word,"bookmark.delete","Bookmarks","Delete a bookmark.","name"));
            x.Add(C(HostType.Word,"comment.add","Review","Add a comment to current selection.","text"));
            x.Add(C(HostType.Word,"hyperlink.add","Links","Add hyperlink to current selection.","url,text"));
            x.Add(C(HostType.Word,"review.track_changes","Review","Enable or disable Track Changes.","enabled=true|false"));
            x.Add(C(HostType.Word,"review.accept_all","Review","Accept all revisions in active document.",""));
            x.Add(C(HostType.Word,"review.reject_all","Review","Reject all revisions in active document.",""));
            x.Add(C(HostType.Word,"page.orientation","Page Layout","Set current section orientation.","orientation=portrait|landscape"));
            x.Add(C(HostType.Word,"page.margins","Page Layout","Set current section margins in points.","top,bottom,left,right"));
            x.Add(C(HostType.Word,"page.columns","Page Layout","Set text column count in current section.","count"));
            x.Add(C(HostType.Word,"field.insert","Fields","Insert a Word field at current selection.","code"));
            x.Add(C(HostType.Word,"footnote.add","References","Add a footnote at current selection.","text"));
            x.Add(C(HostType.Word,"endnote.add","References","Add an endnote at current selection.","text"));

            // PowerPoint — slides, shapes, text, tables, alignment, notes, animation.
            x.Add(C(HostType.PowerPoint,"slide.add","Slides","Add a slide at a one-based position.","index,layout=blank|title|text"));
            x.Add(C(HostType.PowerPoint,"slide.delete","Slides","Delete a slide.","slide"));
            x.Add(C(HostType.PowerPoint,"slide.duplicate","Slides","Duplicate a slide.","slide"));
            x.Add(C(HostType.PowerPoint,"slide.move","Slides","Move a slide to a one-based position.","slide,index"));
            x.Add(C(HostType.PowerPoint,"slide.background","Design","Set solid slide background color.","slide,color=#RRGGBB"));
            x.Add(C(HostType.PowerPoint,"shape.textbox","Shapes","Add a text box.","slide,left,top,width,height,text"));
            x.Add(C(HostType.PowerPoint,"shape.add","Shapes","Add rectangle, ellipse, line, roundedRectangle or arrow.","slide,type,left,top,width,height"));
            x.Add(C(HostType.PowerPoint,"shape.delete","Shapes","Delete a shape by one-based index or name.","slide,shape"));
            x.Add(C(HostType.PowerPoint,"shape.text","Text","Set shape text.","slide,shape,text"));
            x.Add(C(HostType.PowerPoint,"shape.position","Shapes","Set shape position and size.","slide,shape,left,top,width,height"));
            x.Add(C(HostType.PowerPoint,"shape.format","Shapes","Format fill/line/text for a shape.","slide,shape,fillColor,lineColor,fontColor,fontSize,bold"));
            x.Add(C(HostType.PowerPoint,"shape.zorder","Shapes","Move shape front/back.","slide,shape,position=front|back|forward|backward"));
            x.Add(C(HostType.PowerPoint,"shape.align","Arrange","Align selected shape against slide.","slide,shape,alignment=left|center|right|top|middle|bottom"));
            x.Add(C(HostType.PowerPoint,"table.add","Tables","Add a table to a slide.","slide,rows,columns,left,top,width,height"));
            x.Add(C(HostType.PowerPoint,"table.cell_text","Tables","Set table cell text.","slide,shape,row,column,text"));
            x.Add(C(HostType.PowerPoint,"notes.set","Notes","Set speaker notes body.","slide,text"));
            x.Add(C(HostType.PowerPoint,"hyperlink.add","Links","Add hyperlink to a shape.","slide,shape,url"));
            x.Add(C(HostType.PowerPoint,"animation.fade","Animations","Add a fade entrance animation to a shape.","slide,shape"));
            x.Add(C(HostType.PowerPoint,"transition.fade","Transitions","Apply a fade transition to a slide.","slide"));
            return x;
        }
    }

    internal static class CapabilityArgs
    {
        public static JObject Obj(string json)
        {
            try { return string.IsNullOrWhiteSpace(json) ? new JObject() : JObject.Parse(json); }
            catch { throw new ArgumentException("Capability arguments must be valid JSON."); }
        }

        public static string S(JObject o, string key, string fallback = "")
        {
            var t=o[key]; return t==null ? fallback : t.ToString();
        }
        public static int I(JObject o,string key,int fallback,int min,int max)
        {
            int v; if(!int.TryParse(S(o,key,fallback.ToString(CultureInfo.InvariantCulture)),out v)||v<min||v>max)
                throw new ArgumentException(key+" must be between "+min+" and "+max+"."); return v;
        }
        public static double D(JObject o,string key,double fallback,double min,double max)
        {
            double v; if(!double.TryParse(S(o,key,fallback.ToString(CultureInfo.InvariantCulture)),NumberStyles.Float,CultureInfo.InvariantCulture,out v)||v<min||v>max)
                throw new ArgumentException(key+" must be between "+min+" and "+max+"."); return v;
        }
        public static bool B(JObject o,string key,bool fallback=false)
        {
            var t=o[key]; if(t==null) return fallback; bool v; if(bool.TryParse(t.ToString(),out v)) return v;
            throw new ArgumentException(key+" must be true or false.");
        }
        public static int ColorOle(string raw)
        {
            if(string.IsNullOrWhiteSpace(raw)||raw.Length!=7||raw[0]!='#'||!raw.Substring(1).All(Uri.IsHexDigit))
                throw new ArgumentException("Color must be #RRGGBB.");
            return System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml(raw));
        }
        public static string Capability(JObject root) { var id=S(root,"capability","").Trim(); if(id.Length==0) throw new ArgumentException("Missing capability."); return id; }
        public static JObject Args(JObject root) { return root["args"] as JObject ?? new JObject(); }
    }

    public static class ExcelCapabilityEngine
    {
        public static WritePreview Prepare(Excel.Application app,string json)
        {
            var root=CapabilityArgs.Obj(json); string id=CapabilityArgs.Capability(root); var a=CapabilityArgs.Args(root);
            if(!OfficeCapabilityRegistry.Exists(HostType.Excel,id)) throw new ArgumentException("Unsupported Excel capability: "+id);
            var wb=app.ActiveWorkbook; if(wb==null) throw new InvalidOperationException("No active workbook.");
            string target=CapabilityArgs.S(a,"sheet", app.ActiveSheet is Excel.Worksheet ? ((Excel.Worksheet)app.ActiveSheet).Name : "");
            return new WritePreview{ToolName=ToolNames.ExecuteOfficeCapability,Title="Excel capability — "+id,Before="Current workbook: "+wb.Name+(target.Length>0?"; target sheet: "+target:""),After="Execute "+id+" with validated Office Object Model arguments.",ArgumentsJson=json};
        }

        public static void Apply(Excel.Application app,string json)
        {
            var root=CapabilityArgs.Obj(json); string id=CapabilityArgs.Capability(root); var a=CapabilityArgs.Args(root);
            var wb=app.ActiveWorkbook; if(wb==null) throw new InvalidOperationException("No active workbook.");
            Excel.Worksheet ws=null;
            Func<Excel.Worksheet> sheet=()=> {
                if(ws!=null) return ws; string n=CapabilityArgs.S(a,"sheet","");
                ws=string.IsNullOrWhiteSpace(n)?app.ActiveSheet as Excel.Worksheet:wb.Worksheets[n] as Excel.Worksheet;
                if(ws==null) throw new InvalidOperationException("Worksheet not found."); return ws; };
            Func<Excel.Range> range=()=> { string ad=CapabilityArgs.S(a,"address",""); if(ad.Length==0) throw new ArgumentException("Missing address."); var r=sheet().Range[ad]; if(r.Areas.Count!=1) throw new ArgumentException("Use one contiguous range."); if(Convert.ToInt64(r.Cells.CountLarge)>10000) throw new ArgumentException("Range exceeds 10,000-cell capability limit."); return r; };

            switch(id)
            {
                case "worksheet.add": {
                    string name=CapabilityArgs.S(a,"name","Sheet").Trim(); if(name.Length<1||name.Length>31||name.IndexOfAny(new[]{':','\\','/','?','*','[',']'})>=0) throw new ArgumentException("Invalid worksheet name.");
                    foreach(Excel.Worksheet s in wb.Worksheets) if(string.Equals(s.Name,name,StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Worksheet already exists.");
                    var n=(Excel.Worksheet)wb.Worksheets.Add(After:wb.Sheets[wb.Sheets.Count]); n.Name=name; n.Activate(); break; }
                case "worksheet.rename": sheet().Name=CapabilityArgs.S(a,"newName",""); break;
                case "worksheet.move": sheet().Move(Before:wb.Sheets[CapabilityArgs.I(a,"index",1,1,wb.Sheets.Count)]); break;
                case "worksheet.visibility": { string state=CapabilityArgs.S(a,"state","visible").ToLowerInvariant(); sheet().Visible=state=="visible"?Excel.XlSheetVisibility.xlSheetVisible:state=="hidden"?Excel.XlSheetVisibility.xlSheetHidden:state=="veryhidden"?Excel.XlSheetVisibility.xlSheetVeryHidden:throw new ArgumentException("state must be visible, hidden, or veryHidden."); break; }
                case "worksheet.tab_color": sheet().Tab.Color=CapabilityArgs.ColorOle(CapabilityArgs.S(a,"color","")); break;
                case "range.clear_contents": range().ClearContents(); break;
                case "range.clear_formats": range().ClearFormats(); break;
                case "range.insert_rows": range().EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown); break;
                case "range.delete_rows": range().EntireRow.Delete(); break;
                case "range.insert_columns": range().EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight); break;
                case "range.delete_columns": range().EntireColumn.Delete(); break;
                case "range.merge": range().Merge(); break;
                case "range.unmerge": range().UnMerge(); break;
                case "range.autofit": { var r=range(); if(CapabilityArgs.B(a,"rows",true)) r.Rows.AutoFit(); if(CapabilityArgs.B(a,"columns",true)) r.Columns.AutoFit(); break; }
                case "range.column_width": range().EntireColumn.ColumnWidth=CapabilityArgs.D(a,"width",10,0.1,255); break;
                case "range.row_height": range().EntireRow.RowHeight=CapabilityArgs.D(a,"height",15,0.1,409); break;
                case "range.hide_rows": range().EntireRow.Hidden=CapabilityArgs.B(a,"hidden",true); break;
                case "range.hide_columns": range().EntireColumn.Hidden=CapabilityArgs.B(a,"hidden",true); break;
                case "range.find_replace": { var r=range(); r.Replace(CapabilityArgs.S(a,"find",""),CapabilityArgs.S(a,"replace",""),Excel.XlLookAt.xlPart,Excel.XlSearchOrder.xlByRows,CapabilityArgs.B(a,"matchCase",false)); break; }
                case "range.remove_duplicates": { var ar=a["columns"] as JArray; if(ar==null||ar.Count==0) throw new ArgumentException("columns is required."); object[] cols=ar.Select(t=>(object)(int)t).ToArray(); range().RemoveDuplicates(cols,CapabilityArgs.B(a,"hasHeader",true)?Excel.XlYesNoGuess.xlYes:Excel.XlYesNoGuess.xlNo); break; }
                case "validation.list": { var r=range(); r.Validation.Delete(); r.Validation.Add(Excel.XlDVType.xlValidateList,Excel.XlDVAlertStyle.xlValidAlertStop,Excel.XlFormatConditionOperator.xlBetween,CapabilityArgs.S(a,"source","")); r.Validation.IgnoreBlank=true; r.Validation.InCellDropdown=true; break; }
                case "validation.clear": range().Validation.Delete(); break;
                case "filter.apply": range().AutoFilter(CapabilityArgs.I(a,"field",1,1,256),CapabilityArgs.S(a,"criteria","")); break;
                case "filter.clear": { var s=sheet(); if(s.FilterMode) s.ShowAllData(); break; }
                case "sort.range": { var r=range(); int key=CapabilityArgs.I(a,"keyColumn",1,1,r.Columns.Count); var keyRange=(Excel.Range)r.Columns[key]; r.Sort(keyRange,CapabilityArgs.B(a,"descending",false)?Excel.XlSortOrder.xlDescending:Excel.XlSortOrder.xlAscending,Type.Missing,Type.Missing,Excel.XlSortOrder.xlAscending,Type.Missing,Excel.XlSortOrder.xlAscending,CapabilityArgs.B(a,"hasHeader",true)?Excel.XlYesNoGuess.xlYes:Excel.XlYesNoGuess.xlNo); break; }
                case "window.freeze_panes": { var r=range(); sheet().Activate(); app.Goto(r,true); app.ActiveWindow.FreezePanes=false; app.ActiveWindow.SplitRow=r.Row-1; app.ActiveWindow.SplitColumn=r.Column-1; app.ActiveWindow.FreezePanes=true; break; }
                case "window.unfreeze_panes": sheet().Activate(); app.ActiveWindow.FreezePanes=false; app.ActiveWindow.SplitRow=0; app.ActiveWindow.SplitColumn=0; break;
                case "name.define": { string n=CapabilityArgs.S(a,"name",""); if(n.Length==0) throw new ArgumentException("Missing name."); wb.Names.Add(n,range()); break; }
                case "name.delete": wb.Names.Item(CapabilityArgs.S(a,"name","")).Delete(); break;
                case "hyperlink.add": { var r=range(); string text=CapabilityArgs.S(a,"text",""); sheet().Hyperlinks.Add(r,CapabilityArgs.S(a,"url",""),Type.Missing,Type.Missing,text.Length>0?text:r.Text); break; }
                case "note.add": { var r=range(); if(Convert.ToInt64(r.Cells.CountLarge)!=1) throw new ArgumentException("note.add requires one cell."); try{if(r.Comment!=null)r.Comment.Delete();}catch{} r.AddComment(CapabilityArgs.S(a,"text","")); break; }
                case "note.delete": { var r=range(); if(Convert.ToInt64(r.Cells.CountLarge)!=1) throw new ArgumentException("note.delete requires one cell."); if(r.Comment!=null)r.Comment.Delete(); break; }
                case "table.create": { var r=range(); string n=CapabilityArgs.S(a,"name",""); var t=sheet().ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,r,Type.Missing,Excel.XlYesNoGuess.xlYes,Type.Missing); if(n.Length>0)t.Name=n; t.TableStyle=CapabilityArgs.S(a,"style","TableStyleMedium2"); break; }
                case "table.delete": { var t=sheet().ListObjects[CapabilityArgs.S(a,"name","")]; if(CapabilityArgs.B(a,"deleteData",false)) t.Delete(); else t.Unlist(); break; }
                case "chart.create": { var r=range(); var chartObjects=(Excel.ChartObjects)sheet().ChartObjects(); double chartLeft=Convert.ToDouble(r.Left)+Convert.ToDouble(r.Width)+20d; double chartTop=Convert.ToDouble(r.Top); var co=(Excel.ChartObject)chartObjects.Add(chartLeft,chartTop,420d,260d); co.Name=CapabilityArgs.S(a,"name","Chart"+chartObjects.Count); co.Chart.SetSourceData(r); string type=CapabilityArgs.S(a,"type","column").ToLowerInvariant(); co.Chart.ChartType=type=="line"?Excel.XlChartType.xlLine:type=="pie"?Excel.XlChartType.xlPie:type=="bar"?Excel.XlChartType.xlBarClustered:Excel.XlChartType.xlColumnClustered; string title=CapabilityArgs.S(a,"title",""); if(title.Length>0){co.Chart.HasTitle=true;co.Chart.ChartTitle.Text=title;} co.Activate(); break; }
                case "chart.delete": ((Excel.ChartObject)sheet().ChartObjects(CapabilityArgs.S(a,"name",""))).Delete(); break;
                case "chart.title": { var ch=((Excel.ChartObject)sheet().ChartObjects(CapabilityArgs.S(a,"name",""))).Chart; ch.HasTitle=true; ch.ChartTitle.Text=CapabilityArgs.S(a,"title",""); break; }
                case "page.orientation": sheet().PageSetup.Orientation=CapabilityArgs.S(a,"orientation","portrait").Equals("landscape",StringComparison.OrdinalIgnoreCase)?Excel.XlPageOrientation.xlLandscape:Excel.XlPageOrientation.xlPortrait; break;
                case "page.fit": { var p=sheet().PageSetup; p.Zoom=false; p.FitToPagesWide=CapabilityArgs.I(a,"pagesWide",1,1,100); p.FitToPagesTall=CapabilityArgs.I(a,"pagesTall",1,1,100); break; }
                case "print_area.set": sheet().PageSetup.PrintArea=range().Address; break;
                case "print_area.clear": sheet().PageSetup.PrintArea=""; break;
                default: throw new ArgumentException("Unsupported Excel capability: "+id);
            }
        }
    }

    public static class WordCapabilityEngine
    {
        public static WritePreview Prepare(Word.Application app,string json)
        {
            var root=CapabilityArgs.Obj(json); string id=CapabilityArgs.Capability(root);
            if(!OfficeCapabilityRegistry.Exists(HostType.Word,id)) throw new ArgumentException("Unsupported Word capability: "+id);
            var doc=app.ActiveDocument; if(doc==null) throw new InvalidOperationException("No active Word document.");
            return new WritePreview{ToolName=ToolNames.ExecuteOfficeCapability,Title="Word capability — "+id,Before="Document: "+doc.Name+"; selection="+app.Selection.Start+"-"+app.Selection.End,After="Execute "+id+" against the active document/selection.",ArgumentsJson=json};
        }
        public static void Apply(Word.Application app,string json)
        {
            var root=CapabilityArgs.Obj(json); string id=CapabilityArgs.Capability(root); var a=CapabilityArgs.Args(root);
            var doc=app.ActiveDocument; if(doc==null) throw new InvalidOperationException("No active Word document."); var sel=app.Selection;
            Func<Word.Table> table=()=> { if(sel.Tables.Count<1) throw new InvalidOperationException("Selection is not inside a table."); return sel.Tables[1]; };
            switch(id)
            {
                case "text.insert_before": sel.Range.InsertBefore(CapabilityArgs.S(a,"text","")); break;
                case "text.insert_after": sel.Range.InsertAfter(CapabilityArgs.S(a,"text","")); break;
                case "text.find_replace": { var f=doc.Content.Find; f.ClearFormatting(); f.Replacement.ClearFormatting(); f.Text=CapabilityArgs.S(a,"find",""); f.Replacement.Text=CapabilityArgs.S(a,"replace",""); f.MatchCase=CapabilityArgs.B(a,"matchCase",false); f.MatchWholeWord=CapabilityArgs.B(a,"wholeWord",false); f.Execute(Replace:Word.WdReplace.wdReplaceAll); break; }
                case "selection.font": { var ft=sel.Font; string n=CapabilityArgs.S(a,"name",""); if(n.Length>0)ft.Name=n; if(a["size"]!=null)ft.Size=(float)CapabilityArgs.D(a,"size",11,6,96); if(a["bold"]!=null)ft.Bold=CapabilityArgs.B(a,"bold")?-1:0; if(a["italic"]!=null)ft.Italic=CapabilityArgs.B(a,"italic")?-1:0; if(a["underline"]!=null)ft.Underline=CapabilityArgs.B(a,"underline")?Word.WdUnderline.wdUnderlineSingle:Word.WdUnderline.wdUnderlineNone; string color=CapabilityArgs.S(a,"color",""); if(color.Length>0)ft.Color=(Word.WdColor)CapabilityArgs.ColorOle(color); break; }
                case "selection.paragraph": { var p=sel.ParagraphFormat; string al=CapabilityArgs.S(a,"alignment",""); if(al=="left")p.Alignment=Word.WdParagraphAlignment.wdAlignParagraphLeft; else if(al=="center")p.Alignment=Word.WdParagraphAlignment.wdAlignParagraphCenter; else if(al=="right")p.Alignment=Word.WdParagraphAlignment.wdAlignParagraphRight; else if(al=="justify")p.Alignment=Word.WdParagraphAlignment.wdAlignParagraphJustify; if(a["spaceBefore"]!=null)p.SpaceBefore=(float)CapabilityArgs.D(a,"spaceBefore",0,0,500); if(a["spaceAfter"]!=null)p.SpaceAfter=(float)CapabilityArgs.D(a,"spaceAfter",0,0,500); if(a["lineSpacing"]!=null)p.LineSpacing=(float)CapabilityArgs.D(a,"lineSpacing",12,1,500); break; }
                case "selection.style": sel.set_Style(CapabilityArgs.S(a,"style","Normal")); break;
                case "selection.bullets": if(CapabilityArgs.B(a,"enabled",true))sel.Range.ListFormat.ApplyBulletDefault();else sel.Range.ListFormat.RemoveNumbers(); break;
                case "selection.numbering": if(CapabilityArgs.B(a,"enabled",true))sel.Range.ListFormat.ApplyNumberDefault();else sel.Range.ListFormat.RemoveNumbers(); break;
                case "table.insert": doc.Tables.Add(sel.Range,CapabilityArgs.I(a,"rows",2,1,200),CapabilityArgs.I(a,"columns",2,1,63)); break;
                case "table.add_row": { var t=table(); if(CapabilityArgs.S(a,"position","after")=="before")t.Rows.Add(t.Rows[1]);else t.Rows.Add(); break; }
                case "table.add_column": { var t=table(); if(CapabilityArgs.S(a,"position","after")=="before")t.Columns.Add(t.Columns[1]);else t.Columns.Add(); break; }
                case "table.autofit": { string m=CapabilityArgs.S(a,"mode","content"); table().AutoFitBehavior(m=="window"?Word.WdAutoFitBehavior.wdAutoFitWindow:m=="fixed"?Word.WdAutoFitBehavior.wdAutoFitFixed:Word.WdAutoFitBehavior.wdAutoFitContent); break; }
                case "break.page": sel.InsertBreak(Word.WdBreakType.wdPageBreak); break;
                case "break.section": { string t=CapabilityArgs.S(a,"type","nextPage"); sel.InsertBreak(t=="continuous"?Word.WdBreakType.wdSectionBreakContinuous:t=="evenPage"?Word.WdBreakType.wdSectionBreakEvenPage:t=="oddPage"?Word.WdBreakType.wdSectionBreakOddPage:Word.WdBreakType.wdSectionBreakNextPage); break; }
                case "header.set": sel.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text=CapabilityArgs.S(a,"text",""); break;
                case "footer.set": sel.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text=CapabilityArgs.S(a,"text",""); break;
                case "bookmark.add": { string n=CapabilityArgs.S(a,"name",""); if(doc.Bookmarks.Exists(n))doc.Bookmarks[n].Delete(); doc.Bookmarks.Add(n,sel.Range); break; }
                case "bookmark.delete": doc.Bookmarks[CapabilityArgs.S(a,"name","")].Delete(); break;
                case "comment.add": doc.Comments.Add(sel.Range,CapabilityArgs.S(a,"text","")); break;
                case "hyperlink.add": doc.Hyperlinks.Add(sel.Range,CapabilityArgs.S(a,"url",""),Type.Missing,Type.Missing,CapabilityArgs.S(a,"text",sel.Text)); break;
                case "review.track_changes": doc.TrackRevisions=CapabilityArgs.B(a,"enabled",true); break;
                case "review.accept_all": doc.AcceptAllRevisions(); break;
                case "review.reject_all": doc.RejectAllRevisions(); break;
                case "page.orientation": sel.Sections[1].PageSetup.Orientation=CapabilityArgs.S(a,"orientation","portrait")=="landscape"?Word.WdOrientation.wdOrientLandscape:Word.WdOrientation.wdOrientPortrait; break;
                case "page.margins": { var p=sel.Sections[1].PageSetup; p.TopMargin=(float)CapabilityArgs.D(a,"top",72,0,1000);p.BottomMargin=(float)CapabilityArgs.D(a,"bottom",72,0,1000);p.LeftMargin=(float)CapabilityArgs.D(a,"left",72,0,1000);p.RightMargin=(float)CapabilityArgs.D(a,"right",72,0,1000);break; }
                case "page.columns": sel.Sections[1].PageSetup.TextColumns.SetCount(CapabilityArgs.I(a,"count",1,1,12)); break;
                case "field.insert": doc.Fields.Add(sel.Range,Word.WdFieldType.wdFieldEmpty,CapabilityArgs.S(a,"code",""),true); break;
                case "footnote.add": doc.Footnotes.Add(sel.Range,Type.Missing,CapabilityArgs.S(a,"text","")); break;
                case "endnote.add": doc.Endnotes.Add(sel.Range,Type.Missing,CapabilityArgs.S(a,"text","")); break;
                default: throw new ArgumentException("Unsupported Word capability: "+id);
            }
        }
    }

    public static class PowerPointCapabilityEngine
    {
        public static WritePreview Prepare(Ppt.Application app,string json)
        {
            var root=CapabilityArgs.Obj(json); string id=CapabilityArgs.Capability(root);
            if(!OfficeCapabilityRegistry.Exists(HostType.PowerPoint,id)) throw new ArgumentException("Unsupported PowerPoint capability: "+id);
            var p=app.ActivePresentation; if(p==null) throw new InvalidOperationException("No active presentation.");
            return new WritePreview{ToolName=ToolNames.ExecuteOfficeCapability,Title="PowerPoint capability — "+id,Before="Presentation: "+p.Name+"; slides="+p.Slides.Count,After="Execute "+id+" against the specified slide/shape.",ArgumentsJson=json};
        }
        public static void Apply(Ppt.Application app,string json)
        {
            var root=CapabilityArgs.Obj(json); string id=CapabilityArgs.Capability(root); var a=CapabilityArgs.Args(root);
            var p=app.ActivePresentation; if(p==null) throw new InvalidOperationException("No active presentation.");
            Func<Ppt.Slide> slide=()=>p.Slides[CapabilityArgs.I(a,"slide",1,1,p.Slides.Count)];
            Func<Ppt.Shape> shape=()=> { var s=slide(); string raw=CapabilityArgs.S(a,"shape",""); int ix; if(int.TryParse(raw,out ix))return s.Shapes[ix]; return s.Shapes[raw]; };
            switch(id)
            {
                case "slide.add": { int ix=CapabilityArgs.I(a,"index",p.Slides.Count+1,1,p.Slides.Count+1); string l=CapabilityArgs.S(a,"layout","blank"); var layout=l=="title"?Ppt.PpSlideLayout.ppLayoutTitle:l=="text"?Ppt.PpSlideLayout.ppLayoutText:Ppt.PpSlideLayout.ppLayoutBlank; p.Slides.Add(ix,layout); app.ActiveWindow.View.GotoSlide(ix); break; }
                case "slide.delete": slide().Delete(); break;
                case "slide.duplicate": slide().Duplicate(); break;
                case "slide.move": slide().MoveTo(CapabilityArgs.I(a,"index",1,1,p.Slides.Count)); break;
                case "slide.background": { var s=slide(); s.FollowMasterBackground=Office.MsoTriState.msoFalse; s.Background.Fill.ForeColor.RGB=CapabilityArgs.ColorOle(CapabilityArgs.S(a,"color","")); s.Background.Fill.Solid(); break; }
                case "shape.textbox": { var sh=slide().Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,(float)CapabilityArgs.D(a,"left",50,0,5000),(float)CapabilityArgs.D(a,"top",50,0,5000),(float)CapabilityArgs.D(a,"width",300,1,5000),(float)CapabilityArgs.D(a,"height",100,1,5000)); sh.TextFrame.TextRange.Text=CapabilityArgs.S(a,"text",""); sh.Select(); break; }
                case "shape.add": { string t=CapabilityArgs.S(a,"type","rectangle").ToLowerInvariant(); var mt=t=="ellipse"?Office.MsoAutoShapeType.msoShapeOval:t=="roundedrectangle"?Office.MsoAutoShapeType.msoShapeRoundedRectangle:t=="arrow"?Office.MsoAutoShapeType.msoShapeRightArrow:Office.MsoAutoShapeType.msoShapeRectangle; var sh=slide().Shapes.AddShape(mt,(float)CapabilityArgs.D(a,"left",50,0,5000),(float)CapabilityArgs.D(a,"top",50,0,5000),(float)CapabilityArgs.D(a,"width",200,1,5000),(float)CapabilityArgs.D(a,"height",100,1,5000)); sh.Select(); break; }
                case "shape.delete": shape().Delete(); break;
                case "shape.text": { var sh=shape(); if(sh.HasTextFrame!=Office.MsoTriState.msoTrue) throw new InvalidOperationException("Shape has no text frame."); sh.TextFrame.TextRange.Text=CapabilityArgs.S(a,"text",""); break; }
                case "shape.position": { var sh=shape(); if(a["left"]!=null)sh.Left=(float)CapabilityArgs.D(a,"left",sh.Left,-5000,10000);if(a["top"]!=null)sh.Top=(float)CapabilityArgs.D(a,"top",sh.Top,-5000,10000);if(a["width"]!=null)sh.Width=(float)CapabilityArgs.D(a,"width",sh.Width,1,10000);if(a["height"]!=null)sh.Height=(float)CapabilityArgs.D(a,"height",sh.Height,1,10000);break; }
                case "shape.format": { var sh=shape(); string fc=CapabilityArgs.S(a,"fillColor","");if(fc.Length>0){sh.Fill.ForeColor.RGB=CapabilityArgs.ColorOle(fc);sh.Fill.Solid();}string lc=CapabilityArgs.S(a,"lineColor","");if(lc.Length>0)sh.Line.ForeColor.RGB=CapabilityArgs.ColorOle(lc);if(sh.HasTextFrame==Office.MsoTriState.msoTrue){var ft=sh.TextFrame.TextRange.Font;string tc=CapabilityArgs.S(a,"fontColor","");if(tc.Length>0)ft.Color.RGB=CapabilityArgs.ColorOle(tc);if(a["fontSize"]!=null)ft.Size=(float)CapabilityArgs.D(a,"fontSize",18,6,96);if(a["bold"]!=null)ft.Bold=CapabilityArgs.B(a,"bold")?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;}break; }
                case "shape.zorder": { var sh=shape(); string pos=CapabilityArgs.S(a,"position","front"); sh.ZOrder(pos=="back"?Office.MsoZOrderCmd.msoSendToBack:pos=="forward"?Office.MsoZOrderCmd.msoBringForward:pos=="backward"?Office.MsoZOrderCmd.msoSendBackward:Office.MsoZOrderCmd.msoBringToFront);break; }
                case "shape.align": { var sh=shape(); sh.Select(); string al=CapabilityArgs.S(a,"alignment","left"); var range=app.ActiveWindow.Selection.ShapeRange; range.Align(al=="center"?Office.MsoAlignCmd.msoAlignCenters:al=="right"?Office.MsoAlignCmd.msoAlignRights:al=="top"?Office.MsoAlignCmd.msoAlignTops:al=="middle"?Office.MsoAlignCmd.msoAlignMiddles:al=="bottom"?Office.MsoAlignCmd.msoAlignBottoms:Office.MsoAlignCmd.msoAlignLefts,Office.MsoTriState.msoTrue);break; }
                case "table.add": slide().Shapes.AddTable(CapabilityArgs.I(a,"rows",2,1,75),CapabilityArgs.I(a,"columns",2,1,75),(float)CapabilityArgs.D(a,"left",50,0,5000),(float)CapabilityArgs.D(a,"top",50,0,5000),(float)CapabilityArgs.D(a,"width",500,1,5000),(float)CapabilityArgs.D(a,"height",200,1,5000)).Select(); break;
                case "table.cell_text": { var sh=shape(); if(sh.HasTable!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Shape is not a table."); sh.Table.Cell(CapabilityArgs.I(a,"row",1,1,sh.Table.Rows.Count),CapabilityArgs.I(a,"column",1,1,sh.Table.Columns.Count)).Shape.TextFrame.TextRange.Text=CapabilityArgs.S(a,"text","");break; }
                case "notes.set": { var s=slide(); foreach(Ppt.Shape sh in s.NotesPage.Shapes) if(sh.PlaceholderFormat.Type==Ppt.PpPlaceholderType.ppPlaceholderBody){sh.TextFrame.TextRange.Text=CapabilityArgs.S(a,"text","");break;}break; }
                case "hyperlink.add": { var sh=shape(); sh.ActionSettings[Ppt.PpMouseActivation.ppMouseClick].Action=Ppt.PpActionType.ppActionHyperlink; sh.ActionSettings[Ppt.PpMouseActivation.ppMouseClick].Hyperlink.Address=CapabilityArgs.S(a,"url","");break; }
                case "animation.fade": slide().TimeLine.MainSequence.AddEffect(shape(),Ppt.MsoAnimEffect.msoAnimEffectFade,Ppt.MsoAnimateByLevel.msoAnimateLevelNone,Ppt.MsoAnimTriggerType.msoAnimTriggerAfterPrevious); break;
                case "transition.fade": slide().SlideShowTransition.EntryEffect=Ppt.PpEntryEffect.ppEffectFadeSmoothly; break;
                default: throw new ArgumentException("Unsupported PowerPoint capability: "+id);
            }
        }
    }
}
