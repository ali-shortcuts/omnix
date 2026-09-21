using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OMNIX.Core.Context
{
    /// <summary>
    /// Additional professional Office capabilities kept separate from the core registry/executor
    /// so coverage can grow without turning the host adapters into unmaintainable switch blocks.
    /// Every mutation is still reachable only through execute_office_capability and therefore
    /// requires the normal OMNIX preview/confirmation and active-document scope checks.
    /// </summary>
    public static class OfficeCapabilityExtensions
    {
        public static IEnumerable<OfficeCapabilityDescriptor> Descriptors()
        {
            var x = new List<OfficeCapabilityDescriptor>();

            // Excel — workbook calculation/protection, richer data/formatting/chart/shape/page controls.
            Add(x,HostType.Excel,"workbook.protect_structure","Workbook","Protect workbook structure.","password(optional)");
            Add(x,HostType.Excel,"workbook.unprotect","Workbook","Unprotect workbook structure.","password(optional)");
            Add(x,HostType.Excel,"workbook.calculate_full","Calculation","Force a full workbook/application recalculation.","");
            Add(x,HostType.Excel,"worksheet.protect","Worksheet","Protect worksheet contents/UI.","sheet,password(optional)");
            Add(x,HostType.Excel,"worksheet.unprotect","Worksheet","Unprotect worksheet.","sheet,password(optional)");
            Add(x,HostType.Excel,"worksheet.calculate","Calculation","Recalculate one worksheet.","sheet");
            Add(x,HostType.Excel,"view.gridlines","View","Show or hide gridlines in active window.","visible=true|false");
            Add(x,HostType.Excel,"view.headings","View","Show or hide row/column headings.","visible=true|false");
            Add(x,HostType.Excel,"view.zoom","View","Set active Excel window zoom.","percent=10..400");
            Add(x,HostType.Excel,"range.fill_down","Formula","Fill top row formulas/values down through range.","sheet,address");
            Add(x,HostType.Excel,"range.fill_right","Formula","Fill left column formulas/values right through range.","sheet,address");
            Add(x,HostType.Excel,"range.copy","Editing","Copy source range to destination range in workbook.","sheet,address,destinationSheet,destinationAddress");
            Add(x,HostType.Excel,"range.copy_values","Editing","Copy only values to destination range.","sheet,address,destinationSheet,destinationAddress");
            Add(x,HostType.Excel,"range.style","Styles","Apply an existing Excel cell style.","sheet,address,style");
            Add(x,HostType.Excel,"outline.group_rows","Outline","Group target rows.","sheet,address");
            Add(x,HostType.Excel,"outline.ungroup_rows","Outline","Ungroup target rows.","sheet,address");
            Add(x,HostType.Excel,"outline.group_columns","Outline","Group target columns.","sheet,address");
            Add(x,HostType.Excel,"outline.ungroup_columns","Outline","Ungroup target columns.","sheet,address");
            Add(x,HostType.Excel,"validation.whole_number","Data Validation","Validate whole numbers between min/max.","sheet,address,min,max");
            Add(x,HostType.Excel,"validation.decimal","Data Validation","Validate decimals between min/max.","sheet,address,min,max");
            Add(x,HostType.Excel,"validation.date","Data Validation","Validate Excel dates between ISO min/max.","sheet,address,min,max");
            Add(x,HostType.Excel,"conditional.clear","Conditional Formatting","Remove conditional formatting rules.","sheet,address");
            Add(x,HostType.Excel,"conditional.cell_value","Conditional Formatting","Add cell-value comparison formatting.","sheet,address,operator=greaterThan|lessThan|equal,minOrValue,fillColor");
            Add(x,HostType.Excel,"conditional.formula","Conditional Formatting","Add formula-based conditional formatting.","sheet,address,formula,fillColor");
            Add(x,HostType.Excel,"table.resize","Table","Resize a named Excel table.","sheet,name,address");
            Add(x,HostType.Excel,"table.style","Table","Set named Excel table style.","sheet,name,style");
            Add(x,HostType.Excel,"table.totals","Table","Show or hide table totals row.","sheet,name,visible=true|false");
            Add(x,HostType.Excel,"chart.legend","Chart","Show/hide and position chart legend.","sheet,name,visible=true|false,position=right|left|top|bottom");
            Add(x,HostType.Excel,"chart.axis_titles","Chart","Set category/value axis titles.","sheet,name,categoryTitle,valueTitle");
            Add(x,HostType.Excel,"chart.data_labels","Chart","Show or hide series data labels.","sheet,name,visible=true|false");
            Add(x,HostType.Excel,"shape.add","Shapes","Add rectangle, ellipse, roundedRectangle or arrow.","sheet,type,left,top,width,height,name(optional)");
            Add(x,HostType.Excel,"shape.text","Shapes","Set text in a worksheet shape.","sheet,shape,text");
            Add(x,HostType.Excel,"shape.format","Shapes","Format worksheet shape fill/line/text.","sheet,shape,fillColor,lineColor,fontColor,fontSize,bold");
            Add(x,HostType.Excel,"page.margins","Page Layout","Set worksheet page margins in points.","sheet,top,bottom,left,right");
            Add(x,HostType.Excel,"page.header_footer","Page Layout","Set center header/footer text.","sheet,header,footer");
            Add(x,HostType.Excel,"page.center","Page Layout","Center worksheet horizontally/vertically when printing.","sheet,horizontal=true|false,vertical=true|false");
            Add(x,HostType.Excel,"print_titles.rows","Page Layout","Set rows repeated at top when printing.","sheet,address");
            Add(x,HostType.Excel,"print_titles.columns","Page Layout","Set columns repeated at left when printing.","sheet,address");
            Add(x,HostType.Excel,"goal_seek","What-If Analysis","Set a formula cell to goal by changing one cell.","sheet,formulaCell,goal,changingCell");

            // Word — richer formatting, tables, controls, review, references and page structure.
            Add(x,HostType.Word,"selection.case","Editing","Change selected text case.","case=upper|lower|title|sentence");
            Add(x,HostType.Word,"selection.language","Proofing","Set proofing language by numeric Word language ID.","languageId");
            Add(x,HostType.Word,"paragraph.indent","Formatting","Set paragraph left/right/first-line indents in points.","left,right,firstLine");
            Add(x,HostType.Word,"paragraph.pagination","Formatting","Set paragraph keep/page-break behavior.","keepTogether,keepWithNext,pageBreakBefore");
            Add(x,HostType.Word,"selection.shading","Formatting","Set selection shading color.","color=#RRGGBB");
            Add(x,HostType.Word,"selection.border","Formatting","Apply/remove simple selection border.","enabled=true|false");
            Add(x,HostType.Word,"table.delete_row","Tables","Delete current table row.","");
            Add(x,HostType.Word,"table.delete_column","Tables","Delete current table column.","");
            Add(x,HostType.Word,"table.cell_text","Tables","Set text of a cell in current table.","row,column,text");
            Add(x,HostType.Word,"table.cell_shading","Tables","Set table-cell background.","row,column,color=#RRGGBB");
            Add(x,HostType.Word,"table.cell_alignment","Tables","Set vertical alignment of table cell.","row,column,alignment=top|center|bottom");
            Add(x,HostType.Word,"table.repeat_header","Tables","Repeat first table row as header.","enabled=true|false");
            Add(x,HostType.Word,"table.row_height","Tables","Set table row height in points.","row,height");
            Add(x,HostType.Word,"table.column_width","Tables","Set table column width in points.","column,width");
            Add(x,HostType.Word,"table.merge_cells","Tables","Merge a rectangular cell region.","startRow,startColumn,endRow,endColumn");
            Add(x,HostType.Word,"table.split_cell","Tables","Split one table cell.","row,column,rows,columns");
            Add(x,HostType.Word,"content_control.add","Content Controls","Add content control at selection.","type=text|richText|checkBox|date|combo,title,tag");
            Add(x,HostType.Word,"content_control.delete","Content Controls","Delete content control by title/tag, preserving contents optionally.","titleOrTag,deleteContents=false");
            Add(x,HostType.Word,"toc.insert","References","Insert table of contents at selection.","levels=1..9");
            Add(x,HostType.Word,"toc.update","References","Update all tables of contents.","");
            Add(x,HostType.Word,"fields.update","Fields","Update all document fields.","");
            Add(x,HostType.Word,"comment.delete_all","Review","Delete all comments in document.","");
            Add(x,HostType.Word,"review.accept_selection","Review","Accept revisions in current selection.","");
            Add(x,HostType.Word,"review.reject_selection","Review","Reject revisions in current selection.","");
            Add(x,HostType.Word,"header.first_page","Headers/Footers","Set first-page header text in current section.","text");
            Add(x,HostType.Word,"footer.first_page","Headers/Footers","Set first-page footer text in current section.","text");
            Add(x,HostType.Word,"page.first_page_different","Page Layout","Enable/disable different first-page header/footer.","enabled=true|false");
            Add(x,HostType.Word,"page.odd_even_headers","Page Layout","Enable/disable different odd/even headers.","enabled=true|false");
            Add(x,HostType.Word,"page.numbers","Headers/Footers","Add page numbers to primary footer.","alignment=left|center|right");
            Add(x,HostType.Word,"page.size","Page Layout","Set page width/height in points.","width,height");
            Add(x,HostType.Word,"caption.insert","References","Insert caption at selection.","label=Figure|Table|Equation,title");
            Add(x,HostType.Word,"textbox.add","Shapes","Add a text box anchored to selection.","left,top,width,height,text");

            // PowerPoint — more complete slide, shape, text, table, animation and presentation controls.
            Add(x,HostType.PowerPoint,"slide.hidden","Slides","Hide or unhide a slide in slide show.","slide,hidden=true|false");
            Add(x,HostType.PowerPoint,"slide.layout","Slides","Set slide custom layout by one-based master layout index.","slide,layoutIndex");
            Add(x,HostType.PowerPoint,"presentation.size","Design","Set presentation slide width/height in points.","width,height");
            Add(x,HostType.PowerPoint,"shape.duplicate","Shapes","Duplicate a shape.","slide,shape");
            Add(x,HostType.PowerPoint,"shape.rename","Shapes","Rename a shape.","slide,shape,name");
            Add(x,HostType.PowerPoint,"shape.rotate","Shapes","Set shape rotation in degrees.","slide,shape,degrees");
            Add(x,HostType.PowerPoint,"shape.lock_aspect","Shapes","Lock/unlock aspect ratio.","slide,shape,locked=true|false");
            Add(x,HostType.PowerPoint,"shape.transparency","Shapes","Set fill transparency 0..1.","slide,shape,value");
            Add(x,HostType.PowerPoint,"shape.line_weight","Shapes","Set line weight in points.","slide,shape,weight");
            Add(x,HostType.PowerPoint,"shape.shadow","Shapes","Show/hide shape shadow.","slide,shape,visible=true|false");
            Add(x,HostType.PowerPoint,"text.font","Text","Format shape text font.","slide,shape,name,size,bold,italic,color=#RRGGBB");
            Add(x,HostType.PowerPoint,"text.paragraph","Text","Set text paragraph alignment.","slide,shape,alignment=left|center|right|justify");
            Add(x,HostType.PowerPoint,"text.bullets","Text","Enable/disable bullets.","slide,shape,enabled=true|false");
            Add(x,HostType.PowerPoint,"text.autofit","Text","Set text AutoFit behavior.","slide,shape,mode=none|shrink|shapeToFit");
            Add(x,HostType.PowerPoint,"line.add","Shapes","Add a straight line.","slide,x1,y1,x2,y2");
            Add(x,HostType.PowerPoint,"connector.add","Shapes","Add a straight connector.","slide,x1,y1,x2,y2");
            Add(x,HostType.PowerPoint,"table.add_row","Tables","Add row to a table shape.","slide,shape");
            Add(x,HostType.PowerPoint,"table.add_column","Tables","Add column to a table shape.","slide,shape");
            Add(x,HostType.PowerPoint,"table.delete_row","Tables","Delete table row.","slide,shape,row");
            Add(x,HostType.PowerPoint,"table.delete_column","Tables","Delete table column.","slide,shape,column");
            Add(x,HostType.PowerPoint,"table.cell_format","Tables","Format table cell fill/text.","slide,shape,row,column,fillColor,fontColor,fontSize,bold");
            Add(x,HostType.PowerPoint,"transition.duration","Transitions","Set transition duration.","slide,seconds");
            Add(x,HostType.PowerPoint,"transition.advance","Transitions","Configure automatic slide advance.","slide,enabled=true|false,seconds");
            Add(x,HostType.PowerPoint,"animation.appear","Animations","Add appear entrance animation.","slide,shape");
            Add(x,HostType.PowerPoint,"animation.fly","Animations","Add fly entrance animation.","slide,shape");
            Add(x,HostType.PowerPoint,"animation.clear","Animations","Remove all main-sequence animations from slide.","slide");
            Add(x,HostType.PowerPoint,"notes.clear","Notes","Clear speaker notes body.","slide");
            Add(x,HostType.PowerPoint,"shape.visible","Shapes","Show/hide a shape.","slide,shape,visible=true|false");
            Add(x,HostType.PowerPoint,"shape.flip_horizontal","Shapes","Flip shape horizontally.","slide,shape");
            Add(x,HostType.PowerPoint,"shape.flip_vertical","Shapes","Flip shape vertically.","slide,shape");
            return x;
        }

        private static void Add(List<OfficeCapabilityDescriptor> list, HostType host, string id, string category, string description, string args)
        {
            list.Add(new OfficeCapabilityDescriptor { Host=host, Id=id, Category=category, Description=description, Arguments=args });
        }

        public static bool TryApplyExcel(Excel.Application app, string id, JObject a)
        {
            var wb=app.ActiveWorkbook; if(wb==null) throw new InvalidOperationException("No active workbook.");
            Func<Excel.Worksheet> sheet=()=> {
                string n=CapabilityArgs.S(a,"sheet","");
                var ws=string.IsNullOrWhiteSpace(n)?app.ActiveSheet as Excel.Worksheet:wb.Worksheets[n] as Excel.Worksheet;
                if(ws==null) throw new InvalidOperationException("Worksheet not found."); return ws; };
            Func<Excel.Range> range=()=> {
                string ad=CapabilityArgs.S(a,"address",""); if(ad.Length==0) throw new ArgumentException("Missing address.");
                var r=sheet().Range[ad]; if(Convert.ToInt64(r.Cells.CountLarge)>10000) throw new ArgumentException("Range exceeds 10,000 cells."); return r; };

            switch(id)
            {
                case "workbook.protect_structure": wb.Protect(CapabilityArgs.S(a,"password",""),true,false); return true;
                case "workbook.unprotect": wb.Unprotect(CapabilityArgs.S(a,"password","")); return true;
                case "workbook.calculate_full": app.CalculateFull(); return true;
                case "worksheet.protect": sheet().Protect(CapabilityArgs.S(a,"password","")); return true;
                case "worksheet.unprotect": sheet().Unprotect(CapabilityArgs.S(a,"password","")); return true;
                case "worksheet.calculate": sheet().Calculate(); return true;
                case "view.gridlines": app.ActiveWindow.DisplayGridlines=CapabilityArgs.B(a,"visible",true); return true;
                case "view.headings": app.ActiveWindow.DisplayHeadings=CapabilityArgs.B(a,"visible",true); return true;
                case "view.zoom": app.ActiveWindow.Zoom=CapabilityArgs.I(a,"percent",100,10,400); return true;
                case "range.fill_down": range().FillDown(); return true;
                case "range.fill_right": range().FillRight(); return true;
                case "range.copy": {
                    var src=range(); string ds=CapabilityArgs.S(a,"destinationSheet",sheet().Name); string da=CapabilityArgs.S(a,"destinationAddress","");
                    if(da.Length==0) throw new ArgumentException("Missing destinationAddress.");
                    var dest=(wb.Worksheets[ds] as Excel.Worksheet).Range[da]; src.Copy(dest); return true; }
                case "range.copy_values": {
                    var src=range(); string ds=CapabilityArgs.S(a,"destinationSheet",sheet().Name); string da=CapabilityArgs.S(a,"destinationAddress","");
                    var dest=(wb.Worksheets[ds] as Excel.Worksheet).Range[da]; dest.Resize[src.Rows.Count,src.Columns.Count].Value2=src.Value2; return true; }
                case "range.style": range().Style=CapabilityArgs.S(a,"style","Normal"); return true;
                case "outline.group_rows": range().EntireRow.Group(); return true;
                case "outline.ungroup_rows": range().EntireRow.Ungroup(); return true;
                case "outline.group_columns": range().EntireColumn.Group(); return true;
                case "outline.ungroup_columns": range().EntireColumn.Ungroup(); return true;
                case "validation.whole_number": ApplyValidation(range(),Excel.XlDVType.xlValidateWholeNumber,a,false); return true;
                case "validation.decimal": ApplyValidation(range(),Excel.XlDVType.xlValidateDecimal,a,false); return true;
                case "validation.date": ApplyValidation(range(),Excel.XlDVType.xlValidateDate,a,true); return true;
                case "conditional.clear": range().FormatConditions.Delete(); return true;
                case "conditional.cell_value": {
                    var r=range(); string op=CapabilityArgs.S(a,"operator","greaterThan");
                    var xo=op=="lessThan"?Excel.XlFormatConditionOperator.xlLess:op=="equal"?Excel.XlFormatConditionOperator.xlEqual:Excel.XlFormatConditionOperator.xlGreater;
                    var fc=(Excel.FormatCondition)r.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,xo,CapabilityArgs.S(a,"minOrValue","0"));
                    fc.Interior.Color=CapabilityArgs.ColorOle(CapabilityArgs.S(a,"fillColor","#FFF2CC")); return true; }
                case "conditional.formula": {
                    var r=range(); string formula=CapabilityArgs.S(a,"formula",""); if(!formula.StartsWith("=")) throw new ArgumentException("Formula must start with =.");
                    var fc=(Excel.FormatCondition)r.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression,Type.Missing,formula);
                    fc.Interior.Color=CapabilityArgs.ColorOle(CapabilityArgs.S(a,"fillColor","#FFF2CC")); return true; }
                case "table.resize": { var t=sheet().ListObjects[CapabilityArgs.S(a,"name","")]; t.Resize(sheet().Range[CapabilityArgs.S(a,"address","")]); return true; }
                case "table.style": sheet().ListObjects[CapabilityArgs.S(a,"name","")].TableStyle=CapabilityArgs.S(a,"style","TableStyleMedium2"); return true;
                case "table.totals": sheet().ListObjects[CapabilityArgs.S(a,"name","")].ShowTotals=CapabilityArgs.B(a,"visible",true); return true;
                case "chart.legend": {
                    var ch=((Excel.ChartObject)sheet().ChartObjects(CapabilityArgs.S(a,"name",""))).Chart; bool visible=CapabilityArgs.B(a,"visible",true); ch.HasLegend=visible;
                    if(visible){string pos=CapabilityArgs.S(a,"position","right");ch.Legend.Position=pos=="left"?Excel.XlLegendPosition.xlLegendPositionLeft:pos=="top"?Excel.XlLegendPosition.xlLegendPositionTop:pos=="bottom"?Excel.XlLegendPosition.xlLegendPositionBottom:Excel.XlLegendPosition.xlLegendPositionRight;} return true; }
                case "chart.axis_titles": {
                    var ch=((Excel.ChartObject)sheet().ChartObjects(CapabilityArgs.S(a,"name",""))).Chart;
                    string ct=CapabilityArgs.S(a,"categoryTitle",""); string vt=CapabilityArgs.S(a,"valueTitle","");
                    if(ct.Length>0){var ax=(Excel.Axis)ch.Axes(Excel.XlAxisType.xlCategory,Excel.XlAxisGroup.xlPrimary);ax.HasTitle=true;ax.AxisTitle.Text=ct;}
                    if(vt.Length>0){var ax=(Excel.Axis)ch.Axes(Excel.XlAxisType.xlValue,Excel.XlAxisGroup.xlPrimary);ax.HasTitle=true;ax.AxisTitle.Text=vt;} return true; }
                case "chart.data_labels": {
                    var ch=((Excel.ChartObject)sheet().ChartObjects(CapabilityArgs.S(a,"name",""))).Chart; bool show=CapabilityArgs.B(a,"visible",true);
                    foreach(Excel.Series s in (Excel.SeriesCollection)ch.SeriesCollection()) { if(show)s.ApplyDataLabels(); else if(s.HasDataLabels)s.DataLabels().Delete(); } return true; }
                case "shape.add": {
                    var ws=sheet(); string type=CapabilityArgs.S(a,"type","rectangle").ToLowerInvariant();
                    var mt=type=="ellipse"?Office.MsoAutoShapeType.msoShapeOval:type=="roundedrectangle"?Office.MsoAutoShapeType.msoShapeRoundedRectangle:type=="arrow"?Office.MsoAutoShapeType.msoShapeRightArrow:Office.MsoAutoShapeType.msoShapeRectangle;
                    var sh=ws.Shapes.AddShape(mt,(float)CapabilityArgs.D(a,"left",50,0,10000),(float)CapabilityArgs.D(a,"top",50,0,10000),(float)CapabilityArgs.D(a,"width",150,1,10000),(float)CapabilityArgs.D(a,"height",80,1,10000));
                    string name=CapabilityArgs.S(a,"name","");if(name.Length>0)sh.Name=name;sh.Select();return true; }
                case "shape.text": { var sh=sheet().Shapes.Item(CapabilityArgs.S(a,"shape","")); sh.TextFrame2.TextRange.Text=CapabilityArgs.S(a,"text",""); return true; }
                case "shape.format": {
                    var sh=sheet().Shapes.Item(CapabilityArgs.S(a,"shape","")); string v=CapabilityArgs.S(a,"fillColor","");if(v.Length>0){sh.Fill.ForeColor.RGB=CapabilityArgs.ColorOle(v);sh.Fill.Solid();}
                    v=CapabilityArgs.S(a,"lineColor","");if(v.Length>0)sh.Line.ForeColor.RGB=CapabilityArgs.ColorOle(v);
                    if(sh.TextFrame2!=null){var f=sh.TextFrame2.TextRange.Font;v=CapabilityArgs.S(a,"fontColor","");if(v.Length>0)f.Fill.ForeColor.RGB=CapabilityArgs.ColorOle(v);if(a["fontSize"]!=null)f.Size=(float)CapabilityArgs.D(a,"fontSize",11,6,96);if(a["bold"]!=null)f.Bold=CapabilityArgs.B(a,"bold")?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;} return true; }
                case "page.margins": { var p=sheet().PageSetup;p.TopMargin=CapabilityArgs.D(a,"top",54,0,1000);p.BottomMargin=CapabilityArgs.D(a,"bottom",54,0,1000);p.LeftMargin=CapabilityArgs.D(a,"left",54,0,1000);p.RightMargin=CapabilityArgs.D(a,"right",54,0,1000);return true; }
                case "page.header_footer": { var p=sheet().PageSetup;p.CenterHeader=CapabilityArgs.S(a,"header","");p.CenterFooter=CapabilityArgs.S(a,"footer","");return true; }
                case "page.center": { var p=sheet().PageSetup;p.CenterHorizontally=CapabilityArgs.B(a,"horizontal",false);p.CenterVertically=CapabilityArgs.B(a,"vertical",false);return true; }
                case "print_titles.rows": sheet().PageSetup.PrintTitleRows=sheet().Range[CapabilityArgs.S(a,"address","")].EntireRow.Address; return true;
                case "print_titles.columns": sheet().PageSetup.PrintTitleColumns=sheet().Range[CapabilityArgs.S(a,"address","")].EntireColumn.Address; return true;
                case "goal_seek": {
                    var ws=sheet();var formula=ws.Range[CapabilityArgs.S(a,"formulaCell","")];var changing=ws.Range[CapabilityArgs.S(a,"changingCell","")];
                    if(!Convert.ToBoolean(formula.HasFormula))throw new InvalidOperationException("formulaCell must contain a formula.");
                    formula.GoalSeek(CapabilityArgs.D(a,"goal",0,-1e15,1e15),changing);return true; }
                default: return false;
            }
        }

        private static void ApplyValidation(Excel.Range r, Excel.XlDVType type, JObject a, bool date)
        {
            r.Validation.Delete();
            object min=date?DateTime.Parse(CapabilityArgs.S(a,"min","2000-01-01"),CultureInfo.InvariantCulture).ToOADate():(object)CapabilityArgs.D(a,"min",0,-1e15,1e15);
            object max=date?DateTime.Parse(CapabilityArgs.S(a,"max","2100-01-01"),CultureInfo.InvariantCulture).ToOADate():(object)CapabilityArgs.D(a,"max",100,-1e15,1e15);
            r.Validation.Add(type,Excel.XlDVAlertStyle.xlValidAlertStop,Excel.XlFormatConditionOperator.xlBetween,min,max);
            r.Validation.IgnoreBlank=true;
        }

        public static bool TryApplyWord(Word.Application app, string id, JObject a)
        {
            var doc=app.ActiveDocument;if(doc==null)throw new InvalidOperationException("No active Word document.");var sel=app.Selection;
            Func<Word.Table> table=()=>{if(sel.Tables.Count<1)throw new InvalidOperationException("Selection is not inside a table.");return sel.Tables[1];};
            switch(id)
            {
                case "selection.case": { string v=CapabilityArgs.S(a,"case","upper");sel.Range.Case=v=="lower"?Word.WdCharacterCase.wdLowerCase:v=="title"?Word.WdCharacterCase.wdTitleWord:v=="sentence"?Word.WdCharacterCase.wdTitleSentence:Word.WdCharacterCase.wdUpperCase;return true; }
                case "selection.language": sel.Range.LanguageID=(Word.WdLanguageID)CapabilityArgs.I(a,"languageId",1033,0,99999);return true;
                case "paragraph.indent": { var p=sel.ParagraphFormat;p.LeftIndent=(float)CapabilityArgs.D(a,"left",0,-1000,1000);p.RightIndent=(float)CapabilityArgs.D(a,"right",0,-1000,1000);p.FirstLineIndent=(float)CapabilityArgs.D(a,"firstLine",0,-1000,1000);return true; }
                case "paragraph.pagination": { var p=sel.ParagraphFormat;p.KeepTogether=CapabilityArgs.B(a,"keepTogether",false)?-1:0;p.KeepWithNext=CapabilityArgs.B(a,"keepWithNext",false)?-1:0;p.PageBreakBefore=CapabilityArgs.B(a,"pageBreakBefore",false)?-1:0;return true; }
                case "selection.shading": sel.Range.Shading.BackgroundPatternColor=(Word.WdColor)CapabilityArgs.ColorOle(CapabilityArgs.S(a,"color",""));return true;
                case "selection.border": sel.Range.Borders.Enable=CapabilityArgs.B(a,"enabled",true)?1:0;return true;
                case "table.delete_row": table().Rows[1].Delete();return true;
                case "table.delete_column": table().Columns[1].Delete();return true;
                case "table.cell_text": { var t=table();t.Cell(CapabilityArgs.I(a,"row",1,1,t.Rows.Count),CapabilityArgs.I(a,"column",1,1,t.Columns.Count)).Range.Text=CapabilityArgs.S(a,"text","");return true; }
                case "table.cell_shading": { var t=table();t.Cell(CapabilityArgs.I(a,"row",1,1,t.Rows.Count),CapabilityArgs.I(a,"column",1,1,t.Columns.Count)).Shading.BackgroundPatternColor=(Word.WdColor)CapabilityArgs.ColorOle(CapabilityArgs.S(a,"color",""));return true; }
                case "table.cell_alignment": { var t=table();var c=t.Cell(CapabilityArgs.I(a,"row",1,1,t.Rows.Count),CapabilityArgs.I(a,"column",1,1,t.Columns.Count));string v=CapabilityArgs.S(a,"alignment","top");c.VerticalAlignment=v=="center"?Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter:v=="bottom"?Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom:Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;return true; }
                case "table.repeat_header": table().Rows[1].HeadingFormat=CapabilityArgs.B(a,"enabled",true)?-1:0;return true;
                case "table.row_height": {var t=table();t.Rows[CapabilityArgs.I(a,"row",1,1,t.Rows.Count)].Height=(float)CapabilityArgs.D(a,"height",18,1,1000);return true;}
                case "table.column_width": {var t=table();t.Columns[CapabilityArgs.I(a,"column",1,1,t.Columns.Count)].Width=(float)CapabilityArgs.D(a,"width",72,1,1000);return true;}
                case "table.merge_cells": { var t=table();t.Cell(CapabilityArgs.I(a,"startRow",1,1,t.Rows.Count),CapabilityArgs.I(a,"startColumn",1,1,t.Columns.Count)).Merge(t.Cell(CapabilityArgs.I(a,"endRow",1,1,t.Rows.Count),CapabilityArgs.I(a,"endColumn",1,1,t.Columns.Count)));return true; }
                case "table.split_cell": { var t=table();t.Cell(CapabilityArgs.I(a,"row",1,1,t.Rows.Count),CapabilityArgs.I(a,"column",1,1,t.Columns.Count)).Split(CapabilityArgs.I(a,"rows",1,1,20),CapabilityArgs.I(a,"columns",2,1,20));return true; }
                case "content_control.add": { string type=CapabilityArgs.S(a,"type","text");var wt=type=="richtext"?Word.WdContentControlType.wdContentControlRichText:type=="checkbox"?Word.WdContentControlType.wdContentControlCheckBox:type=="date"?Word.WdContentControlType.wdContentControlDate:type=="combo"?Word.WdContentControlType.wdContentControlComboBox:Word.WdContentControlType.wdContentControlText;var cc=doc.ContentControls.Add(wt,sel.Range);cc.Title=CapabilityArgs.S(a,"title","");cc.Tag=CapabilityArgs.S(a,"tag","");return true; }
                case "content_control.delete": { string key=CapabilityArgs.S(a,"titleOrTag","");foreach(Word.ContentControl cc in doc.ContentControls){if(string.Equals(cc.Title,key,StringComparison.OrdinalIgnoreCase)||string.Equals(cc.Tag,key,StringComparison.OrdinalIgnoreCase)){cc.Delete(CapabilityArgs.B(a,"deleteContents",false));return true;}}throw new InvalidOperationException("Content control not found."); }
                case "toc.insert": doc.TablesOfContents.Add(sel.Range,true,1,CapabilityArgs.I(a,"levels",3,1,9));return true;
                case "toc.update": foreach(Word.TableOfContents t in doc.TablesOfContents)t.Update();return true;
                case "fields.update": doc.Fields.Update();return true;
                case "comment.delete_all": while(doc.Comments.Count>0)doc.Comments[1].Delete();return true;
                case "review.accept_selection": sel.Range.Revisions.AcceptAll();return true;
                case "review.reject_selection": sel.Range.Revisions.RejectAll();return true;
                case "header.first_page": sel.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text=CapabilityArgs.S(a,"text","");return true;
                case "footer.first_page": sel.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text=CapabilityArgs.S(a,"text","");return true;
                case "page.first_page_different": sel.Sections[1].PageSetup.DifferentFirstPageHeaderFooter=CapabilityArgs.B(a,"enabled",true)?-1:0;return true;
                case "page.odd_even_headers": sel.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter=CapabilityArgs.B(a,"enabled",true)?-1:0;return true;
                case "page.numbers": { string al=CapabilityArgs.S(a,"alignment","center");sel.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Add(al=="left"?Word.WdPageNumberAlignment.wdAlignPageNumberLeft:al=="right"?Word.WdPageNumberAlignment.wdAlignPageNumberRight:Word.WdPageNumberAlignment.wdAlignPageNumberCenter,true);return true; }
                case "page.size": {var p=sel.Sections[1].PageSetup;p.PageWidth=(float)CapabilityArgs.D(a,"width",612,100,2000);p.PageHeight=(float)CapabilityArgs.D(a,"height",792,100,3000);return true;}
                case "caption.insert": { string label=CapabilityArgs.S(a,"label","Figure");sel.InsertCaption(label,CapabilityArgs.S(a,"title",""),Type.Missing,Word.WdCaptionPosition.wdCaptionPositionBelow);return true; }
                case "textbox.add": { var sh=doc.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,(float)CapabilityArgs.D(a,"left",50,0,5000),(float)CapabilityArgs.D(a,"top",50,0,5000),(float)CapabilityArgs.D(a,"width",250,1,5000),(float)CapabilityArgs.D(a,"height",80,1,5000),sel.Range);sh.TextFrame.TextRange.Text=CapabilityArgs.S(a,"text","");return true; }
                default:return false;
            }
        }

        public static bool TryApplyPowerPoint(Ppt.Application app, string id, JObject a)
        {
            var p=app.ActivePresentation;if(p==null)throw new InvalidOperationException("No active presentation.");
            Func<Ppt.Slide> slide=()=>p.Slides[CapabilityArgs.I(a,"slide",1,1,p.Slides.Count)];
            Func<Ppt.Shape> shape=()=>{var s=slide();string raw=CapabilityArgs.S(a,"shape","");int ix;if(int.TryParse(raw,out ix))return s.Shapes[ix];return s.Shapes[raw];};
            switch(id)
            {
                case "slide.hidden": slide().SlideShowTransition.Hidden=CapabilityArgs.B(a,"hidden",true)?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;return true;
                case "slide.layout": { int ix=CapabilityArgs.I(a,"layoutIndex",1,1,p.SlideMaster.CustomLayouts.Count);slide().CustomLayout=p.SlideMaster.CustomLayouts[ix];return true; }
                case "presentation.size": p.PageSetup.SlideWidth=(float)CapabilityArgs.D(a,"width",960,100,5000);p.PageSetup.SlideHeight=(float)CapabilityArgs.D(a,"height",540,100,5000);return true;
                case "shape.duplicate": shape().Duplicate().Select();return true;
                case "shape.rename": shape().Name=CapabilityArgs.S(a,"name","Shape");return true;
                case "shape.rotate": shape().Rotation=(float)CapabilityArgs.D(a,"degrees",0,-360,360);return true;
                case "shape.lock_aspect": shape().LockAspectRatio=CapabilityArgs.B(a,"locked",true)?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;return true;
                case "shape.transparency": shape().Fill.Transparency=(float)CapabilityArgs.D(a,"value",0,0,1);return true;
                case "shape.line_weight": shape().Line.Weight=(float)CapabilityArgs.D(a,"weight",1,0.1,50);return true;
                case "shape.shadow": shape().Shadow.Visible=CapabilityArgs.B(a,"visible",true)?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;return true;
                case "text.font": {var sh=shape();if(sh.HasTextFrame!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Shape has no text.");var f=sh.TextFrame.TextRange.Font;string n=CapabilityArgs.S(a,"name","");if(n.Length>0)f.Name=n;if(a["size"]!=null)f.Size=(float)CapabilityArgs.D(a,"size",18,6,96);if(a["bold"]!=null)f.Bold=CapabilityArgs.B(a,"bold")?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;if(a["italic"]!=null)f.Italic=CapabilityArgs.B(a,"italic")?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;string co=CapabilityArgs.S(a,"color","");if(co.Length>0)f.Color.RGB=CapabilityArgs.ColorOle(co);return true;}
                case "text.paragraph": {var sh=shape();var pf=sh.TextFrame.TextRange.ParagraphFormat;string al=CapabilityArgs.S(a,"alignment","left");pf.Alignment=al=="center"?Ppt.PpParagraphAlignment.ppAlignCenter:al=="right"?Ppt.PpParagraphAlignment.ppAlignRight:al=="justify"?Ppt.PpParagraphAlignment.ppAlignJustify:Ppt.PpParagraphAlignment.ppAlignLeft;return true;}
                case "text.bullets": shape().TextFrame.TextRange.ParagraphFormat.Bullet.Visible=CapabilityArgs.B(a,"enabled",true)?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;return true;
                case "text.autofit": {string m=CapabilityArgs.S(a,"mode","none");shape().TextFrame2.AutoSize=m=="shrink"?Office.MsoAutoSize.msoAutoSizeTextToFitShape:m=="shapetofit"?Office.MsoAutoSize.msoAutoSizeShapeToFitText:Office.MsoAutoSize.msoAutoSizeNone;return true;}
                case "line.add": slide().Shapes.AddLine((float)CapabilityArgs.D(a,"x1",10,-5000,10000),(float)CapabilityArgs.D(a,"y1",10,-5000,10000),(float)CapabilityArgs.D(a,"x2",200,-5000,10000),(float)CapabilityArgs.D(a,"y2",10,-5000,10000)).Select();return true;
                case "connector.add": slide().Shapes.AddConnector(Office.MsoConnectorType.msoConnectorStraight,(float)CapabilityArgs.D(a,"x1",10,-5000,10000),(float)CapabilityArgs.D(a,"y1",10,-5000,10000),(float)CapabilityArgs.D(a,"x2",200,-5000,10000),(float)CapabilityArgs.D(a,"y2",10,-5000,10000)).Select();return true;
                case "table.add_row": {var sh=shape();if(sh.HasTable!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Shape is not a table.");sh.Table.Rows.Add();return true;}
                case "table.add_column": {var sh=shape();if(sh.HasTable!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Shape is not a table.");sh.Table.Columns.Add();return true;}
                case "table.delete_row": {var sh=shape();sh.Table.Rows[CapabilityArgs.I(a,"row",1,1,sh.Table.Rows.Count)].Delete();return true;}
                case "table.delete_column": {var sh=shape();sh.Table.Columns[CapabilityArgs.I(a,"column",1,1,sh.Table.Columns.Count)].Delete();return true;}
                case "table.cell_format": {var sh=shape();var cell=sh.Table.Cell(CapabilityArgs.I(a,"row",1,1,sh.Table.Rows.Count),CapabilityArgs.I(a,"column",1,1,sh.Table.Columns.Count)).Shape;string v=CapabilityArgs.S(a,"fillColor","");if(v.Length>0){cell.Fill.ForeColor.RGB=CapabilityArgs.ColorOle(v);cell.Fill.Solid();}var f=cell.TextFrame.TextRange.Font;v=CapabilityArgs.S(a,"fontColor","");if(v.Length>0)f.Color.RGB=CapabilityArgs.ColorOle(v);if(a["fontSize"]!=null)f.Size=(float)CapabilityArgs.D(a,"fontSize",14,6,96);if(a["bold"]!=null)f.Bold=CapabilityArgs.B(a,"bold")?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;return true;}
                case "transition.duration": slide().SlideShowTransition.Duration=(float)CapabilityArgs.D(a,"seconds",1,0,60);return true;
                case "transition.advance": {var t=slide().SlideShowTransition;bool en=CapabilityArgs.B(a,"enabled",true);t.AdvanceOnTime=en?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;t.AdvanceTime=(float)CapabilityArgs.D(a,"seconds",5,0,3600);return true;}
                case "animation.appear": slide().TimeLine.MainSequence.AddEffect(shape(),Ppt.MsoAnimEffect.msoAnimEffectAppear,Ppt.MsoAnimateByLevel.msoAnimateLevelNone,Ppt.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);return true;
                case "animation.fly": slide().TimeLine.MainSequence.AddEffect(shape(),Ppt.MsoAnimEffect.msoAnimEffectFly,Ppt.MsoAnimateByLevel.msoAnimateLevelNone,Ppt.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);return true;
                case "animation.clear": {var seq=slide().TimeLine.MainSequence;while(seq.Count>0)seq[1].Delete();return true;}
                case "notes.clear": {foreach(Ppt.Shape sh in slide().NotesPage.Shapes)if(sh.PlaceholderFormat.Type==Ppt.PpPlaceholderType.ppPlaceholderBody){sh.TextFrame.TextRange.Text="";break;}return true;}
                case "shape.visible": shape().Visible=CapabilityArgs.B(a,"visible",true)?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;return true;
                case "shape.flip_horizontal": shape().Flip(Office.MsoFlipCmd.msoFlipHorizontal);return true;
                case "shape.flip_vertical": shape().Flip(Office.MsoFlipCmd.msoFlipVertical);return true;
                default:return false;
            }
        }
    }
}
