using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using OMNIX.Core.Errors;
using OMNIX.Core.Context;
using OMNIX.Core.Util;

namespace OMNIX.Core.Tools
{
    internal static class ExcelAdvancedCapabilities
    {
        private const int MaxRangeCells = 10000;
        private const int MaxInspectCells = 256;
        private const int MaxText = 12000;

        public static string Inspect(ExcelHostAdapter adapter, ToolArguments args)
        {
            string op = Op(args);
            if (!OfficeCapabilityRegistry.IsKnown(HostType.Excel, op, false))
                throw new OmnixException(ErrorCode.CORE_ERROR, "Unknown Excel inspection operation: " + op, op, "Use list_office_capabilities.");

            var app = adapter.App;
            var wb = app.ActiveWorkbook;
            if (wb == null) throw new InvalidOperationException("No Excel workbook is active.");

            switch (op)
            {
                case "workbook.summary": return WorkbookSummary(wb);
                case "sheet.inspect": return SheetInspect(wb, args);
                case "range.inspect": return RangeInspect(wb, args);
                case "table.inspect": return TableInspect(wb, args);
                case "name.inspect": return NameInspect(wb, args);
                case "chart.inspect": return ChartInspect(wb, args);
                case "pivot.inspect": return PivotInspect(wb, args);
                case "validation.inspect": return ValidationInspect(wb, args);
                case "conditional_format.inspect": return ConditionalInspect(wb, args);
                case "filter.inspect": return FilterInspect(wb, args);
                case "page_setup.inspect": return PageSetupInspect(wb, args);
                case "view.inspect": return ViewInspect(app);
                default: throw new InvalidOperationException("Unhandled Excel inspection operation: " + op);
            }
        }

        public static WritePreview Prepare(ExcelHostAdapter adapter, string json)
        {
            var args = ToolArguments.Parse(json);
            string op = Op(args);
            if (!OfficeCapabilityRegistry.IsKnown(HostType.Excel, op, true))
                throw new OmnixException(ErrorCode.CORE_ERROR, "Unknown Excel write operation: " + op, op, "Use list_office_capabilities.");

            ValidateWrite(adapter, args, op);
            string before = PreviewBefore(adapter, args, op);
            string after = DescribeAfter(args, op);
            return new WritePreview
            {
                ToolName = ToolNames.ApplyOfficeCapability,
                Title = "Excel — " + op,
                Before = before,
                After = after,
                ArgumentsJson = json
            };
        }

        public static void Apply(ExcelHostAdapter adapter, string json)
        {
            var args = ToolArguments.Parse(json);
            string op = Op(args);
            if (!OfficeCapabilityRegistry.IsKnown(HostType.Excel, op, true))
                throw new OmnixException(ErrorCode.CORE_ERROR, "Unknown Excel write operation: " + op, op, "Use list_office_capabilities.");
            ValidateWrite(adapter, args, op);

            var app = adapter.App;
            var wb = app.ActiveWorkbook;
            if (wb == null) throw new InvalidOperationException("No Excel workbook is active.");

            switch (op)
            {
                case "sheet.create": SheetCreate(wb,args); break;
                case "sheet.rename": SheetRename(wb,args); break;
                case "sheet.delete": Sheet(wb,args.Get("sheet","")).Delete(); break;
                case "sheet.copy": SheetCopy(wb,args); break;
                case "sheet.move": SheetMove(wb,args); break;
                case "sheet.visibility": SheetVisibility(wb,args); break;
                case "range.clear": RangeClear(wb,args); break;
                case "range.insert": RangeInsert(wb,args); break;
                case "range.delete": RangeDelete(wb,args); break;
                case "range.merge": RangeMerge(wb,args); break;
                case "range.autofill": RangeAutoFill(wb,args); break;
                case "row.height": RowHeight(wb,args); break;
                case "column.width": ColumnWidth(wb,args); break;
                case "row.visibility": RowVisibility(wb,args); break;
                case "column.visibility": ColumnVisibility(wb,args); break;
                case "table.create": TableCreate(wb,args); break;
                case "table.resize": TableResize(wb,args); break;
                case "table.rename": TableRename(wb,args); break;
                case "table.style": TableStyle(wb,args); break;
                case "table.delete": TableDelete(wb,args); break;
                case "sort.apply": SortApply(wb,args); break;
                case "filter.apply": FilterApply(wb,args); break;
                case "filter.clear": FilterClear(wb,args); break;
                case "validation.add": ValidationAdd(wb,args); break;
                case "validation.delete": Range(wb,args).Validation.Delete(); break;
                case "conditional_format.add": ConditionalAdd(wb,args); break;
                case "conditional_format.clear": Range(wb,args).FormatConditions.Delete(); break;
                case "name.add": NameAdd(wb,args); break;
                case "name.delete": wb.Names.Item(args.Get("name","")).Delete(); break;
                case "chart.create": ChartCreate(wb,args); break;
                case "chart.delete": ChartObject(wb,args).Delete(); break;
                case "chart.title": ChartTitle(wb,args); break;
                case "pivot.refresh": PivotRefresh(wb,args); break;
                case "comment.set": CommentSet(wb,args); break;
                case "comment.delete": RangeSingle(wb,args).ClearComments(); break;
                case "hyperlink.add": HyperlinkAdd(wb,args); break;
                case "view.freeze": ViewFreeze(app,wb,args); break;
                case "view.zoom": app.ActiveWindow.Zoom = Int(args,"zoom",10,400); break;
                case "page_setup.set": PageSetupSet(wb,args); break;
                default: throw new InvalidOperationException("Unhandled Excel write operation: " + op);
            }
        }

        private static string Op(ToolArguments args)
        {
            string op = (args.Get("operation","") ?? "").Trim();
            if (op.Length == 0) throw new ArgumentException("operation is required.");
            return op;
        }

        private static void ValidateWrite(ExcelHostAdapter adapter, ToolArguments args, string op)
        {
            var wb = adapter.App.ActiveWorkbook;
            if (wb == null || wb.ReadOnly) throw new InvalidOperationException("An editable Excel workbook is required.");

            if (op.StartsWith("range.",StringComparison.Ordinal) || op.StartsWith("row.",StringComparison.Ordinal) ||
                op.StartsWith("column.",StringComparison.Ordinal) || op.StartsWith("validation.",StringComparison.Ordinal) ||
                op.StartsWith("conditional_format.",StringComparison.Ordinal))
            {
                var r = Range(wb,args);
                if (CellCount(r) > Math.Min(MaxRangeCells, Math.Max(1,adapter.MaxCells)))
                    throw new ArgumentException("The requested Excel range is too large for one confirmed capability operation.");
            }

            if (op == "sheet.delete" && wb.Worksheets.Count <= 1)
                throw new InvalidOperationException("Excel must keep at least one worksheet.");

            if (op == "view.zoom") Int(args,"zoom",10,400);
            if (op == "sheet.move") Int(args,"index",1,wb.Sheets.Count);
            if (op == "chart.create")
            {
                Double(args,"width",80,2000,480); Double(args,"height",60,1600,280);
            }
            if (op == "table.create" || op == "table.resize")
            {
                var r=Range(wb,args); if(CellCount(r)>MaxRangeCells) throw new ArgumentException("Table range is too large.");
            }
        }

        private static string PreviewBefore(ExcelHostAdapter adapter, ToolArguments args, string op)
        {
            var wb=adapter.App.ActiveWorkbook;
            try
            {
                if (op.StartsWith("range.",StringComparison.Ordinal) || op.StartsWith("row.",StringComparison.Ordinal) ||
                    op.StartsWith("column.",StringComparison.Ordinal) || op.StartsWith("validation.",StringComparison.Ordinal) ||
                    op.StartsWith("conditional_format.",StringComparison.Ordinal))
                {
                    var r=Range(wb,args);
                    return "Target " + r.Worksheet.Name + "!" + r.Address[false,false] + " (" + CellCount(r) +
                           " cells). Current values preview:\n" + adapter.BuildValuesTable(r, Math.Min(30,adapter.MaxCells));
                }
                if (op.StartsWith("sheet.",StringComparison.Ordinal) && op!="sheet.create")
                {
                    var ws=Sheet(wb,args.Get("sheet",""));
                    return "Worksheet '" + ws.Name + "'; used=" + ws.UsedRange.Address[false,false] +
                           "; visibility=" + ws.Visible + "; tables=" + ws.ListObjects.Count + ".";
                }
                if (op.StartsWith("table.",StringComparison.Ordinal) && op!="table.create")
                {
                    var t=Table(wb,args);
                    return "Table '" + t.Name + "' at " + t.Range.Address[false,false] +
                           "; rows=" + t.ListRows.Count + "; columns=" + t.ListColumns.Count + ".";
                }
                if (op.StartsWith("chart.",StringComparison.Ordinal) && op!="chart.create")
                {
                    var co=ChartObject(wb,args);
                    var parent = co.Parent as Excel.Worksheet; return "Chart '" + co.Name + "' on " + (parent != null ? parent.Name : "?") + ".";
                }
            }
            catch { }
            return "Active workbook: " + wb.Name + ". Existing content is not modified outside the explicitly described operation.";
        }

        private static string DescribeAfter(ToolArguments a,string op)
        {
            switch(op)
            {
                case "sheet.create": return "Create worksheet '" + a.Get("name","") + "'.";
                case "sheet.rename": return "Rename '" + a.Get("sheet","") + "' to '" + a.Get("newName","") + "'.";
                case "sheet.delete": return "Delete worksheet '" + a.Get("sheet","") + "'.";
                case "sheet.copy": return "Copy worksheet '" + a.Get("sheet","") + "'.";
                case "sheet.move": return "Move worksheet '" + a.Get("sheet","") + "' to position " + a.Get("index","") + ".";
                case "sheet.visibility": return "Set worksheet visibility to " + a.Get("state","visible") + ".";
                case "range.clear": return "Clear " + a.Get("what","contents") + " from " + Target(a) + ".";
                case "range.insert": return "Insert " + a.Get("mode","cellsDown") + " at " + Target(a) + ".";
                case "range.delete": return "Delete " + a.Get("mode","cellsUp") + " at " + Target(a) + ".";
                case "range.merge": return (Bool(a,"merge",true)?"Merge ":"Unmerge ") + Target(a) + ".";
                case "range.autofill": return "AutoFill from " + a.Get("source","") + " into " + a.Get("destination","") + ".";
                case "row.height": return "Set/AutoFit row height for " + Target(a) + ".";
                case "column.width": return "Set/AutoFit column width for " + Target(a) + ".";
                case "row.visibility": return (Bool(a,"hidden",true)?"Hide ":"Show ") + "rows in " + Target(a) + ".";
                case "column.visibility": return (Bool(a,"hidden",true)?"Hide ":"Show ") + "columns in " + Target(a) + ".";
                case "table.create": return "Create Excel table over " + Target(a) + ".";
                case "table.resize": return "Resize table '" + a.Get("table","") + "' to " + a.Get("address","") + ".";
                case "table.rename": return "Rename table '" + a.Get("table","") + "' to '" + a.Get("newName","") + "'.";
                case "table.style": return "Update table style/options for '" + a.Get("table","") + "'.";
                case "table.delete": return "Remove table '" + a.Get("table","") + "'; keepData=" + a.Get("keepData","true") + ".";
                case "sort.apply": return "Sort " + Target(a) + " by key " + a.Get("key","") + " order=" + a.Get("order","asc") + ".";
                case "filter.apply": return "Apply filter field " + a.Get("field","") + " criteria=" + a.Get("criteria","") + ".";
                case "filter.clear": return "Clear active filters on worksheet '" + a.Get("sheet","") + "'.";
                case "validation.add": return "Add " + a.Get("type","list") + " validation to " + Target(a) + ".";
                case "validation.delete": return "Remove validation from " + Target(a) + ".";
                case "conditional_format.add": return "Add conditional formatting to " + Target(a) + ".";
                case "conditional_format.clear": return "Clear conditional formatting from " + Target(a) + ".";
                case "name.add": return "Create/update defined name '" + a.Get("name","") + "'.";
                case "name.delete": return "Delete defined name '" + a.Get("name","") + "'.";
                case "chart.create": return "Create chart from " + Target(a) + ".";
                case "chart.delete": return "Delete chart '" + a.Get("chart","") + "'.";
                case "chart.title": return "Set title of chart '" + a.Get("chart","") + "'.";
                case "pivot.refresh": return "Refresh PivotTable '" + a.Get("pivot","") + "'.";
                case "comment.set": return "Set note/comment on " + Target(a) + ".";
                case "comment.delete": return "Delete note/comment from " + Target(a) + ".";
                case "hyperlink.add": return "Add hyperlink to " + Target(a) + ".";
                case "view.freeze": return (Bool(a,"freeze",true)?"Freeze panes at ":"Unfreeze panes; target ") + a.Get("address","") + ".";
                case "view.zoom": return "Set Excel window zoom to " + a.Get("zoom","") + "%.";
                case "page_setup.set": return "Update page setup for worksheet '" + a.Get("sheet","") + "'.";
                default:return op;
            }
        }

        private static string WorkbookSummary(Excel.Workbook wb)
        {
            var sb=new StringBuilder();
            sb.AppendLine("Workbook="+wb.Name+"; readOnly="+wb.ReadOnly+"; worksheets="+wb.Worksheets.Count+
                          "; names="+wb.Names.Count+"; structureProtected="+wb.ProtectStructure);
            foreach(Excel.Worksheet ws in wb.Worksheets)
            {
                int charts=0,pivots=0;
                try{charts=((Excel.ChartObjects)ws.ChartObjects()).Count;}catch{}
                try{pivots=((Excel.PivotTables)ws.PivotTables()).Count;}catch{}
                sb.AppendLine("sheet="+ws.Name+"; visible="+ws.Visible+"; used="+ws.UsedRange.Address[false,false]+
                              "; tables="+ws.ListObjects.Count+"; charts="+charts+"; pivots="+pivots+"; shapes="+ws.Shapes.Count);
                if(sb.Length>MaxText) break;
            }
            return TextUtil.Truncate(sb.ToString(),MaxText);
        }

        private static string SheetInspect(Excel.Workbook wb,ToolArguments a)
        {
            var ws=Sheet(wb,a.Get("sheet",""));
            int charts=0,pivots=0; try{charts=((Excel.ChartObjects)ws.ChartObjects()).Count;}catch{} try{pivots=((Excel.PivotTables)ws.PivotTables()).Count;}catch{}
            var sb=new StringBuilder();
            sb.AppendLine("Sheet="+ws.Name+"; visible="+ws.Visible+"; used="+ws.UsedRange.Address[false,false]+
                "; rows="+ws.UsedRange.Rows.Count+"; columns="+ws.UsedRange.Columns.Count+"; tables="+ws.ListObjects.Count+
                "; charts="+charts+"; pivots="+pivots+"; shapes="+ws.Shapes.Count);
            try{sb.AppendLine("freezePanes="+wb.Application.ActiveWindow.FreezePanes+"; zoom="+wb.Application.ActiveWindow.Zoom);}catch{}
            foreach(Excel.ListObject t in ws.ListObjects) sb.AppendLine("table="+t.Name+"; range="+t.Range.Address[false,false]);
            return TextUtil.Truncate(sb.ToString(),MaxText);
        }

        private static string RangeInspect(Excel.Workbook wb,ToolArguments a)
        {
            var r=Range(wb,a); if(CellCount(r)>MaxInspectCells) r=r.Resize[Math.Max(1,MaxInspectCells/Math.Max(1,r.Columns.Count)),Math.Min(r.Columns.Count,32)];
            var sb=new StringBuilder(); sb.AppendLine("Range="+r.Worksheet.Name+"!"+r.Address[false,false]+"; cells="+CellCount(r));
            object vals=r.Value2, forms=r.Formula, fmts=r.NumberFormat;
            int rows=r.Rows.Count, cols=r.Columns.Count;
            for(int y=1;y<=rows;y++) for(int x=1;x<=cols;x++)
            {
                var cell=(Excel.Range)r.Cells[y,x];
                string v=One(vals,y,x), f=One(forms,y,x), nf=One(fmts,y,x);
                string comment=""; try{ if(cell.Comment!=null) comment=cell.Comment.Text(); }catch{}
                sb.AppendLine(cell.Address[false,false]+"; value="+Q(v)+"; formula="+Q(f)+"; numberFormat="+Q(nf)+
                              "; style="+Q(Convert.ToString(cell.Style))+"; note="+Q(TextUtil.Truncate(comment,200)));
                if(sb.Length>MaxText) return TextUtil.Truncate(sb.ToString(),MaxText);
            }
            return sb.ToString();
        }

        private static string TableInspect(Excel.Workbook wb,ToolArguments a)
        {
            var t=Table(wb,a); var sb=new StringBuilder();
            sb.AppendLine("Table="+t.Name+"; range="+t.Range.Address[false,false]+"; style="+t.TableStyle+
                          "; rows="+t.ListRows.Count+"; columns="+t.ListColumns.Count+"; totals="+t.ShowTotals+"; headers="+t.ShowHeaders);
            foreach(Excel.ListColumn c in t.ListColumns) sb.AppendLine("column="+c.Index+"; name="+c.Name);
            return TextUtil.Truncate(sb.ToString(),MaxText);
        }

        private static string NameInspect(Excel.Workbook wb,ToolArguments a)
        {
            string q=a.Get("query",""); var sb=new StringBuilder();
            foreach(Excel.Name n in wb.Names)
            {
                if(q.Length>0 && n.Name.IndexOf(q,StringComparison.OrdinalIgnoreCase)<0) continue;
                sb.AppendLine(n.Name+" = "+n.RefersTo+"; visible="+n.Visible);
                if(sb.Length>MaxText) break;
            }
            return sb.Length==0?"(no matching defined names)":sb.ToString();
        }

        private static string ChartInspect(Excel.Workbook wb,ToolArguments a)
        {
            var co=ChartObject(wb,a); var ch=co.Chart; var sb=new StringBuilder();
            var parent = co.Parent as Excel.Worksheet;
            sb.AppendLine("Chart="+co.Name+"; sheet="+(parent!=null?parent.Name:"?")+"; type="+ch.ChartType+"; hasTitle="+ch.HasTitle+
                          "; left="+co.Left+"; top="+co.Top+"; width="+co.Width+"; height="+co.Height);
            if(ch.HasTitle) sb.AppendLine("title="+ch.ChartTitle.Text);
            try
            {
                var series=(Excel.SeriesCollection)ch.SeriesCollection();
                sb.AppendLine("seriesCount="+series.Count);
                for(int i=1;i<=series.Count && i<=30;i++){var s=series.Item(i); sb.AppendLine("series="+i+"; name="+s.Name+"; formula="+s.Formula);}
            }catch{}
            return TextUtil.Truncate(sb.ToString(),MaxText);
        }

        private static string PivotInspect(Excel.Workbook wb,ToolArguments a)
        {
            var p=Pivot(wb,a); var sb=new StringBuilder();
            sb.AppendLine("Pivot="+p.Name+"; tableRange="+p.TableRange2.Address[false,false]+"; cacheIndex="+p.CacheIndex);
            try{foreach(Excel.PivotField f in (Excel.PivotFields)p.PivotFields()) sb.AppendLine("field="+f.Name+"; orientation="+f.Orientation+"; position="+f.Position);}catch{}
            return TextUtil.Truncate(sb.ToString(),MaxText);
        }

        private static string ValidationInspect(Excel.Workbook wb,ToolArguments a)
        {
            var r=Range(wb,a); var sb=new StringBuilder(); int emitted=0;
            foreach(Excel.Range cell in r.Cells)
            {
                try
                {
                    var v=cell.Validation;
                    if(v.Type!=(int)Excel.XlDVType.xlValidateInputOnly)
                        sb.AppendLine(cell.Address[false,false]+"; type="+v.Type+"; operator="+v.Operator+"; formula1="+v.Formula1+"; formula2="+v.Formula2+"; allowBlank="+v.IgnoreBlank);
                }catch{}
                if(++emitted>=MaxInspectCells) break;
            }
            return sb.Length==0?"(no validation detected in bounded range)":TextUtil.Truncate(sb.ToString(),MaxText);
        }

        private static string ConditionalInspect(Excel.Workbook wb,ToolArguments a)
        {
            var r=Range(wb,a); var sb=new StringBuilder();
            sb.AppendLine("Range="+r.Address[false,false]+"; conditionalFormatCount="+r.FormatConditions.Count);
            for(int i=1;i<=r.FormatConditions.Count && i<=50;i++)
            {
                var fc = r.FormatConditions.Item(i) as Excel.FormatCondition;
                if(fc==null){sb.AppendLine("rule="+i+"; type=(non-basic rule)");continue;}
                try{sb.AppendLine("rule="+i+"; type="+fc.Type+"; formula1="+fc.Formula1+"; formula2="+fc.Formula2);}
                catch{sb.AppendLine("rule="+i+"; type="+fc.Type);}
            }
            return TextUtil.Truncate(sb.ToString(),MaxText);
        }

        private static string FilterInspect(Excel.Workbook wb,ToolArguments a)
        {
            var ws=Sheet(wb,a.Get("sheet","")); var sb=new StringBuilder();
            sb.AppendLine("sheet="+ws.Name+"; autoFilterMode="+ws.AutoFilterMode+"; filterMode="+ws.FilterMode);
            string tableName=a.Get("table","");
            if(tableName.Length>0)
            {
                var t=ws.ListObjects[tableName]; sb.AppendLine("table="+t.Name+"; range="+t.Range.Address[false,false]+"; showAutoFilter="+t.ShowAutoFilter);
                try{var fs=t.AutoFilter.Filters; for(int i=1;i<=fs.Count;i++){var fl=fs.Item[i]; sb.AppendLine("field="+i+"; on="+fl.On);}}catch{}
            }
            return sb.ToString();
        }

        private static string PageSetupInspect(Excel.Workbook wb,ToolArguments a)
        {
            var ws=Sheet(wb,a.Get("sheet","")); var p=ws.PageSetup;
            return "sheet="+ws.Name+"; orientation="+p.Orientation+"; printArea="+p.PrintArea+"; printTitleRows="+p.PrintTitleRows+
                   "; printTitleColumns="+p.PrintTitleColumns+"; zoom="+p.Zoom+"; fitWide="+p.FitToPagesWide+"; fitTall="+p.FitToPagesTall+
                   "; margins(in points)="+p.LeftMargin+","+p.RightMargin+","+p.TopMargin+","+p.BottomMargin;
        }

        private static string ViewInspect(Excel.Application app)
        {
            var w=app.ActiveWindow; var ws=app.ActiveSheet as Excel.Worksheet; var sel=app.Selection as Excel.Range;
            return "sheet="+(ws!=null?ws.Name:"?")+"; selection="+(sel!=null?sel.Address[false,false]:"?")+"; zoom="+(w!=null?w.Zoom:0)+
                   "; freezePanes="+(w!=null&&w.FreezePanes)+"; splitRow="+(w!=null?w.SplitRow:0)+"; splitColumn="+(w!=null?w.SplitColumn:0);
        }

        private static void SheetCreate(Excel.Workbook wb,ToolArguments a)
        {
            string name=ValidSheetName(a.Get("name","Sheet"));
            if(Bool(a,"uniqueName",true)) name=UniqueSheetName(wb,name); else if(SheetExists(wb,name)) throw new InvalidOperationException("Worksheet already exists.");
            var ws=(Excel.Worksheet)wb.Worksheets.Add(After:wb.Sheets[wb.Sheets.Count]); ws.Name=name; ws.Activate();
        }
        private static void SheetRename(Excel.Workbook wb,ToolArguments a){string n=ValidSheetName(a.Get("newName","")); if(SheetExists(wb,n)) throw new InvalidOperationException("Target worksheet name already exists."); Sheet(wb,a.Get("sheet","")).Name=n;}
        private static void SheetCopy(Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet","")); ws.Copy(After:wb.Sheets[wb.Sheets.Count]); var copy=wb.Application.ActiveSheet as Excel.Worksheet; string n=a.Get("newName",""); if(copy!=null&&n.Length>0){n=ValidSheetName(n); if(SheetExists(wb,n)) throw new InvalidOperationException("Target worksheet name already exists."); copy.Name=n;}}
        private static void SheetMove(Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet","")); int i=Int(a,"index",1,wb.Sheets.Count); ws.Move(Before:wb.Sheets[i]);}
        private static void SheetVisibility(Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet","")); string s=a.Get("state","visible").ToLowerInvariant(); ws.Visible=s=="visible"?Excel.XlSheetVisibility.xlSheetVisible:s=="hidden"?Excel.XlSheetVisibility.xlSheetHidden:s=="veryhidden"?Excel.XlSheetVisibility.xlSheetVeryHidden:throw new ArgumentException("state must be visible, hidden or veryHidden.");}
        private static void RangeClear(Excel.Workbook wb,ToolArguments a){var r=Range(wb,a); switch(a.Get("what","contents").ToLowerInvariant()){case"contents":r.ClearContents();break;case"formats":r.ClearFormats();break;case"comments":r.ClearComments();break;case"all":r.Clear();break;default:throw new ArgumentException("what must be contents, formats, comments or all.");}}
        private static void RangeInsert(Excel.Workbook wb,ToolArguments a){var r=Range(wb,a); switch(a.Get("mode","cellsDown").ToLowerInvariant()){case"cellsdown":r.Insert(Excel.XlInsertShiftDirection.xlShiftDown);break;case"cellsright":r.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);break;case"rows":r.EntireRow.Insert();break;case"columns":r.EntireColumn.Insert();break;default:throw new ArgumentException("Invalid insert mode.");}}
        private static void RangeDelete(Excel.Workbook wb,ToolArguments a){var r=Range(wb,a); switch(a.Get("mode","cellsUp").ToLowerInvariant()){case"cellsup":r.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);break;case"cellsleft":r.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);break;case"rows":r.EntireRow.Delete();break;case"columns":r.EntireColumn.Delete();break;default:throw new ArgumentException("Invalid delete mode.");}}
        private static void RangeMerge(Excel.Workbook wb,ToolArguments a){var r=Range(wb,a); if(Bool(a,"merge",true)) r.Merge(); else r.UnMerge();}
        private static void RangeAutoFill(Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet","")); var src=ws.Range[a.Get("source","")]; var dst=ws.Range[a.Get("destination","")]; if(CellCount(dst)>MaxRangeCells)throw new ArgumentException("Destination too large."); src.AutoFill(dst,Excel.XlAutoFillType.xlFillDefault);}
        private static void RowHeight(Excel.Workbook wb,ToolArguments a){var rows=Range(wb,a).EntireRow; if(Bool(a,"autofit",false))rows.AutoFit(); else rows.RowHeight=Double(a,"height",5,409,15);}
        private static void ColumnWidth(Excel.Workbook wb,ToolArguments a){var cols=Range(wb,a).EntireColumn; if(Bool(a,"autofit",false))cols.AutoFit(); else cols.ColumnWidth=Double(a,"width",0,255,10);}
        private static void RowVisibility(Excel.Workbook wb,ToolArguments a){Range(wb,a).EntireRow.Hidden=Bool(a,"hidden",true);}
        private static void ColumnVisibility(Excel.Workbook wb,ToolArguments a){Range(wb,a).EntireColumn.Hidden=Bool(a,"hidden",true);}
        private static void TableCreate(Excel.Workbook wb,ToolArguments a){var r=Range(wb,a); var t=r.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,r,Type.Missing,Excel.XlYesNoGuess.xlYes,Type.Missing); string n=a.Get("name",""); if(n.Length>0)t.Name=n; string s=a.Get("style",""); if(s.Length>0)t.TableStyle=s;}
        private static void TableResize(Excel.Workbook wb,ToolArguments a){Table(wb,a).Resize(Range(wb,a));}
        private static void TableRename(Excel.Workbook wb,ToolArguments a){Table(wb,a).Name=a.Get("newName","");}
        private static void TableStyle(Excel.Workbook wb,ToolArguments a){var t=Table(wb,a); string s=a.Get("style",""); if(s.Length>0)t.TableStyle=s; if(a.Token("showTotals")!=null)t.ShowTotals=Bool(a,"showTotals",false); if(a.Token("showHeaders")!=null)t.ShowHeaders=Bool(a,"showHeaders",true); if(a.Token("bandedRows")!=null)t.ShowTableStyleRowStripes=Bool(a,"bandedRows",true);}
        private static void TableDelete(Excel.Workbook wb,ToolArguments a){var t=Table(wb,a); if(Bool(a,"keepData",true))t.Unlist(); else t.Delete();}
        private static void SortApply(Excel.Workbook wb,ToolArguments a){var r=Range(wb,a); var key=r.Worksheet.Range[a.Get("key","")]; r.Sort(Key1:key,Order1:a.Get("order","asc").Equals("desc",StringComparison.OrdinalIgnoreCase)?Excel.XlSortOrder.xlDescending:Excel.XlSortOrder.xlAscending,Header:Bool(a,"header",true)?Excel.XlYesNoGuess.xlYes:Excel.XlYesNoGuess.xlNo);}
        private static void FilterApply(Excel.Workbook wb,ToolArguments a){var r=Range(wb,a); int field=Int(a,"field",1,r.Columns.Count); r.AutoFilter(field,a.Get("criteria",""));}
        private static void FilterClear(Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet","")); try{if(ws.FilterMode)ws.ShowAllData();}catch{}}
        private static void ValidationAdd(Excel.Workbook wb,ToolArguments a)
        {
            var r=Range(wb,a); r.Validation.Delete(); string type=a.Get("type","list").ToLowerInvariant();
            Excel.XlDVType t=type=="list"?Excel.XlDVType.xlValidateList:type=="whole"?Excel.XlDVType.xlValidateWholeNumber:type=="decimal"?Excel.XlDVType.xlValidateDecimal:type=="date"?Excel.XlDVType.xlValidateDate:type=="textlength"?Excel.XlDVType.xlValidateTextLength:throw new ArgumentException("Unsupported validation type.");
            string op=a.Get("operator","between").ToLowerInvariant();
            Excel.XlFormatConditionOperator o=op=="between"?Excel.XlFormatConditionOperator.xlBetween:op=="notbetween"?Excel.XlFormatConditionOperator.xlNotBetween:op=="equal"?Excel.XlFormatConditionOperator.xlEqual:op=="notequal"?Excel.XlFormatConditionOperator.xlNotEqual:op=="greater"?Excel.XlFormatConditionOperator.xlGreater:op=="less"?Excel.XlFormatConditionOperator.xlLess:op=="greaterequal"?Excel.XlFormatConditionOperator.xlGreaterEqual:op=="lessequal"?Excel.XlFormatConditionOperator.xlLessEqual:Excel.XlFormatConditionOperator.xlBetween;
            r.Validation.Add(t,Excel.XlDVAlertStyle.xlValidAlertStop,(Excel.XlFormatConditionOperator)o,a.Get("formula1",""),a.Get("formula2","")); r.Validation.IgnoreBlank=Bool(a,"allowBlank",true); r.Validation.InCellDropdown=true;
        }
        private static void ConditionalAdd(Excel.Workbook wb,ToolArguments a)
        {
            var r=Range(wb,a); string type=a.Get("type","formula").ToLowerInvariant(); Excel.FormatCondition fc;
            if(type=="formula") fc=(Excel.FormatCondition)r.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression,Type.Missing,a.Get("formula1",""));
            else
            {
                string op=a.Get("operator","equal").ToLowerInvariant();
                Excel.XlFormatConditionOperator o=op=="between"?Excel.XlFormatConditionOperator.xlBetween:op=="greater"?Excel.XlFormatConditionOperator.xlGreater:op=="less"?Excel.XlFormatConditionOperator.xlLess:op=="notequal"?Excel.XlFormatConditionOperator.xlNotEqual:Excel.XlFormatConditionOperator.xlEqual;
                fc=(Excel.FormatCondition)r.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,o,a.Get("formula1",""),a.Get("formula2",""));
            }
            string fill=a.Get("fillColor",""); if(fill.Length>0)fc.Interior.Color=Ole(fill); string font=a.Get("fontColor",""); if(font.Length>0)fc.Font.Color=Ole(font);
        }
        private static void NameAdd(Excel.Workbook wb,ToolArguments a){string n=a.Get("name",""); if(n.Length==0)throw new ArgumentException("name is required."); string refers=a.Get("refersTo",""); if(!refers.StartsWith("="))throw new ArgumentException("refersTo must begin with '='."); try{wb.Names.Item(n).Delete();}catch{} wb.Names.Add(n,refers);}
        private static void ChartCreate(Excel.Workbook wb,ToolArguments a){var r=Range(wb,a); var ws=r.Worksheet; var co=((Excel.ChartObjects)ws.ChartObjects()).Add((float)Double(a,"left",0,5000,300),(float)Double(a,"top",0,5000,20),(float)Double(a,"width",80,2000,480),(float)Double(a,"height",60,1600,280)); co.Chart.SetSourceData(r); co.Chart.ChartType=ChartType(a.Get("chartType","column")); string title=a.Get("title",""); if(title.Length>0){co.Chart.HasTitle=true;co.Chart.ChartTitle.Text=title;} co.Activate();}
        private static void ChartTitle(Excel.Workbook wb,ToolArguments a){var ch=ChartObject(wb,a).Chart; string t=a.Get("title",""); ch.HasTitle=t.Length>0;if(t.Length>0)ch.ChartTitle.Text=t;}
        private static void PivotRefresh(Excel.Workbook wb,ToolArguments a){var p=Pivot(wb,a); p.PivotCache().Refresh();p.RefreshTable();}
        private static void CommentSet(Excel.Workbook wb,ToolArguments a){var c=RangeSingle(wb,a);c.ClearComments();c.AddComment(a.Get("text",""));}
        private static void HyperlinkAdd(Excel.Workbook wb,ToolArguments a){var c=RangeSingle(wb,a);string url=a.Get("url","");RequireHttpsOrMailto(url); c.Worksheet.Hyperlinks.Add(c,url,Type.Missing,Type.Missing,a.Get("text","").Length>0?a.Get("text",""):Convert.ToString(c.Text));}
        private static void ViewFreeze(Excel.Application app,Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet",""));ws.Activate();var w=app.ActiveWindow;w.FreezePanes=false;if(Bool(a,"freeze",true)){var address=a.Get("address","B2");app.Goto(ws.Range[address],true);ws.Range[address].Select();w.FreezePanes=true;}}
        private static void PageSetupSet(Excel.Workbook wb,ToolArguments a){var p=Sheet(wb,a.Get("sheet","")).PageSetup;string o=a.Get("orientation","");if(o.Length>0)p.Orientation=o.Equals("landscape",StringComparison.OrdinalIgnoreCase)?Excel.XlPageOrientation.xlLandscape:Excel.XlPageOrientation.xlPortrait;if(a.Token("fitToPagesWide")!=null||a.Token("fitToPagesTall")!=null)p.Zoom=false;if(a.Token("fitToPagesWide")!=null)p.FitToPagesWide=Int(a,"fitToPagesWide",1,100);if(a.Token("fitToPagesTall")!=null)p.FitToPagesTall=Int(a,"fitToPagesTall",1,100);if(a.Token("printArea")!=null)p.PrintArea=a.Get("printArea","");}

        private static Excel.Worksheet Sheet(Excel.Workbook wb,string name){if(string.IsNullOrWhiteSpace(name)){var ws=wb.Application.ActiveSheet as Excel.Worksheet;if(ws==null)throw new InvalidOperationException("No active worksheet.");return ws;}try{return (Excel.Worksheet)wb.Worksheets[name];}catch{throw new ArgumentException("Worksheet not found: "+name);}}
        private static Excel.Range Range(Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet",""));string ad=a.Get("address","");if(string.IsNullOrWhiteSpace(ad))throw new ArgumentException("address is required.");Excel.Range r;try{r=ws.Range[ad];}catch{throw new ArgumentException("Invalid Excel range address: "+ad);}if(r.Areas.Count!=1)throw new ArgumentException("Multi-area ranges are not supported.");return r;}
        private static Excel.Range RangeSingle(Excel.Workbook wb,ToolArguments a){var r=Range(wb,a);if(CellCount(r)!=1)throw new ArgumentException("This operation requires exactly one cell.");return r;}
        private static Excel.ListObject Table(Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet",""));string n=a.Get("table","");if(n.Length==0)throw new ArgumentException("table is required.");try{return ws.ListObjects[n];}catch{throw new ArgumentException("Table not found: "+n);}}
        private static Excel.ChartObject ChartObject(Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet",""));string n=a.Get("chart","");foreach(Excel.ChartObject co in (Excel.ChartObjects)ws.ChartObjects())if(string.Equals(co.Name,n,StringComparison.OrdinalIgnoreCase))return co;throw new ArgumentException("Chart not found: "+n);}
        private static Excel.PivotTable Pivot(Excel.Workbook wb,ToolArguments a){var ws=Sheet(wb,a.Get("sheet",""));string n=a.Get("pivot","");try{return (Excel.PivotTable)ws.PivotTables(n);}catch{throw new ArgumentException("PivotTable not found: "+n);}}
        private static long CellCount(Excel.Range r){try{return Convert.ToInt64(r.Cells.CountLarge);}catch{return Convert.ToInt64(r.Cells.Count);}}
        private static int Int(ToolArguments a,string k,int min,int max,int fallback=0){int v;if(!int.TryParse(a.Get(k,fallback.ToString(CultureInfo.InvariantCulture)),NumberStyles.Integer,CultureInfo.InvariantCulture,out v)||v<min||v>max)throw new ArgumentException(k+" must be "+min+".."+max+".");return v;}
        private static double Double(ToolArguments a,string k,double min,double max,double fallback){double v;if(!double.TryParse(a.Get(k,fallback.ToString(CultureInfo.InvariantCulture)),NumberStyles.Float,CultureInfo.InvariantCulture,out v)||double.IsNaN(v)||double.IsInfinity(v)||v<min||v>max)throw new ArgumentException(k+" must be "+min+".."+max+".");return v;}
        private static bool Bool(ToolArguments a,string k,bool fallback){var t=a.Token(k);if(t==null)return fallback;bool v;if(!bool.TryParse(t.ToString(),out v))throw new ArgumentException(k+" must be true or false.");return v;}
        private static string Target(ToolArguments a){return (a.Get("sheet","").Length>0?a.Get("sheet","")+"!":"")+a.Get("address","");}
        private static string Q(string s){return Newtonsoft.Json.JsonConvert.SerializeObject(TextUtil.Truncate(s??"",600));}
        private static string One(object v,int r,int c){var ar=v as Array;try{return Convert.ToString(ar==null?v:ar.GetValue(r,c))??"";}catch{return Convert.ToString(v)??"";}}
        private static bool SheetExists(Excel.Workbook wb,string n){foreach(object o in wb.Sheets){var ws=o as Excel.Worksheet;if(ws!=null&&string.Equals(ws.Name,n,StringComparison.OrdinalIgnoreCase))return true;var ch=o as Excel.Chart;if(ch!=null&&string.Equals(ch.Name,n,StringComparison.OrdinalIgnoreCase))return true;}return false;}
        private static string ValidSheetName(string n){n=(n??"").Trim();if(n.Length<1||n.Length>31||n.IndexOfAny(new[]{':','\\','/','?','*','[',']'})>=0)throw new ArgumentException("Invalid worksheet name.");return n;}
        private static string UniqueSheetName(Excel.Workbook wb,string n){if(!SheetExists(wb,n))return n;for(int i=2;i<1000;i++){string s=" ("+i+")";string b=n.Length>31-s.Length?n.Substring(0,31-s.Length):n;string c=b+s;if(!SheetExists(wb,c))return c;}throw new InvalidOperationException("Could not resolve unique sheet name.");}
        private static int Ole(string html){if(string.IsNullOrWhiteSpace(html)||html.Length!=7||html[0]!='#'||!html.Substring(1).All(Uri.IsHexDigit))throw new ArgumentException("Color must be #RRGGBB.");return System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml(html));}
        private static Excel.XlChartType ChartType(string s){switch((s??"").ToLowerInvariant()){case"bar":return Excel.XlChartType.xlBarClustered;case"line":return Excel.XlChartType.xlLine;case"pie":return Excel.XlChartType.xlPie;case"area":return Excel.XlChartType.xlArea;case"scatter":return Excel.XlChartType.xlXYScatter;default:return Excel.XlChartType.xlColumnClustered;}}
        private static void RequireHttpsOrMailto(string url){Uri u;if(!Uri.TryCreate(url,UriKind.Absolute,out u)||(!u.Scheme.Equals("https",StringComparison.OrdinalIgnoreCase)&&!u.Scheme.Equals("mailto",StringComparison.OrdinalIgnoreCase)))throw new ArgumentException("Hyperlinks must use https or mailto.");}
    }
}
