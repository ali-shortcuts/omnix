using System;
using System.Globalization;
using Newtonsoft.Json.Linq;
using OMNIX.Core.Context;
using Excel=Microsoft.Office.Interop.Excel;
using Word=Microsoft.Office.Interop.Word;
using Ppt=Microsoft.Office.Interop.PowerPoint;

namespace OMNIX.Core.Agent
{
    public static class OfficePostconditions
    {
        private static int Index(JObject c,string name,int fallback,int max)
        {
            int n=c[name]==null?fallback:(int)c[name];
            if(n<1 || n>max) throw new ArgumentException("Invalid "+name);
            return n;
        }
        public static bool ValuesEqual(object actual,JToken expected)
        {
            if(expected==null || expected.Type==JTokenType.Null) return actual==null;
            if(expected.Type==JTokenType.Boolean) return actual is bool && (bool)actual==(bool)expected;
            if(expected.Type==JTokenType.Integer || expected.Type==JTokenType.Float)
            {
                if(!(actual is double || actual is int || actual is decimal || actual is float || actual is long)) return false;
                double a=Convert.ToDouble(actual,CultureInfo.InvariantCulture), b=(double)expected;
                return !double.IsNaN(a) && !double.IsInfinity(a) && Math.Abs(a-b)<=1e-9*Math.Max(1,Math.Abs(b));
            }
            return actual is string && string.Equals((string)actual,(string)expected,StringComparison.Ordinal);
        }
        public static string ExcelCheck(Excel.Application app,JObject c)
        {
            ExecutionPlan.ValidateCheck(c,HostType.Excel);
            var sheet=(Excel.Worksheet)app.ActiveWorkbook.Worksheets[(string)c["sheet"]];
            var range=sheet.Range[(string)c["address"]];
            if(range.Areas.Count!=1 || Convert.ToInt64(range.Cells.CountLarge)>512) return "Check range must be contiguous and at most 512 cells.";
            string kind=(string)c["kind"];
            if(kind=="cell_value" || kind=="formula")
            {
                if(Convert.ToInt64(range.Cells.CountLarge)!=1) return "Value/formula checks require one cell.";
                if(kind=="formula")
                {
                    if(!Convert.ToBoolean(range.HasFormula)) return "Expected a real formula at "+(string)c["address"];
                    range.Calculate();
                    if(c["formula"]!=null && !string.Equals(Convert.ToString(range.Formula), (string)c["formula"],StringComparison.OrdinalIgnoreCase))
                        return "Formula reference differs from the plan at "+(string)c["address"];
                }
                return ValuesEqual(range.Value2,c["value"])?null:"Computed value/type differs from the plan at "+(string)c["address"];
            }
            if(kind=="format")
            {
                if(c["bold"]!=null && !ValuesEqual(range.Font.Bold,c["bold"])) return "Bold differs from the plan.";
                if(c["italic"]!=null && !ValuesEqual(range.Font.Italic,c["italic"])) return "Italic differs from the plan.";
                if(c["wrapText"]!=null && !ValuesEqual(range.WrapText,c["wrapText"])) return "Text wrapping differs from the plan.";
                if(c["fontSize"]!=null && !ValuesEqual(range.Font.Size,c["fontSize"])) return "Font size differs from the plan.";
                if(c["numberFormat"]!=null && !ValuesEqual(range.NumberFormat,c["numberFormat"])) return "Number format differs from the plan.";
                if(c["horizontalAlignment"]!=null)
                {
                    string alignment=(string)c["horizontalAlignment"];
                    int expected=alignment=="center"?(int)Excel.XlHAlign.xlHAlignCenter:alignment=="right"?(int)Excel.XlHAlign.xlHAlignRight:alignment=="left"?(int)Excel.XlHAlign.xlHAlignLeft:(int)Excel.XlHAlign.xlHAlignGeneral;
                    if(!ValuesEqual(range.HorizontalAlignment,new JValue(expected))) return "Horizontal alignment differs from the plan.";
                }
                return null;
            }
            if(kind=="heading")
            {
                var first=(Excel.Range)range.Cells[1,1];
                if(!Convert.ToBoolean(first.MergeCells) || first.MergeArea.Address[false,false]!=range.Address[false,false])
                    return "Heading frame does not match the planned range.";
                if(!ValuesEqual(first.Value2,c["text"])) return "Heading text differs from the plan.";
                foreach(Excel.ListObject table in sheet.ListObjects)
                    if(app.Intersect(range,table.Range)!=null) return "Heading overlaps a data table.";
                return null;
            }
            if(kind=="table")
            {
                foreach(Excel.ListObject table in sheet.ListObjects)
                    if(table.Range.Address[false,false]==range.Address[false,false]) return null;
                return "No native table matches the planned range.";
            }
            foreach(Excel.Range cell in range.Cells)
            {
                if(Convert.ToBoolean(app.WorksheetFunction.IsError(cell))) return "Excel error at "+cell.Address[false,false];
            }
            return null;
        }
        public static string WordCheck(Word.Application app,JObject c)
        {
            ExecutionPlan.ValidateCheck(c,HostType.Word);
            var doc=app.ActiveDocument; string kind=(string)c["kind"];
            if(kind=="table_count") return doc.Tables.Count==(int)c["count"]?null:"Word table count differs from the plan.";
            int index=Index(c,"paragraph",1,doc.Paragraphs.Count);
            var p=doc.Paragraphs[index];
            if(kind=="paragraph_style")
            {
                var style=p.Range.get_Style() as Word.Style;
                return style!=null && style.NameLocal==(string)c["style"]?null:"Paragraph style differs from the plan.";
            }
            var r=p.Range.Duplicate;
            if(r.End-r.Start>10000) return "Paragraph exceeds the bounded text check; choose a smaller target.";
            return (r.Text??"").Contains((string)c["text"])?null:"Expected text is absent from the target paragraph.";
        }
        public static string PowerPointCheck(Ppt.Application app,JObject c)
        {
            ExecutionPlan.ValidateCheck(c,HostType.PowerPoint);
            var p=app.ActivePresentation; string kind=(string)c["kind"];
            if(kind=="slide_count") return p.Slides.Count==(int)c["count"]?null:"Slide count differs from the plan.";
            var slide=p.Slides[Index(c,"slide",1,p.Slides.Count)];
            var shape=slide.Shapes[Index(c,"shape",1,slide.Shapes.Count)];
            if(kind=="shape_bounds") return shape.Left>=0 && shape.Top>=0 && shape.Left+shape.Width<=p.PageSetup.SlideWidth+1 && shape.Top+shape.Height<=p.PageSetup.SlideHeight+1?null:"Shape extends outside the slide.";
            if(shape.HasTextFrame!=Microsoft.Office.Core.MsoTriState.msoTrue) return "Target shape has no text frame.";
            var text=shape.TextFrame.TextRange;
            if(text.Length>10000) return "Shape exceeds the bounded text check.";
            return (text.Text??"").Contains((string)c["text"])?null:"Expected text is absent from the target shape.";
        }
    }
}
