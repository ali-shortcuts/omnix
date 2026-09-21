using System;
using System.Globalization;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using OMNIX.Core.Context;
using OMNIX.Core.Errors;
using OMNIX.Core.Util;

namespace OMNIX.Core.Tools
{
    internal static class WordAdvancedCapabilities
    {
        private const int MaxText = 12000;
        private const int MaxItems = 120;

        public static string Inspect(WordHostAdapter adapter, ToolArguments args)
        {
            string op=Op(args);
            if(!OfficeCapabilityRegistry.IsKnown(HostType.Word,op,false))
                throw new OmnixException(ErrorCode.CORE_ERROR,"Unknown Word inspection operation: "+op,op,"Use list_office_capabilities.");
            var app=adapter.App; var doc=app.ActiveDocument;
            if(doc==null) throw new InvalidOperationException("No Word document is active.");
            switch(op)
            {
                case "document.summary": return DocumentSummary(doc);
                case "selection.inspect": return SelectionInspect(app);
                case "paragraph.inspect": return ParagraphInspect(doc,args);
                case "table.inspect": return TableInspect(doc,args);
                case "styles.inspect": return StylesInspect(doc,args);
                case "bookmarks.inspect": return BookmarksInspect(doc,args);
                case "fields.inspect": return FieldsInspect(doc,args);
                case "comments.inspect": return CommentsInspect(doc,args);
                case "revisions.inspect": return RevisionsInspect(doc,args);
                case "content_controls.inspect": return ContentControlsInspect(doc,args);
                case "sections.inspect": return SectionsInspect(doc,args);
                default: throw new InvalidOperationException("Unhandled Word inspection: "+op);
            }
        }

        public static WritePreview Prepare(WordHostAdapter adapter,string json)
        {
            var a=ToolArguments.Parse(json); string op=Op(a);
            if(!OfficeCapabilityRegistry.IsKnown(HostType.Word,op,true))
                throw new OmnixException(ErrorCode.CORE_ERROR,"Unknown Word write operation: "+op,op,"Use list_office_capabilities.");
            Validate(adapter,a,op);
            return new WritePreview
            {
                ToolName=ToolNames.ApplyOfficeCapability,
                Title="Word — "+op,
                Before=PreviewBefore(adapter,a,op),
                After=DescribeAfter(a,op),
                ArgumentsJson=json
            };
        }

        public static void Apply(WordHostAdapter adapter,string json)
        {
            var a=ToolArguments.Parse(json); string op=Op(a);
            if(!OfficeCapabilityRegistry.IsKnown(HostType.Word,op,true))
                throw new OmnixException(ErrorCode.CORE_ERROR,"Unknown Word write operation: "+op,op,"Use list_office_capabilities.");
            Validate(adapter,a,op);
            var app=adapter.App; var doc=app.ActiveDocument;
            if(doc==null) throw new InvalidOperationException("No Word document is active.");
            switch(op)
            {
                case "text.insert": TextInsert(app,doc,a); break;
                case "text.replace_range": TextReplace(doc,a); break;
                case "find_replace": FindReplace(doc,a); break;
                case "paragraph.format": ParagraphFormat(app,a); break;
                case "font.format": FontFormat(app,a); break;
                case "style.apply": StyleApply(app,a); break;
                case "table.create": TableCreate(app,doc,a); break;
                case "table.add_row": TableAddRow(doc,a); break;
                case "table.add_column": Table(doc,a).Columns.Add(); break;
                case "table.delete_row": Table(doc,a).Rows[Int(a,"row",1,Table(doc,a).Rows.Count)].Delete(); break;
                case "table.delete_column": Table(doc,a).Columns[Int(a,"column",1,Table(doc,a).Columns.Count)].Delete(); break;
                case "table.style": TableStyle(doc,a); break;
                case "bookmark.add": BookmarkAdd(app,doc,a); break;
                case "bookmark.delete": doc.Bookmarks[a.Get("name","")].Delete(); break;
                case "comment.add": CommentAdd(app,doc,a); break;
                case "comment.delete": doc.Comments[Int(a,"comment",1,doc.Comments.Count)].Delete(); break;
                case "track_changes": doc.TrackRevisions=Bool(a,"enabled",true); break;
                case "revisions.accept": Revisions(doc,app,a,true); break;
                case "revisions.reject": Revisions(doc,app,a,false); break;
                case "header_footer.set": HeaderFooter(doc,a); break;
                case "page_setup.set": PageSetup(doc,a); break;
                case "field.insert": FieldInsert(app,doc,a); break;
                case "hyperlink.add": HyperlinkAdd(app,doc,a); break;
                case "content_control.add": ContentControlAdd(app,doc,a); break;
                default: throw new InvalidOperationException("Unhandled Word write operation: "+op);
            }
        }

        private static void Validate(WordHostAdapter adapter,ToolArguments a,string op)
        {
            var doc=adapter.App.ActiveDocument;
            if(doc==null) throw new InvalidOperationException("No Word document is active.");
            if(op=="text.replace_range")
            {
                int start=Int(a,"start",doc.Content.Start,doc.Content.End); int end=Int(a,"end",start,doc.Content.End);
                if(end<start) throw new ArgumentException("end must be >= start.");
            }
            if(op=="table.create"){Int(a,"rows",1,100);Int(a,"columns",1,30);}
            if(op=="page_setup.set"||op=="header_footer.set") Int(a,"section",1,doc.Sections.Count);
            if(op=="comment.delete") Int(a,"comment",1,Math.Max(1,doc.Comments.Count));
            if(op=="bookmark.add"||op=="bookmark.delete")
            {
                string n=a.Get("name",""); if(string.IsNullOrWhiteSpace(n)||n.Length>40||char.IsDigit(n[0])||n.Any(ch=>!(char.IsLetterOrDigit(ch)||ch=='_')))
                    throw new ArgumentException("Bookmark name must start with a letter and contain only letters, numbers or underscore.");
            }
        }

        private static string DocumentSummary(Word.Document d)
        {
            int pages=0;try{pages=d.ComputeStatistics(Word.WdStatistic.wdStatisticPages,false);}catch{}
            return "Document="+d.Name+"; pages="+pages+"; sections="+d.Sections.Count+"; paragraphs="+d.Paragraphs.Count+
                   "; tables="+d.Tables.Count+"; styles="+d.Styles.Count+"; fields="+d.Fields.Count+"; bookmarks="+d.Bookmarks.Count+
                   "; comments="+d.Comments.Count+"; revisions="+d.Revisions.Count+"; contentControls="+d.ContentControls.Count+
                   "; inlineShapes="+d.InlineShapes.Count+"; shapes="+d.Shapes.Count+"; trackChanges="+d.TrackRevisions;
        }
        private static string SelectionInspect(Word.Application app)
        {
            var s=app.Selection;if(s==null)return "(no selection)"; string style="";try{style=Convert.ToString(s.get_Style());}catch{}
            int table=0;try{table=Convert.ToBoolean(s.Information[Word.WdInformation.wdWithInTable])?1:0;}catch{}
            return "selection=["+s.Start+","+s.End+"); text="+Q(TextUtil.Truncate(s.Text??"",3000))+"; style="+Q(style)+
                   "; alignment="+s.ParagraphFormat.Alignment+"; inTable="+table+"; font="+s.Font.Name+" "+s.Font.Size;
        }
        private static string ParagraphInspect(Word.Document d,ToolArguments a)
        {
            int i=Int(a,"index",1,d.Paragraphs.Count);var p=d.Paragraphs[i];string style="";try{style=Convert.ToString(p.Range.get_Style());}catch{}
            return "paragraph="+i+"; range=["+p.Range.Start+","+p.Range.End+"); style="+Q(style)+"; alignment="+p.Alignment+
                   "; text="+Q(TextUtil.Truncate(p.Range.Text??"",4000));
        }
        private static string TableInspect(Word.Document d,ToolArguments a)
        {
            var t=Table(d,a);var sb=new StringBuilder();sb.AppendLine("table="+a.Get("table","")+"; rows="+t.Rows.Count+"; columns="+t.Columns.Count);
            int n=0;for(int r=1;r<=t.Rows.Count;r++)for(int c=1;c<=t.Columns.Count;c++){string text="";try{text=t.Cell(r,c).Range.Text.TrimEnd('\r','\a');}catch{}sb.AppendLine("r="+r+",c="+c+"; text="+Q(TextUtil.Truncate(text,500)));if(++n>=120)return sb.ToString();}
            return TextUtil.Truncate(sb.ToString(),MaxText);
        }
        private static string StylesInspect(Word.Document d,ToolArguments a)
        {
            string q=a.Get("query","");var sb=new StringBuilder();int n=0;foreach(Word.Style s in d.Styles){string name="";try{name=s.NameLocal;}catch{}if(q.Length>0&&name.IndexOf(q,StringComparison.OrdinalIgnoreCase)<0)continue;sb.AppendLine("style="+name+"; type="+s.Type+"; builtIn="+s.BuiltIn);if(++n>=MaxItems)break;}return sb.Length==0?"(no matching styles)":sb.ToString();
        }
        private static string BookmarksInspect(Word.Document d,ToolArguments a)
        {
            string q=a.Get("query","");var sb=new StringBuilder();int n=0;foreach(Word.Bookmark b in d.Bookmarks){if(q.Length>0&&b.Name.IndexOf(q,StringComparison.OrdinalIgnoreCase)<0)continue;sb.AppendLine("bookmark="+b.Name+"; range=["+b.Range.Start+","+b.Range.End+")");if(++n>=MaxItems)break;}return sb.Length==0?"(no matching bookmarks)":sb.ToString();
        }
        private static string FieldsInspect(Word.Document d,ToolArguments a)
        {
            int offset=Int(a,"offset",0,Math.Max(0,d.Fields.Count));var sb=new StringBuilder();int end=Math.Min(d.Fields.Count,offset+50);for(int i=offset+1;i<=end;i++){var f=d.Fields[i];sb.AppendLine("field="+i+"; type="+f.Type+"; code="+Q(TextUtil.Truncate(f.Code.Text,600))+"; result="+Q(TextUtil.Truncate(f.Result.Text,600)));}sb.AppendLine("nextOffset="+(end<d.Fields.Count?end.ToString():"none"));return sb.ToString();
        }
        private static string CommentsInspect(Word.Document d,ToolArguments a)
        {
            int offset=Int(a,"offset",0,Math.Max(0,d.Comments.Count));var sb=new StringBuilder();int end=Math.Min(d.Comments.Count,offset+50);for(int i=offset+1;i<=end;i++){var c=d.Comments[i];sb.AppendLine("comment="+i+"; author="+Q(c.Author)+"; range=["+c.Scope.Start+","+c.Scope.End+"); text="+Q(TextUtil.Truncate(c.Range.Text,800)));}sb.AppendLine("nextOffset="+(end<d.Comments.Count?end.ToString():"none"));return sb.ToString();
        }
        private static string RevisionsInspect(Word.Document d,ToolArguments a)
        {
            int offset=Int(a,"offset",0,Math.Max(0,d.Revisions.Count));var sb=new StringBuilder();sb.AppendLine("trackChanges="+d.TrackRevisions+"; revisions="+d.Revisions.Count);int end=Math.Min(d.Revisions.Count,offset+50);for(int i=offset+1;i<=end;i++){var r=d.Revisions[i];sb.AppendLine("revision="+i+"; type="+r.Type+"; author="+Q(r.Author)+"; range=["+r.Range.Start+","+r.Range.End+"); text="+Q(TextUtil.Truncate(r.Range.Text,600)));}sb.AppendLine("nextOffset="+(end<d.Revisions.Count?end.ToString():"none"));return sb.ToString();
        }
        private static string ContentControlsInspect(Word.Document d,ToolArguments a)
        {
            int offset=Int(a,"offset",0,Math.Max(0,d.ContentControls.Count));var sb=new StringBuilder();int end=Math.Min(d.ContentControls.Count,offset+50);for(int i=offset+1;i<=end;i++){var c=d.ContentControls[i];sb.AppendLine("control="+i+"; type="+c.Type+"; title="+Q(c.Title)+"; tag="+Q(c.Tag)+"; range=["+c.Range.Start+","+c.Range.End+")");}sb.AppendLine("nextOffset="+(end<d.ContentControls.Count?end.ToString():"none"));return sb.ToString();
        }
        private static string SectionsInspect(Word.Document d,ToolArguments a)
        {
            int offset=Int(a,"offset",0,Math.Max(0,d.Sections.Count));var sb=new StringBuilder();int end=Math.Min(d.Sections.Count,offset+30);for(int i=offset+1;i<=end;i++){var s=d.Sections[i];var p=s.PageSetup;sb.AppendLine("section="+i+"; range=["+s.Range.Start+","+s.Range.End+"); orientation="+p.Orientation+"; margins="+p.TopMargin+","+p.BottomMargin+","+p.LeftMargin+","+p.RightMargin+"; headers="+s.Headers.Count+"; footers="+s.Footers.Count);}sb.AppendLine("nextOffset="+(end<d.Sections.Count?end.ToString():"none"));return sb.ToString();
        }

        private static string PreviewBefore(WordHostAdapter adapter,ToolArguments a,string op)
        {
            var app=adapter.App;var d=app.ActiveDocument;var s=app.Selection;
            if(op.StartsWith("table.",StringComparison.Ordinal)&&op!="table.create"){try{return TableInspect(d,a);}catch{}}
            if(op=="header_footer.set"||op=="page_setup.set"){try{return SectionsInspect(d,new ToolArgumentsProxy(a.Get("section","1")).Args);}catch{}}
            return s!=null?"Current selection ["+s.Start+","+s.End+"): "+Q(TextUtil.Truncate(s.Text??"",2000)):"Document="+d.Name;
        }
        private static string DescribeAfter(ToolArguments a,string op){return "Apply Word operation '"+op+"' with the confirmed arguments. Target remains inside the active document.";}

        private static void TextInsert(Word.Application app,Word.Document d,ToolArguments a)
        {
            string text=a.Get("text","");string where=a.Get("where","afterSelection").ToLowerInvariant();if(where=="beforeselection")app.Selection.Range.InsertBefore(text);else if(where=="afterselection")app.Selection.Range.InsertAfter(text);else if(where=="position"){int p=Int(a,"position",d.Content.Start,d.Content.End);d.Range(p,p).InsertAfter(text);}else throw new ArgumentException("where must be beforeSelection, afterSelection or position.");
        }
        private static void TextReplace(Word.Document d,ToolArguments a){int start=Int(a,"start",d.Content.Start,d.Content.End);int end=Int(a,"end",start,d.Content.End);d.Range(start,end).Text=a.Get("text","");}
        private static void FindReplace(Word.Document d,ToolArguments a)
        {
            string find=a.Get("find","");if(find.Length==0)throw new ArgumentException("find is required.");var r=d.Content.Duplicate;var f=r.Find;f.ClearFormatting();f.Replacement.ClearFormatting();f.Text=find;f.Replacement.Text=a.Get("replace","");f.Forward=true;f.Wrap=Word.WdFindWrap.wdFindContinue;f.MatchCase=Bool(a,"matchCase",false);f.MatchWholeWord=Bool(a,"wholeWord",false);f.Execute(Replace:Bool(a,"replaceAll",true)?Word.WdReplace.wdReplaceAll:Word.WdReplace.wdReplaceOne);
        }
        private static void ParagraphFormat(Word.Application app,ToolArguments a)
        {
            var p=app.Selection.ParagraphFormat;string al=a.Get("alignment","").ToLowerInvariant();if(al=="left")p.Alignment=Word.WdParagraphAlignment.wdAlignParagraphLeft;else if(al=="center")p.Alignment=Word.WdParagraphAlignment.wdAlignParagraphCenter;else if(al=="right")p.Alignment=Word.WdParagraphAlignment.wdAlignParagraphRight;else if(al=="justify")p.Alignment=Word.WdParagraphAlignment.wdAlignParagraphJustify;else if(al.Length>0)throw new ArgumentException("Unsupported alignment.");
            if(a.Token("spaceBefore")!=null)p.SpaceBefore=(float)Double(a,"spaceBefore",0,300,0);if(a.Token("spaceAfter")!=null)p.SpaceAfter=(float)Double(a,"spaceAfter",0,300,0);if(a.Token("leftIndent")!=null)p.LeftIndent=(float)Double(a,"leftIndent",-500,500,0);if(a.Token("rightIndent")!=null)p.RightIndent=(float)Double(a,"rightIndent",-500,500,0);if(a.Token("firstLineIndent")!=null)p.FirstLineIndent=(float)Double(a,"firstLineIndent",-500,500,0);if(a.Token("keepWithNext")!=null)p.KeepWithNext=Bool(a,"keepWithNext",false)?-1:0;
        }
        private static void FontFormat(Word.Application app,ToolArguments a)
        {
            var f=app.Selection.Font;string n=a.Get("fontName","");if(n.Length>0)f.Name=n;if(a.Token("fontSize")!=null)f.Size=(float)Double(a,"fontSize",6,96,11);if(a.Token("bold")!=null)f.Bold=Bool(a,"bold",false)?1:0;if(a.Token("italic")!=null)f.Italic=Bool(a,"italic",false)?1:0;if(a.Token("underline")!=null)f.Underline=Bool(a,"underline",false)?Word.WdUnderline.wdUnderlineSingle:Word.WdUnderline.wdUnderlineNone;string color=a.Get("fontColor","");if(color.Length>0)f.Color=(Word.WdColor)Ole(color);
        }
        private static void StyleApply(Word.Application app,ToolArguments a){object style=a.Get("style","");app.Selection.set_Style(ref style);}
        private static void TableCreate(Word.Application app,Word.Document d,ToolArguments a){var t=d.Tables.Add(app.Selection.Range,Int(a,"rows",1,100),Int(a,"columns",1,30));string style=a.Get("style","");if(style.Length>0){object tableStyle=style;t.set_Style(ref tableStyle);}t.Select();}
        private static void TableAddRow(Word.Document d,ToolArguments a){var t=Table(d,a);string p=a.Get("position","end");if(p.Equals("end",StringComparison.OrdinalIgnoreCase))t.Rows.Add();else{int idx=Int(a,"position",1,t.Rows.Count);Word.Row before=t.Rows[idx];t.Rows.Add(ref before);}}
        private static void TableStyle(Word.Document d,ToolArguments a){var t=Table(d,a);string s=a.Get("style","");if(s.Length>0){object tableStyle=s;t.set_Style(ref tableStyle);}string fit=a.Get("autofit","").ToLowerInvariant();if(fit=="content")t.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);else if(fit=="window")t.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);else if(fit=="fixed")t.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed);}
        private static void BookmarkAdd(Word.Application app,Word.Document d,ToolArguments a){string n=a.Get("name","");if(d.Bookmarks.Exists(n))d.Bookmarks[n].Delete();d.Bookmarks.Add(n,app.Selection.Range);}
        private static void CommentAdd(Word.Application app,Word.Document d,ToolArguments a){object text=a.Get("text","");d.Comments.Add(app.Selection.Range,ref text);}
        private static void Revisions(Word.Document d,Word.Application app,ToolArguments a,bool accept){string scope=a.Get("scope","selection").ToLowerInvariant();if(scope=="document"){if(accept)d.Revisions.AcceptAll();else d.Revisions.RejectAll();}else{var r=app.Selection.Range;if(accept)r.Revisions.AcceptAll();else r.Revisions.RejectAll();}}
        private static void HeaderFooter(Word.Document d,ToolArguments a){var s=d.Sections[Int(a,"section",1,d.Sections.Count)];var hf=a.Get("type","header").Equals("footer",StringComparison.OrdinalIgnoreCase)?s.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]:s.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];hf.Range.Text=a.Get("text","");}
        private static void PageSetup(Word.Document d,ToolArguments a){var p=d.Sections[Int(a,"section",1,d.Sections.Count)].PageSetup;string o=a.Get("orientation","");if(o.Length>0)p.Orientation=o.Equals("landscape",StringComparison.OrdinalIgnoreCase)?Word.WdOrientation.wdOrientLandscape:Word.WdOrientation.wdOrientPortrait;if(a.Token("topMargin")!=null)p.TopMargin=(float)Double(a,"topMargin",0,1000,72);if(a.Token("bottomMargin")!=null)p.BottomMargin=(float)Double(a,"bottomMargin",0,1000,72);if(a.Token("leftMargin")!=null)p.LeftMargin=(float)Double(a,"leftMargin",0,1000,72);if(a.Token("rightMargin")!=null)p.RightMargin=(float)Double(a,"rightMargin",0,1000,72);}
        private static void FieldInsert(Word.Application app,Word.Document d,ToolArguments a){string t=a.Get("type","page").ToLowerInvariant();Word.WdFieldType ft=t=="page"?Word.WdFieldType.wdFieldPage:t=="numpages"?Word.WdFieldType.wdFieldNumPages:t=="date"?Word.WdFieldType.wdFieldDate:t=="time"?Word.WdFieldType.wdFieldTime:t=="filename"?Word.WdFieldType.wdFieldFileName:t=="author"?Word.WdFieldType.wdFieldAuthor:throw new ArgumentException("Unsupported field type.");d.Fields.Add(app.Selection.Range,ft);}
        private static void HyperlinkAdd(Word.Application app,Word.Document d,ToolArguments a){string url=a.Get("url","");RequireHttpsOrMailto(url);string text=a.Get("text","");if(text.Length>0){var r=app.Selection.Range;r.Text=text;d.Hyperlinks.Add(r,url);}else d.Hyperlinks.Add(app.Selection.Range,url);}
        private static void ContentControlAdd(Word.Application app,Word.Document d,ToolArguments a){string t=a.Get("type","plainText").ToLowerInvariant();Word.WdContentControlType type=t=="richtext"?Word.WdContentControlType.wdContentControlRichText:t=="checkbox"?Word.WdContentControlType.wdContentControlCheckBox:Word.WdContentControlType.wdContentControlText;var c=d.ContentControls.Add(type,app.Selection.Range);c.Title=a.Get("title","");c.Tag=a.Get("tag","");}

        private static Word.Table Table(Word.Document d,ToolArguments a){int i=Int(a,"table",1,d.Tables.Count);return d.Tables[i];}
        private static int Int(ToolArguments a,string k,int min,int max,int fallback=0){int v;if(!int.TryParse(a.Get(k,fallback.ToString(CultureInfo.InvariantCulture)),NumberStyles.Integer,CultureInfo.InvariantCulture,out v)||v<min||v>max)throw new ArgumentException(k+" must be "+min+".."+max+".");return v;}
        private static double Double(ToolArguments a,string k,double min,double max,double fallback){double v;if(!double.TryParse(a.Get(k,fallback.ToString(CultureInfo.InvariantCulture)),NumberStyles.Float,CultureInfo.InvariantCulture,out v)||double.IsNaN(v)||double.IsInfinity(v)||v<min||v>max)throw new ArgumentException(k+" must be "+min+".."+max+".");return v;}
        private static bool Bool(ToolArguments a,string k,bool fallback){var t=a.Token(k);if(t==null)return fallback;bool v;if(!bool.TryParse(t.ToString(),out v))throw new ArgumentException(k+" must be true or false.");return v;}
        private static int Ole(string html){if(string.IsNullOrWhiteSpace(html)||html.Length!=7||html[0]!='#'||!html.Substring(1).All(Uri.IsHexDigit))throw new ArgumentException("Color must be #RRGGBB.");return System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml(html));}
        private static string Q(string s){return Newtonsoft.Json.JsonConvert.SerializeObject(s??"");}
        private static string Op(ToolArguments a){string op=(a.Get("operation","")??"").Trim();if(op.Length==0)throw new ArgumentException("operation is required.");return op;}
        private static void RequireHttpsOrMailto(string url){Uri u;if(!Uri.TryCreate(url,UriKind.Absolute,out u)||(!u.Scheme.Equals("https",StringComparison.OrdinalIgnoreCase)&&!u.Scheme.Equals("mailto",StringComparison.OrdinalIgnoreCase)))throw new ArgumentException("Hyperlinks must use https or mailto.");}

        // Small adapter only to reuse sections inspection in a preview without exposing JObject.
        private sealed class ToolArgumentsProxy
        {
            public ToolArguments Args { get; private set; }
            public ToolArgumentsProxy(string section){Args=ToolArguments.Parse("{\"offset\":"+Math.Max(0,int.Parse(section)-1).ToString(CultureInfo.InvariantCulture)+"}");}
        }
    }
}
