using System;
using System.Globalization;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using OMNIX.Core.Context;
using OMNIX.Core.Errors;
using OMNIX.Core.Util;

namespace OMNIX.Core.Tools
{
    internal static class PowerPointAdvancedCapabilities
    {
        private const int MaxText=12000;

        public static string Inspect(PowerPointHostAdapter adapter,ToolArguments a)
        {
            string op=Op(a);if(!OfficeCapabilityRegistry.IsKnown(HostType.PowerPoint,op,false))throw new OmnixException(ErrorCode.CORE_ERROR,"Unknown PowerPoint inspection operation: "+op,op,"Use list_office_capabilities.");
            var app=adapter.App;var p=app.ActivePresentation;if(p==null)throw new InvalidOperationException("No PowerPoint presentation is active.");
            switch(op)
            {
                case "presentation.summary":return PresentationSummary(p);
                case "slide.inspect":return SlideInspect(p,a);
                case "shape.inspect":return ShapeInspect(p,a);
                case "table.inspect":return TableInspect(p,a);
                case "notes.inspect":return NotesInspect(p,a);
                case "animations.inspect":return AnimationsInspect(p,a);
                case "sections.inspect":return SectionsInspect(p);
                default:throw new InvalidOperationException("Unhandled PowerPoint inspection: "+op);
            }
        }

        public static WritePreview Prepare(PowerPointHostAdapter adapter,string json)
        {
            var a=ToolArguments.Parse(json);string op=Op(a);if(!OfficeCapabilityRegistry.IsKnown(HostType.PowerPoint,op,true))throw new OmnixException(ErrorCode.CORE_ERROR,"Unknown PowerPoint write operation: "+op,op,"Use list_office_capabilities.");
            Validate(adapter,a,op);
            return new WritePreview{ToolName=ToolNames.ApplyOfficeCapability,Title="PowerPoint — "+op,Before=PreviewBefore(adapter,a,op),After="Apply PowerPoint operation '"+op+"' to the confirmed active presentation.",ArgumentsJson=json};
        }

        public static void Apply(PowerPointHostAdapter adapter,string json)
        {
            var a=ToolArguments.Parse(json);string op=Op(a);if(!OfficeCapabilityRegistry.IsKnown(HostType.PowerPoint,op,true))throw new OmnixException(ErrorCode.CORE_ERROR,"Unknown PowerPoint write operation: "+op,op,"Use list_office_capabilities.");
            Validate(adapter,a,op);
            var app=adapter.App;var p=app.ActivePresentation;if(p==null)throw new InvalidOperationException("No PowerPoint presentation is active.");
            switch(op)
            {
                case "slide.delete":Slide(p,a).Delete();break;
                case "slide.duplicate":SlideDuplicate(p,a);break;
                case "slide.move":Slide(p,a).MoveTo(Int(a,"index",1,p.Slides.Count));break;
                case "slide.layout":Slide(p,a).CustomLayout=p.SlideMaster.CustomLayouts[Int(a,"layout",1,p.SlideMaster.CustomLayouts.Count)];break;
                case "shape.add_textbox":AddTextBox(p,a);break;
                case "shape.add_shape":AddShape(p,a);break;
                case "shape.delete":Shape(p,a).Delete();break;
                case "shape.move_resize":MoveResize(p,a);break;
                case "shape.rotate":Shape(p,a).Rotation=(float)Double(a,"rotation",-360,360,0);break;
                case "shape.fill":ShapeFill(p,a);break;
                case "shape.line":ShapeLine(p,a);break;
                case "text.set":TextSet(p,a);break;
                case "text.format":TextFormat(p,a);break;
                case "table.create":TableCreate(p,a);break;
                case "table.set_cell":TableSetCell(p,a);break;
                case "notes.set":NotesSet(p,a);break;
                case "hyperlink.add":HyperlinkAdd(p,a);break;
                case "transition.set":TransitionSet(p,a);break;
                case "animation.add":AnimationAdd(p,a);break;
                case "animation.clear":AnimationClear(p,a);break;
                case "view.goto_slide":if(app.ActiveWindow!=null)app.ActiveWindow.View.GotoSlide(Int(a,"slide",1,p.Slides.Count));break;
                default:throw new InvalidOperationException("Unhandled PowerPoint write operation: "+op);
            }
        }

        private static void Validate(PowerPointHostAdapter adapter,ToolArguments a,string op)
        {
            var p=adapter.App.ActivePresentation;if(p==null)throw new InvalidOperationException("No PowerPoint presentation is active.");
            if(op!="shape.add_textbox"&&op!="shape.add_shape"&&op!="table.create"&&op!="view.goto_slide"&&!op.StartsWith("slide.",StringComparison.Ordinal)) { if(a.Token("slide")!=null) Int(a,"slide",1,p.Slides.Count); }
            if(op.StartsWith("slide.",StringComparison.Ordinal)||op=="view.goto_slide") Int(a,"slide",1,p.Slides.Count);
            if(op.StartsWith("shape.",StringComparison.Ordinal)||op.StartsWith("text.",StringComparison.Ordinal)||op=="hyperlink.add"||op=="animation.add")
            {
                if(op!="shape.add_textbox"&&op!="shape.add_shape"){var s=Slide(p,a);Int(a,"shape",1,s.Shapes.Count);}
            }
            if(op=="table.create"){Int(a,"rows",1,40);Int(a,"columns",1,20);}
            if(op=="table.set_cell"){var sh=Shape(p,a);if(sh.HasTable!=Office.MsoTriState.msoTrue)throw new ArgumentException("Target shape is not a table.");Int(a,"row",1,sh.Table.Rows.Count);Int(a,"column",1,sh.Table.Columns.Count);}
        }

        private static string PresentationSummary(Ppt.Presentation p)
        {
            int sections=0;try{sections=p.SectionProperties.Count;}catch{}
            return "Presentation="+p.Name+"; slides="+p.Slides.Count+"; sections="+sections+"; layouts="+p.SlideMaster.CustomLayouts.Count+
                   "; pageSize="+p.PageSetup.SlideWidth+"x"+p.PageSetup.SlideHeight;
        }
        private static string SlideInspect(Ppt.Presentation p,ToolArguments a)
        {
            var s=Slide(p,a);var sb=new StringBuilder();sb.AppendLine("slide="+s.SlideIndex+"; id="+s.SlideID+"; name="+s.Name+"; layout="+(s.CustomLayout!=null?s.CustomLayout.Name:"?")+"; shapes="+s.Shapes.Count+"; hidden="+s.SlideShowTransition.Hidden);
            for(int i=1;i<=s.Shapes.Count&&i<=100;i++){var sh=s.Shapes[i];sb.AppendLine("shape="+i+"; name="+sh.Name+"; type="+sh.Type+"; left="+sh.Left+"; top="+sh.Top+"; width="+sh.Width+"; height="+sh.Height+"; text="+Q(ShapeText(sh,300)));if(sb.Length>MaxText)break;}
            return TextUtil.Truncate(sb.ToString(),MaxText);
        }
        private static string ShapeInspect(Ppt.Presentation p,ToolArguments a)
        {
            var sh=Shape(p,a);var sb=new StringBuilder();sb.AppendLine("name="+sh.Name+"; type="+sh.Type+"; left="+sh.Left+"; top="+sh.Top+"; width="+sh.Width+"; height="+sh.Height+"; rotation="+sh.Rotation+"; text="+Q(ShapeText(sh,3000)));
            try{sb.AppendLine("fillVisible="+sh.Fill.Visible+"; fillRGB="+sh.Fill.ForeColor.RGB+"; lineVisible="+sh.Line.Visible+"; lineRGB="+sh.Line.ForeColor.RGB+"; lineWeight="+sh.Line.Weight);}catch{}
            try{if(sh.HasTable==Office.MsoTriState.msoTrue)sb.AppendLine("table="+sh.Table.Rows.Count+"x"+sh.Table.Columns.Count);}catch{}
            try{if(sh.Type==Office.MsoShapeType.msoGroup)sb.AppendLine("groupItems="+sh.GroupItems.Count);}catch{}
            return sb.ToString();
        }
        private static string TableInspect(Ppt.Presentation p,ToolArguments a)
        {
            var sh=Shape(p,a);if(sh.HasTable!=Office.MsoTriState.msoTrue)throw new ArgumentException("Target shape is not a table.");var t=sh.Table;var sb=new StringBuilder();sb.AppendLine("table="+t.Rows.Count+"x"+t.Columns.Count);int n=0;for(int r=1;r<=t.Rows.Count;r++)for(int c=1;c<=t.Columns.Count;c++){string text="";try{text=t.Cell(r,c).Shape.TextFrame.TextRange.Text;}catch{}sb.AppendLine("r="+r+",c="+c+"; text="+Q(TextUtil.Truncate(text,500)));if(++n>=120)return sb.ToString();}return TextUtil.Truncate(sb.ToString(),MaxText);
        }
        private static string NotesInspect(Ppt.Presentation p,ToolArguments a){return "slide="+Int(a,"slide",1,p.Slides.Count)+"; notes="+Q(TextUtil.Truncate(GetNotes(Slide(p,a)),5000));}
        private static string AnimationsInspect(Ppt.Presentation p,ToolArguments a)
        {
            var s=Slide(p,a);var seq=s.TimeLine.MainSequence;var sb=new StringBuilder();sb.AppendLine("slide="+s.SlideIndex+"; effects="+seq.Count);for(int i=1;i<=seq.Count&&i<=100;i++){var e=seq[i];string shape="";try{shape=e.Shape.Name;}catch{}sb.AppendLine("effect="+i+"; type="+e.EffectType+"; trigger="+e.Timing.TriggerType+"; shape="+shape);}return sb.ToString();
        }
        private static string SectionsInspect(Ppt.Presentation p)
        {
            var sb=new StringBuilder();try{var sp=p.SectionProperties;sb.AppendLine("sections="+sp.Count);for(int i=1;i<=sp.Count&&i<=100;i++)sb.AppendLine("section="+i+"; name="+sp.Name(i)+"; firstSlide="+sp.FirstSlide(i)+"; slides="+sp.SlidesCount(i));}catch(Exception ex){sb.AppendLine("Section metadata unavailable: "+ex.Message);}return sb.ToString();
        }

        private static string PreviewBefore(PowerPointHostAdapter adapter,ToolArguments a,string op)
        {
            var p=adapter.App.ActivePresentation;if(p==null)return "(no presentation)";
            try{if(a.Token("shape")!=null)return ShapeInspect(p,a);if(a.Token("slide")!=null)return SlideInspect(p,a);}catch{}
            return "Presentation="+p.Name+"; slides="+p.Slides.Count;
        }

        private static void SlideDuplicate(Ppt.Presentation p,ToolArguments a){var range=Slide(p,a).Duplicate();int index=0;if(int.TryParse(a.Get("index",""),out index)&&index>=1&&index<=p.Slides.Count)range.MoveTo(index);}
        private static void AddTextBox(Ppt.Presentation p,ToolArguments a){var s=Slide(p,a);var sh=s.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,F(a,"left",0,5000,50),F(a,"top",0,5000,50),F(a,"width",10,5000,300),F(a,"height",10,5000,100));sh.TextFrame.TextRange.Text=a.Get("text","");sh.Select();}
        private static void AddShape(Ppt.Presentation p,ToolArguments a){var s=Slide(p,a);string type=a.Get("type","rectangle").ToLowerInvariant();Ppt.Shape sh;if(type=="line")sh=s.Shapes.AddLine(F(a,"left",0,5000,50),F(a,"top",0,5000,50),F(a,"left",0,5000,50)+F(a,"width",1,5000,200),F(a,"top",0,5000,50)+F(a,"height",1,5000,1));else{Office.MsoAutoShapeType t=type=="ellipse"?Office.MsoAutoShapeType.msoShapeOval:type=="roundedrectangle"?Office.MsoAutoShapeType.msoShapeRoundedRectangle:Office.MsoAutoShapeType.msoShapeRectangle;sh=s.Shapes.AddShape(t,F(a,"left",0,5000,50),F(a,"top",0,5000,50),F(a,"width",1,5000,200),F(a,"height",1,5000,100));}sh.Select();}
        private static void MoveResize(Ppt.Presentation p,ToolArguments a){var sh=Shape(p,a);if(a.Token("left")!=null)sh.Left=F(a,"left",-5000,10000,sh.Left);if(a.Token("top")!=null)sh.Top=F(a,"top",-5000,10000,sh.Top);if(a.Token("width")!=null)sh.Width=F(a,"width",1,10000,sh.Width);if(a.Token("height")!=null)sh.Height=F(a,"height",1,10000,sh.Height);}
        private static void ShapeFill(Ppt.Presentation p,ToolArguments a){var sh=Shape(p,a);if(Bool(a,"none",false)){sh.Fill.Visible=Office.MsoTriState.msoFalse;return;}sh.Fill.Visible=Office.MsoTriState.msoTrue;sh.Fill.Solid();string c=a.Get("color","");if(c.Length>0)sh.Fill.ForeColor.RGB=Ole(c);if(a.Token("transparency")!=null)sh.Fill.Transparency=F(a,"transparency",0,1,0);}
        private static void ShapeLine(Ppt.Presentation p,ToolArguments a){var sh=Shape(p,a);bool vis=Bool(a,"visible",true);sh.Line.Visible=vis?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;if(!vis)return;string c=a.Get("color","");if(c.Length>0)sh.Line.ForeColor.RGB=Ole(c);if(a.Token("weight")!=null)sh.Line.Weight=F(a,"weight",0.25f,20,1);}
        private static void TextSet(Ppt.Presentation p,ToolArguments a){var sh=Shape(p,a);if(sh.HasTextFrame!=Office.MsoTriState.msoTrue)throw new ArgumentException("Shape has no text frame.");sh.TextFrame.TextRange.Text=a.Get("text","");}
        private static void TextFormat(Ppt.Presentation p,ToolArguments a){var sh=Shape(p,a);if(sh.HasTextFrame!=Office.MsoTriState.msoTrue)throw new ArgumentException("Shape has no text frame.");var tr=sh.TextFrame.TextRange;string n=a.Get("fontName","");if(n.Length>0)tr.Font.Name=n;if(a.Token("fontSize")!=null)tr.Font.Size=F(a,"fontSize",6,120,24);if(a.Token("bold")!=null)tr.Font.Bold=Bool(a,"bold",false)?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;if(a.Token("italic")!=null)tr.Font.Italic=Bool(a,"italic",false)?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;string c=a.Get("fontColor","");if(c.Length>0)tr.Font.Color.RGB=Ole(c);string al=a.Get("alignment","").ToLowerInvariant();if(al=="left")tr.ParagraphFormat.Alignment=Ppt.PpParagraphAlignment.ppAlignLeft;else if(al=="center")tr.ParagraphFormat.Alignment=Ppt.PpParagraphAlignment.ppAlignCenter;else if(al=="right")tr.ParagraphFormat.Alignment=Ppt.PpParagraphAlignment.ppAlignRight;else if(al=="justify")tr.ParagraphFormat.Alignment=Ppt.PpParagraphAlignment.ppAlignJustify;}
        private static void TableCreate(Ppt.Presentation p,ToolArguments a){var s=Slide(p,a);var sh=s.Shapes.AddTable(Int(a,"rows",1,40),Int(a,"columns",1,20),F(a,"left",0,5000,50),F(a,"top",0,5000,100),F(a,"width",20,5000,600),F(a,"height",20,5000,300));sh.Select();}
        private static void TableSetCell(Ppt.Presentation p,ToolArguments a){var sh=Shape(p,a);sh.Table.Cell(Int(a,"row",1,sh.Table.Rows.Count),Int(a,"column",1,sh.Table.Columns.Count)).Shape.TextFrame.TextRange.Text=a.Get("text","");}
        private static void NotesSet(Ppt.Presentation p,ToolArguments a){var body=NotesBody(Slide(p,a));if(body==null)throw new InvalidOperationException("Speaker-notes body is unavailable.");body.TextFrame.TextRange.Text=a.Get("text","");}
        private static void HyperlinkAdd(Ppt.Presentation p,ToolArguments a){string url=a.Get("url","");RequireHttpsOrMailto(url);var sh=Shape(p,a);var action=sh.ActionSettings[Ppt.PpMouseActivation.ppMouseClick];action.Action=Ppt.PpActionType.ppActionHyperlink;action.Hyperlink.Address=url;}
        private static void TransitionSet(Ppt.Presentation p,ToolArguments a){var t=Slide(p,a).SlideShowTransition;string type=a.Get("type","none").ToLowerInvariant();t.EntryEffect=type=="fade"?Ppt.PpEntryEffect.ppEffectFade:type=="push"?Ppt.PpEntryEffect.ppEffectPushLeft:type=="wipe"?Ppt.PpEntryEffect.ppEffectWipeRight:Ppt.PpEntryEffect.ppEffectNone;t.AdvanceOnClick=Bool(a,"advanceOnClick",true)?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;if(a.Token("advanceAfterSeconds")!=null){t.AdvanceOnTime=Office.MsoTriState.msoTrue;t.AdvanceTime=F(a,"advanceAfterSeconds",0,3600,0);}else t.AdvanceOnTime=Office.MsoTriState.msoFalse;}
        private static void AnimationAdd(Ppt.Presentation p,ToolArguments a){var s=Slide(p,a);var sh=Shape(p,a);string type=a.Get("type","appear").ToLowerInvariant();Office.MsoAnimEffect e=type=="fade"?Office.MsoAnimEffect.msoAnimEffectFade:type=="fly"?Office.MsoAnimEffect.msoAnimEffectFly:Office.MsoAnimEffect.msoAnimEffectAppear;s.TimeLine.MainSequence.AddEffect(sh,e,Office.MsoAnimateByLevel.msoAnimateLevelNone,Office.MsoAnimTriggerType.msoAnimTriggerOnPageClick);}
        private static void AnimationClear(Ppt.Presentation p,ToolArguments a){var seq=Slide(p,a).TimeLine.MainSequence;for(int i=seq.Count;i>=1;i--)seq[i].Delete();}

        private static Ppt.Slide Slide(Ppt.Presentation p,ToolArguments a){return p.Slides[Int(a,"slide",1,p.Slides.Count)];}
        private static Ppt.Shape Shape(Ppt.Presentation p,ToolArguments a){var s=Slide(p,a);return s.Shapes[Int(a,"shape",1,s.Shapes.Count)];}
        private static string ShapeText(Ppt.Shape sh,int max){try{if(sh.HasTextFrame==Office.MsoTriState.msoTrue&&sh.TextFrame.HasText==Office.MsoTriState.msoTrue)return TextUtil.Truncate(sh.TextFrame.TextRange.Text??"",max);}catch{}return "";}
        private static Ppt.Shape NotesBody(Ppt.Slide s){try{foreach(Ppt.Shape sh in s.NotesPage.Shapes)if(sh.Type==Office.MsoShapeType.msoPlaceholder&&sh.PlaceholderFormat.Type==Ppt.PpPlaceholderType.ppPlaceholderBody)return sh;}catch{}return null;}
        private static string GetNotes(Ppt.Slide s){var sh=NotesBody(s);try{return sh!=null?sh.TextFrame.TextRange.Text??"":"";}catch{return "";}}
        private static int Int(ToolArguments a,string k,int min,int max,int fallback=0){int v;if(!int.TryParse(a.Get(k,fallback.ToString(CultureInfo.InvariantCulture)),NumberStyles.Integer,CultureInfo.InvariantCulture,out v)||v<min||v>max)throw new ArgumentException(k+" must be "+min+".."+max+".");return v;}
        private static double Double(ToolArguments a,string k,double min,double max,double fallback){double v;if(!double.TryParse(a.Get(k,fallback.ToString(CultureInfo.InvariantCulture)),NumberStyles.Float,CultureInfo.InvariantCulture,out v)||double.IsNaN(v)||double.IsInfinity(v)||v<min||v>max)throw new ArgumentException(k+" must be "+min+".."+max+".");return v;}
        private static float F(ToolArguments a,string k,double min,double max,double fallback){return (float)Double(a,k,min,max,fallback);}
        private static bool Bool(ToolArguments a,string k,bool fallback){var t=a.Token(k);if(t==null)return fallback;bool v;if(!bool.TryParse(t.ToString(),out v))throw new ArgumentException(k+" must be true or false.");return v;}
        private static int Ole(string html){if(string.IsNullOrWhiteSpace(html)||html.Length!=7||html[0]!='#'||!html.Substring(1).All(Uri.IsHexDigit))throw new ArgumentException("Color must be #RRGGBB.");return System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml(html));}
        private static string Q(string s){return Newtonsoft.Json.JsonConvert.SerializeObject(s??"");}
        private static string Op(ToolArguments a){string op=(a.Get("operation","")??"").Trim();if(op.Length==0)throw new ArgumentException("operation is required.");return op;}
        private static void RequireHttpsOrMailto(string url){Uri u;if(!Uri.TryCreate(url,UriKind.Absolute,out u)||(!u.Scheme.Equals("https",StringComparison.OrdinalIgnoreCase)&&!u.Scheme.Equals("mailto",StringComparison.OrdinalIgnoreCase)))throw new ArgumentException("Hyperlinks must use https or mailto.");}
    }
}
