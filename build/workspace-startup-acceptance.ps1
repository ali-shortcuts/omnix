$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $root 'src\OMNIX.Core\bin\Release'
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$wpf = Join-Path $framework 'WPF'
$source = @'
using System;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Windows.Documents;
using OMNIX.Core.AiGateway.Http;
using OMNIX.Core.AiGateway.Adapters;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Controls.Primitives;
using OMNIX.Core.Settings;
using OMNIX.Core.Context;
using OMNIX.Core.Tools;
using OMNIX.Core.Storage;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using OMNIX.Core.Ui;
using OMNIX.Core.Localization;
using OMNIX.Core.Theming;
using OMNIX.Core.AiGateway;
class WorkspaceStartupRegression {
    static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
    static double Luminance(Color c) {
        Func<byte,double> f = b => { double v=b/255.0; return v<=0.04045 ? v/12.92 : Math.Pow((v+0.055)/1.055,2.4); };
        return 0.2126*f(c.R)+0.7152*f(c.G)+0.0722*f(c.B);
    }
    static void Contrast(Brush foreground, Brush background) {
        var f=foreground as SolidColorBrush; var b=background as SolidColorBrush;
        Check(f!=null && b!=null, "Theme brush missing");
        double a=Luminance(f.Color), z=Luminance(b.Color);
        Check((Math.Max(a,z)+0.05)/(Math.Min(a,z)+0.05)>=4.5, "Settings text contrast below 4.5:1: "+f.Color+" / "+b.Color);
    }
    static void Snapshot(FrameworkElement element, string name) {
        int width=(int)Math.Ceiling(element.ActualWidth), height=(int)Math.Ceiling(element.ActualHeight);
        Check(width>0 && height>0,"Screenshot layout missing");
        var bitmap=new RenderTargetBitmap(width,height,96,96,PixelFormats.Pbgra32); bitmap.Render(element);
        var encoder=new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap));
        Directory.CreateDirectory("build/artifact");
        using(var stream=File.Create("build/artifact/"+name+".png")) encoder.Save(stream);
    }
    static void SettingsRegression(WorkspaceView view) {
        view.ShowSettingsTab();
        foreach (ThemeMode mode in new[]{ThemeMode.Dark,ThemeMode.Light}) {
            SettingsManager.Instance.Settings.Theme=mode; ThemeManager.Instance.ApplyTo(view);
            var settings=view.Settings;
            ((Expander)settings.FindName("GeneralSettingsExpander")).IsExpanded = true;
            var provider=(ComboBox)settings.FindName("ProviderCombo");
            provider.ItemsSource=new ProviderRegistry().All.Select(p=>p.Info).ToList();
            provider.DisplayMemberPath="DisplayName";
            provider.SelectedIndex=0;
            Check(provider.SelectedItem.ToString()=="Custom Provider","Provider selected label must display its name");
            var model=(ComboBox)settings.FindName("ModelCombo");
            model.ItemsSource=new[]{"model-one", "model-two-with-a-long-name"};
            foreach (string name in new[]{"ProviderCombo","ModelCombo","ThemeCombo","LanguageCombo"}) {
                var combo=(ComboBox)settings.FindName(name); combo.ApplyTemplate();
                combo.IsDropDownOpen=true; combo.UpdateLayout();
                var frame=new DispatcherFrame();
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,new Action(()=>frame.Continue=false));
                Dispatcher.PushFrame(frame);
                var border=(Border)combo.Template.FindName("DropDownBorder",combo);
                Contrast(combo.Foreground,combo.Background); Contrast(combo.Foreground,border.Background);
                var item=(ComboBoxItem)combo.ItemContainerGenerator.ContainerFromIndex(0);
                Check(item!=null,"Drop-down item missing"); item.ApplyTemplate(); item.UpdateLayout();
                Contrast(item.Foreground,item.Background);
                if(name=="ProviderCombo") Snapshot(border,"settings-dropdown-"+mode);
                combo.IsDropDownOpen=false;
            }
            model.Text="manually-entered-model"; model.ApplyTemplate();
            var editor=(TextBox)model.Template.FindName("PART_EditableTextBox",model);
            Check(editor!=null && editor.Visibility==Visibility.Visible,"Editable model input missing");
            editor.Text="edited-model-id";
            Check(model.Text=="edited-model-id","Editable model binding failed");
            Contrast(editor.Foreground,editor.Background);
            view.UpdateLayout(); Snapshot(view,"settings-"+mode);
        }
    }
    sealed class FakeHost : IHostAdapter, IIndexedHostAdapter {
        public int Reads;
        public HostType Host { get { return HostType.Excel; } }
        public string HostDisplayName { get { return "Excel"; } }
        public OfficeContext ReadContext() { return new OfficeContext { Host=HostType.Excel,DocumentName="navigation-test.xlsx" }; }
        public string ReadSelection() { return "selection"; }
        public string ReadDocument(int max) { return "document"; }
        public string ReadDocumentMap(int offset) { Reads++; return "offset="+offset; }
        public string ReadDocumentSection(ToolArguments args) { Reads++; return "cell-data"; }
        public byte[] CaptureChartAsImage(string name) { return null; }
        public byte[] CaptureSlideAsImage(int index) { return null; }
        public byte[] CaptureCurrentViewAsImage() { return null; }
        public WritePreview PrepareWrite(string name,string json) { throw new NotSupportedException(); }
        public void ApplyWrite(string name,string json) { throw new NotSupportedException(); }
    }
    static void CapabilityRegression() {
        var host=new FakeHost(); var executor=new ToolExecutor();
        var map=new ToolCall { Name=ToolNames.ReadDocumentMap,ArgumentsJson="{\"offset\":20}" };
        Check(executor.ExecuteAsync(map,host).GetAwaiter().GetResult().Success && host.Reads==1,"Map navigation failed");
        var section=new ToolCall { Name=ToolNames.ReadDocumentSection,ArgumentsJson="{}" };
        Check(executor.ExecuteAsync(section,host).GetAwaiter().GetResult().Success && host.Reads==2,"Section navigation failed");
        executor.RequestScopeValidator=()=>false;
        try { executor.ExecuteAsync(section,host).GetAwaiter().GetResult(); throw new Exception("Stale navigation was allowed"); }
        catch(OperationCanceledException) {}
        Check(host.Reads==2,"Stale request read another document");
        executor.RequestScopeValidator=()=>true;
        map.ArgumentsJson="{\"offset\":-1}";
        Check(!executor.ExecuteAsync(map,host).GetAwaiter().GetResult().Success && host.Reads==2,"Invalid offset crossed host boundary");
        string valid="{\"sheet\":\"Products\",\"headers\":[\"ID\",\"Price\"],\"rows\":[[\"001\",12.5]]}";
        Check(ExcelTableBuilder.ValidatePlan(valid)!=null,"Valid table rejected");
        foreach(string invalid in new[]{valid.Replace("Products","Bad/Name"),valid.Replace("Price","ID"),valid.Replace("12.5]","12.5,4]"),"{}",new string('x',32001)}) {
            bool rejected=false; try { ExcelTableBuilder.ValidatePlan(invalid); } catch { rejected=true; }
            Check(rejected,"Invalid table plan accepted");
        }
        string prompt=SystemPromptBuilder.Build(host,host.ReadContext());
        Check(prompt.Contains("ACTIVE OFFICE HOST: Excel") && prompt.Contains("read_document_section") && prompt.Contains("create_data_table"),"Host capabilities missing from prompt");
        Check(!prompt.Contains("rewrite_selected_text {text}"),"Foreign host write tool advertised");
        using(var controller=new WorkspaceController(host,new ChatHistoryStore())) {
            var method=typeof(WorkspaceController).GetMethod("RunOnUiThread",BindingFlags.Instance|BindingFlags.NonPublic).MakeGenericMethod(typeof(bool));
            int uiThread=Thread.CurrentThread.ManagedThreadId;
            Func<bool> action=()=>Thread.CurrentThread.ManagedThreadId==uiThread;
            var same=(Task<bool>)method.Invoke(controller,new object[]{action});
            Check(same.GetAwaiter().GetResult(),"Office pane UI callback failed without Application.Current");
            var background=Task.Run(async ()=> await (Task<bool>)method.Invoke(controller,new object[]{action}));
            var deadline=Stopwatch.StartNew();
            while(!background.IsCompleted && deadline.ElapsedMilliseconds<5000) {
                var frame=new DispatcherFrame();
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,new Action(()=>frame.Continue=false));
                Dispatcher.PushFrame(frame);
            }
            Check(background.IsCompleted && background.GetAwaiter().GetResult(),"Background callback missed pane dispatcher");
        }
    }
    static void TransportRegression() {
        // Actual HTTP transport against a loopback fixture: no provider account or secret.
        foreach(bool anthropic in new[]{false,true}) foreach(bool streaming in new[]{false,true}) {
            var portPicker=new TcpListener(IPAddress.Loopback,0); portPicker.Start();
            int port=((IPEndPoint)portPicker.LocalEndpoint).Port; portPicker.Stop();
            using(var server=new HttpListener()) using(var timeout=new CancellationTokenSource(5000)) {
                string origin="http://127.0.0.1:"+port; server.Prefixes.Add(origin+"/"); server.Start();
                var serving=Task.Run(async ()=> {
                    var context=await server.GetContextAsync();
                    Check(context.Request.Url.AbsolutePath==(anthropic?"/v1/messages":"/v1/chat/completions"),"Wrong chat route");
                    Check(anthropic ? context.Request.Headers["x-api-key"]=="fixture-key" && context.Request.Headers["anthropic-version"]=="2023-06-01" : context.Request.Headers["Authorization"]=="Bearer fixture-key","Wrong provider auth headers");
                    using(var reader=new StreamReader(context.Request.InputStream)) {
                        string body=await reader.ReadToEndAsync(); Check(body.Contains("fixture-model") && body.Contains("Reply with OK."),"Diagnostic lost model or text");
                    }
                    string reply=streaming
                        ? (anthropic ? "data: {\"type\":\"content_block_delta\",\"delta\":{\"type\":\"text_delta\",\"text\":\"OK\"}}\n\ndata: {\"type\":\"message_stop\"}\n\n" : "data: {\"choices\":[{\"delta\":{\"content\":\"OK\"}}]}\n\ndata: [DONE]\n\n")
                        : (anthropic ? "{\"content\":[{\"type\":\"text\",\"text\":\"OK\"}]}" : "{\"choices\":[{\"message\":{\"content\":\"OK\"}}]}");
                    byte[] bytes=Encoding.UTF8.GetBytes(reply); context.Response.ContentType=streaming?"text/event-stream":"application/json";
                    context.Response.ContentLength64=bytes.Length; await context.Response.OutputStream.WriteAsync(bytes,0,bytes.Length); context.Response.Close();
                });
                var client=new OpenAiCompatibleClient(origin+"/v1","Fixture",null,anthropic);
                var text=new StringBuilder(); Action<string> delta=streaming ? new Action<string>(x=>text.Append(x)) : null;
                var answer=client.SendAsync(new ChatRequest{UserTurn=new ChatTurn{Role=ChatRole.User,Text="Reply with OK."}},"fixture-key","fixture-model",delta,timeout.Token).GetAwaiter().GetResult();
                serving.GetAwaiter().GetResult();
                Check(answer.Text=="OK" && (!streaming || text.ToString()=="OK"),"Protocol response parsing failed");
            }
        }
    }
    static void AsyncContextRegression() {
        SynchronizationContext.SetSynchronizationContext(null);
        int owner=Thread.CurrentThread.ManagedThreadId;
        var bubble=new ChatBubble(new ChatTurn {Role=ChatRole.Assistant,Text="test"});
        var runner=typeof(WorkspaceController).Assembly.GetType("OMNIX.Core.Ui.OfficeUi").GetMethod("RunAsync",BindingFlags.Public|BindingFlags.Static);
        foreach(bool fail in new[]{false,true}) {
            bool caught=false,finalized=false;
            Func<Task> work=async ()=> {
                try {
                    await Task.Delay(25);
                    Check(Thread.CurrentThread.ManagedThreadId==owner,"Async continuation left Office dispatcher");
                    bubble.ReplaceText("پاسخ **خوانا**");
                    if(fail) throw new InvalidOperationException("synthetic failure");
                } catch(InvalidOperationException) { caught=true; bubble.ReplaceText("Handled"); }
                finally { Check(Thread.CurrentThread.ManagedThreadId==owner,"Finally left Office dispatcher"); finalized=true; }
            };
            var task=(Task)runner.Invoke(null,new object[]{Dispatcher.CurrentDispatcher,work});
            var timer=Stopwatch.StartNew();
            while(!task.IsCompleted && timer.ElapsedMilliseconds<5000) {
                var frame=new DispatcherFrame();
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,new Action(()=>frame.Continue=false));
                Dispatcher.PushFrame(frame);
            }
            Check(task.IsCompleted,"Dispatcher async test timed out"); task.GetAwaiter().GetResult();
            Check(finalized && caught==fail,"Async error recovery failed");
        }
        Check(CustomOpenAiCompatibleAdapter.NormalizeBaseUrl("https://example.org")=="https://example.org/v1","Root URL routing failed");
        Check(CustomOpenAiCompatibleAdapter.NormalizeBaseUrl("https://example.org/proxy/v1/messages")=="https://example.org/proxy/v1","Anthropic path normalization failed");
        var anthropic=new OpenAiCompatibleClient("https://example.org/v1","Fixture",null,true);
        anthropic.SetModel("fixture-model");
        string payload=anthropic.BuildPayload(new ChatRequest{SystemPrompt="Office context",UserTurn=new ChatTurn{Role=ChatRole.User,Text="Hello"}},false);
        Check(payload.Contains("\"max_tokens\":4096") && payload.Contains("\"system\":\"Office context\"") && !payload.Contains("\"role\":\"system\""),"Anthropic payload format invalid");
    }
    [STAThread] static int Main() {
        try {
            // Cold lookup on a background thread before there is a WPF Application or view.
            Exception backgroundError = null;
            var thread = new Thread(() => {
                try { Check(Strings.T("S.Tab.Chat") == "Chat", "Cold background localization failed"); }
                catch (Exception ex) { backgroundError = ex; }
            });
            thread.SetApartmentState(ApartmentState.MTA); thread.Start();
            Check(thread.Join(10000), "Cold localization timed out");
            if (backgroundError != null) throw backgroundError;
            Check(Application.Current == null, "Test must model Office without a WPF Application");
            var xmlCall = OMNIX.Core.Tools.ToolCallParser.Parse("Checking...<tool_call>omnix_tool\n{\"tool\":\"read_document_section\",\"args\":{\"sheet\":\"test\"}}</tool_call>");
            Check(xmlCall != null && xmlCall.Name == "read_document_section", "XML tool call not parsed");
            Check(OMNIX.Core.Tools.ToolCallParser.Parse("```omnix_tool {\"tool\":\"read_selection\"}```").Name == "read_selection", "Inline fenced call not parsed");
            var nativeCall = OMNIX.Core.Tools.ToolCallParser.Parse("<|tool_call_start|>[write_to_cell(sheet='Sheet1', address='B2', value='کد محصول')]<|tool_call_end|>");
            Check(nativeCall != null && nativeCall.Name == "write_to_cell" && nativeCall.ArgumentsJson.Contains("کد محصول"), "Provider-native tool call not parsed");
            var nativeTable = OMNIX.Core.Tools.ToolCallParser.Parse("<|tool_call_start|>[create_data_table(sheet='محصولات', uniqueName=True, headers=['کد','وزن'], rows=[['T001',3]])]<|tool_call_end|>");
            Check(nativeTable != null && nativeTable.Name == "create_data_table" && nativeTable.ArgumentsJson.Contains("\"uniqueName\":true"), "Nested native table arguments not parsed");
            Check(OMNIX.Core.Tools.ToolCallParser.Parse("<tool_call>broken").Name == "", "Incomplete call must fail closed");
            Check(OMNIX.Core.Tools.ToolCallParser.Parse("plain answer") == null, "Plain answer treated as a tool");
            Check(OMNIX.Core.Tools.ToolCallParser.Parse("<tool_call>{\"tool\":\"read_selection\"}</tool_call><tool_call>{}</tool_call>").Name == "", "Ambiguous calls accepted");
            var filterType = typeof(OMNIX.Core.Tools.ToolCallParser).Assembly.GetType("OMNIX.Core.AiGateway.ToolProtocolDeltaFilter", true);
            foreach (string protocol in new[] { "<tool_call>omnix_tool\n{}\n</tool_call>", "```omnix_tool\n{}\n```", "<|tool_call_start|>[read_selection()]<|tool_call_end|>" })
            {
                var visible = new System.Text.StringBuilder();
                var filter = Activator.CreateInstance(filterType, new object[] { new Action<string>(part => visible.Append(part)) });
                foreach (char ch in "Visible prefix " + protocol)
                    filterType.GetMethod("OnDelta").Invoke(filter, new object[] { ch.ToString() });
                filterType.GetMethod("Complete").Invoke(filter, new object[] { false, protocol });
                Check(visible.ToString() == "Visible prefix ", "Split tool protocol leaked into chat");
            }
            Check(OMNIX.Core.Reference.OfficeReference.Search("Excel", "DSUM").Contains("functions/dsum-function"), "Reference catalog missing DSUM");
            Check(OMNIX.Core.Reference.OfficeReference.Search("Excel", "SUM", 0, "fa").Contains("مرجع Excel"), "Persian reference mode missing");
            TransportRegression();
            AsyncContextRegression();
            CapabilityRegression();
            var watch = Stopwatch.StartNew();
            var router = new ProviderRouter(new ProviderRegistry());
            router.BuildCredentials("ollama"); router.BuildCredentials("lmstudio");
            Check(watch.ElapsedMilliseconds < 1000, "Credential construction blocks on discovery");
            for (int i = 0; i < 3; i++) {
                var view = new WorkspaceView(null); // Loads actual compiled BAML, including Checked events.
                view.Resources.MergedDictionaries.Add(Strings.Dictionary);
                ThemeManager.Instance.ApplyTo(view);
                using (var host = new TaskPaneHostControl(view)) {
                    host.Width = 360; host.Height = 640; host.CreateControl();
                    view.Measure(new Size(360, 640)); view.Arrange(new Rect(0, 0, 360, 640)); view.UpdateLayout();
                    var chat = (FrameworkElement)view.FindName("ChatPage");
                    var settings = (FrameworkElement)view.FindName("SettingsPage");
                    var about = (FrameworkElement)view.FindName("AboutPage");
                    Check(chat.Visibility == Visibility.Visible && settings.Visibility == Visibility.Collapsed, "Initial chat visibility wrong");
                    view.ShowSettingsTab();
                    Check(settings.Visibility == Visibility.Visible && chat.Visibility == Visibility.Collapsed, "Settings navigation failed");
                    ((RadioButton)view.FindName("TabAbout")).IsChecked = true;
                    Check(about.Visibility == Visibility.Visible && settings.Visibility == Visibility.Collapsed, "About navigation failed");
                    ((RadioButton)view.FindName("TabChat")).IsChecked = true;
                    Check(chat.Visibility == Visibility.Visible && about.Visibility == Visibility.Collapsed, "Chat navigation failed");
                    Check(view.FindResource("S.Tab.Chat") as string == "Chat", "UI localization failed after background lookup");
                    if(i==0) SettingsRegression(view);
                    var frame = new DispatcherFrame();
                    Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() => frame.Continue = false));
                    Dispatcher.PushFrame(frame);
                }
            }
            Console.WriteLine("PASS: cold background localization; nonblocking credentials; three compiled WPF/ElementHost construction, navigation and dispatcher cycles.");
            return 0;
        } catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
'@
$file = Join-Path $bin 'workspace-startup-test.cs'
$exe = Join-Path $bin 'workspace-startup-test.exe'
$source | Set-Content $file -Encoding UTF8
$refs = @('System.dll','System.Core.dll','System.Xaml.dll','System.Windows.Forms.dll','System.Drawing.dll') | ForEach-Object { Join-Path $framework $_ }
$refs += @('WindowsBase.dll','PresentationCore.dll','PresentationFramework.dll','WindowsFormsIntegration.dll') | ForEach-Object { Join-Path $wpf $_ }
$refs += Join-Path $bin 'OMNIX.Core.dll'
$refs += Join-Path $bin 'Newtonsoft.Json.dll'
$args = @('/nologo','/target:exe',('/out:' + $exe)) + @($refs | ForEach-Object { '/reference:' + $_ }) + @($file)
& (Join-Path $framework 'csc.exe') @args
if ($LASTEXITCODE -ne 0) { throw 'WPF regression harness failed to compile.' }
$stdout = Join-Path $bin 'workspace-startup-test.stdout'
$stderr = Join-Path $bin 'workspace-startup-test.stderr'
$info = New-Object Diagnostics.ProcessStartInfo
$info.FileName = $exe
$info.UseShellExecute = $false
$info.RedirectStandardOutput = $true
$info.RedirectStandardError = $true
$p = New-Object Diagnostics.Process
$p.StartInfo = $info
[void]$p.Start()
$outRead = $p.StandardOutput.ReadToEndAsync()
$errRead = $p.StandardError.ReadToEndAsync()
if (-not $p.WaitForExit(30000)) { $p.Kill(); throw 'WPF startup/dispatcher exceeded 30 seconds.' }
$p.WaitForExit()
$outRead.Result | Set-Content $stdout
$errRead.Result | Set-Content $stderr
Get-Content $stdout | Write-Host
Get-Content $stderr | Write-Host
if ($p.ExitCode -ne 0) { throw "WPF startup regression failed ($($p.ExitCode))." }
New-Item -ItemType Directory -Force (Join-Path $root 'build\artifact') | Out-Null
@{TestId='WORKSPACE-STARTUP-WPF-001';OverallPass=$true;Cycles=3;ColdBackgroundLocalization=$true;CredentialConstructionBounded=$true;DarkAndLightDropdownContrastPass=$true;EditableModelBindingPass=$true;OfficePaneDispatcherPass=$true;AsyncContinuationWithoutSynchronizationContextPass=$true;AnthropicPayloadPass=$true;OpenAIAndAnthropicHttpAndStreamingPass=$true;NavigationScopeIsolationPass=$true;TablePlanValidationPass=$true;RealOfficeTested=$false} | ConvertTo-Json | Set-Content (Join-Path $root 'build\artifact\workspace-startup-acceptance.json')

Remove-Item -LiteralPath $file,$exe,$stdout,$stderr -Force
