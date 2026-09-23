$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $root 'src\OMNIX.Core\bin\Release'
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$wpf = Join-Path $framework 'WPF'
$source = @'
using System;
using System.Collections.Generic;
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
            model.ItemsSource=new[]{"Custom Model"}; model.SelectedIndex=0;
            var manual=(TextBox)settings.FindName("ManualModelBox"); manual.Text="private/ExactModel";
            var effective=settings.GetType().GetProperty("EffectiveModelId",BindingFlags.NonPublic|BindingFlags.Instance);
            Check((string)effective.GetValue(settings,null)=="private/ExactModel","Manual model ID was lost");
            var busy=settings.GetType().GetMethod("SetProviderOperationBusy",BindingFlags.NonPublic|BindingFlags.Instance);
            busy.Invoke(settings,new object[]{true}); Check(!manual.IsEnabled,"Manual model can change during diagnostics");
            busy.Invoke(settings,new object[]{false}); Check(manual.IsEnabled,"Manual model stays disabled after diagnostics");
            var discovered=(List<string>)settings.GetType().GetField("_discoveredModels",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(settings);
            discovered.Clear(); discovered.AddRange(new[]{"CaseModel","casemodel"});
            settings.GetType().GetMethod("RefreshModelOptions",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(settings,new object[]{"private/ExactModel"});
            Check(model.Items.Contains("CaseModel") && model.Items.Contains("casemodel") && model.Text=="private/ExactModel","Catalog refresh changed model identity");

            var verified=(Dictionary<string,ModelVerificationResult>)settings.GetType().GetField("_modelVerification",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(settings);
            verified.Clear();
            verified["working-model"]=new ModelVerificationResult {ModelId="working-model",State=ModelVerificationState.Working,ToolCallingVerified=true};
            verified["text-model"]=new ModelVerificationResult {ModelId="text-model",State=ModelVerificationState.TextOnly};
            verified["denied-model"]=new ModelVerificationResult {ModelId="denied-model",State=ModelVerificationState.AccessDenied};
            settings.GetType().GetMethod("RenderVerificationSummary",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(settings,new object[]{3,3});
            var choices=(StackPanel)settings.FindName("VerifiedModelsPanel");
            Check(choices.Children.Count==3,"Verified model choices missing");
            foreach(StackPanel row in choices.Children) {
                var button=(Button)row.Children[1];
                if((string)button.Content=="working-model") {
                    button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    Check(model.Text=="working-model","Verified model click did not select model");
                    Check(SettingsManager.Instance.Settings.SavedModels["custom"].Contains("working-model"),"Verified choice not retained");
                }
                if((string)button.Content=="denied-model") Check(!button.IsEnabled,"Denied model is selectable");
            }

            var connectionButton=(Button)settings.FindName("TestButton");
            var modelButton=(Button)settings.FindName("TestModelButton");
            var verifyButton=(Button)settings.FindName("VerifyModelsButton");
            Check(connectionButton!=null && (string)connectionButton.Content=="Test connection","Separate connection-test control missing");
            Check(modelButton!=null && (string)modelButton.Content=="Test model","Separate model-test control missing");
            Check(verifyButton!=null && (string)verifyButton.Content=="Verify models","Catalog verification control missing");
            Check(view.Chat.FindName("ActivityBorder")==null,"Chat live-activity panel must not exist; execution belongs in Office");
            view.UpdateLayout(); Snapshot(view,"settings-"+mode);
        }
    }
    sealed class FakeHost : IHostAdapter, IIndexedHostAdapter, IOfficeAccessHost {
        public bool AllowWrites; public int Writes;
        public int AccessReads;
        public string ReadOfficeAccess() { AccessReads++; return "documentPresent=true; writeToolsExposed=true; readOnly=false; workbookStructureProtected=false"; }
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
        public WritePreview PrepareWrite(string name,string json) { if (!AllowWrites) throw new NotSupportedException(); return new WritePreview { ToolName=name, ArgumentsJson=json, Title="Test", Before="empty", After="sample" }; }
        public void ApplyWrite(string name,string json) { if (!AllowWrites) throw new NotSupportedException(); Writes++; }
    }
    sealed class AccessProvider : IProviderAdapter {
        public int Calls; public bool CancelScenario;
        public ProviderInfo Info { get; private set; }
        public AccessProvider() { Info = new ProviderInfo { Id="custom", DisplayName="Test", Kind=ProviderKind.Cloud, Vision=VisionSupport.No }; }
        public void Configure(ProviderCredentials credentials) {}
        public bool SupportsVisionNow() { return false; }
        public Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct) { return Task.FromResult<IReadOnlyList<string>>(new string[0]); }
        public Task<bool> TestConnectionAsync(CancellationToken ct) { return Task.FromResult(true); }
        public Task<ChatResponse> SendAsync(ChatRequest request,Action<string> delta,CancellationToken ct) {
            Calls++;
            string write="```omnix_tool\n{\"tool\":\"create_data_table\",\"args\":{\"sheet\":\"Test\",\"headers\":[\"ID\"],\"rows\":[[1]]}}\n```";
            string verify="```omnix_tool\n{\"tool\":\"read_document_section\",\"args\":{\"sheet\":\"Test\",\"row\":1,\"column\":1,\"rows\":2,\"columns\":1}}\n```";
            string answer = CancelScenario ? (Calls==1 ? write : "Write access is unavailable")
                : Calls==1 ? "Write access is unavailable" : Calls==2 ? write : Calls==3 ? verify : "Verified completed";
            if(delta!=null) delta(answer);
            return Task.FromResult(new ChatResponse { Text=answer });
        }
    }
    sealed class NativeWriteProvider : IProviderAdapter {
        public int Calls;
        public ProviderInfo Info { get; private set; }
        public NativeWriteProvider() { Info=new ProviderInfo{Id="custom",DisplayName="Native Fixture",Kind=ProviderKind.Cloud,Vision=VisionSupport.No}; }
        public void Configure(ProviderCredentials credentials) {}
        public bool SupportsVisionNow() { return false; }
        public Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct) { return Task.FromResult<IReadOnlyList<string>>(new string[0]); }
        public Task<bool> TestConnectionAsync(CancellationToken ct) { return Task.FromResult(true); }
        public Task<ChatResponse> SendAsync(ChatRequest request,Action<string> delta,CancellationToken ct) {
            Calls++;
            if(Calls==1) return Task.FromResult(new ChatResponse {
                Text="",
                ToolCalls=new List<ProviderToolCall> {
                    new ProviderToolCall { Id="call-1", Name="omnix.create_data_table",
                        ArgumentsJson="{\"sheet\":\"NativeTest\",\"headers\":[\"ID\"],\"rows\":[[1]]}" }
                }
            });
            if(Calls==2) return Task.FromResult(new ChatResponse {
                Text="",
                ToolCalls=new List<ProviderToolCall> {
                    new ProviderToolCall { Id="call-2", Name="read_document_section",
                        ArgumentsJson="{\"sheet\":\"NativeTest\",\"row\":1,\"column\":1,\"rows\":2,\"columns\":1}" }
                }
            });
            return Task.FromResult(new ChatResponse { Text="Verified completed" });
        }
    }

    sealed class NoReadbackProvider : IProviderAdapter {
        public int Calls;
        public ProviderInfo Info { get; private set; }
        public NoReadbackProvider() { Info=new ProviderInfo{Id="custom",DisplayName="No Readback Fixture",Kind=ProviderKind.Cloud,Vision=VisionSupport.No}; }
        public void Configure(ProviderCredentials credentials) {}
        public bool SupportsVisionNow() { return false; }
        public Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct) { return Task.FromResult<IReadOnlyList<string>>(new string[0]); }
        public Task<bool> TestConnectionAsync(CancellationToken ct) { return Task.FromResult(true); }
        public Task<ChatResponse> SendAsync(ChatRequest request,Action<string> delta,CancellationToken ct) {
            Calls++;
            if(Calls==1) return Task.FromResult(new ChatResponse {
                ToolCalls=new List<ProviderToolCall> {
                    new ProviderToolCall { Id="call-1", Name="write_to_cell", ArgumentsJson="{\"address\":\"A1\",\"value\":\"done\"}" }
                }
            });
            return Task.FromResult(new ChatResponse { Text="Everything is complete." });
        }
    }

    sealed class PlanProbe : OMNIX.Core.Agent.IPlanVerificationHost {
        public bool First = true;
        public bool Second = true;
        public string CheckPostcondition(Newtonsoft.Json.Linq.JObject check) {
            return ((string)check["address"] == "A1" ? First : Second) ? null : "Wrong native value";
        }
    }
    static void ExecutionPlanRegression() {
        var plan = new OMNIX.Core.Agent.ExecutionPlan();
        plan.Begin("write two cells", true);
        var first = new ToolCall { Name = ToolNames.WriteToCell, ArgumentsJson = "{\"sheet\":\"Sheet1\",\"address\":\"A1\",\"value\":1}" };
        Check(plan.BeforeWrite(first) != null, "Unplanned write accepted");
        string json = "{\"steps\":[{\"id\":\"one\",\"tool\":\"write_to_cell\",\"args\":{\"sheet\":\"Sheet1\",\"address\":\"A1\",\"value\":1},\"checks\":[{\"kind\":\"cell_value\",\"sheet\":\"Sheet1\",\"address\":\"A1\",\"value\":1}]},{\"id\":\"two\",\"tool\":\"write_to_cell\",\"args\":{\"sheet\":\"Sheet1\",\"address\":\"A2\",\"value\":2},\"checks\":[{\"kind\":\"cell_value\",\"sheet\":\"Sheet1\",\"address\":\"A2\",\"value\":2}]}]}";
        plan.Submit(json, HostType.Excel);
        var probe = new PlanProbe();
        Check(plan.BeforeWrite(new ToolCall { Name=ToolNames.WriteToCell, ArgumentsJson="{}" }) != null, "Mismatched plan arguments accepted");
        Check(plan.BeforeWrite(first) == null, "Exact planned write rejected");
        plan.AfterWrite(probe);
        Check(!plan.Complete, "Partial plan reported complete");
        bool rejected = false;
        try { plan.Submit(json.Replace("\"checks\":[", "\"checks\":[],\"ignored\":["), HostType.Excel); } catch(ArgumentException) { rejected=true; }
        Check(rejected, "Acceptance criteria were weakened after applying a write");
        var second = new ToolCall { Name=ToolNames.WriteToCell, ArgumentsJson="{\"sheet\":\"Sheet1\",\"address\":\"A2\",\"value\":2}" };
        Check(plan.BeforeWrite(second) == null, "Next planned write rejected");
        probe.First=false;
        plan.AfterWrite(probe);
        Check(!plan.Complete, "Later write invalidated an earlier step without detection");
        probe.First=true;
        plan.VerifyAll(probe);
        Check(plan.Complete, "Native postconditions did not complete plan");
        Check(plan.BeforeWrite(first)!=null, "Completed writes replayed");
        Check(!OMNIX.Core.Agent.OfficePostconditions.ValuesEqual("5",new Newtonsoft.Json.Linq.JValue(5)), "Numeric text accepted as a real number");
        Check(OMNIX.Core.Agent.OfficePostconditions.ValuesEqual(5.0,new Newtonsoft.Json.Linq.JValue(5)), "Equivalent numeric value rejected");
        foreach(string name in new[]{"gold","inventory","invoice"}) {
            var template=OMNIX.Core.Agent.OfficePlaybooks.Template(name,"Demo",HostType.Excel);
            var templatePlan=Newtonsoft.Json.Linq.JObject.Parse(template);
            ExcelTableBuilder.ValidatePlan(templatePlan["steps"][0]["args"].ToString());
            var candidate=new OMNIX.Core.Agent.ExecutionPlan(); candidate.Begin("demo",true); candidate.Submit(template,HostType.Excel);
        }
        foreach(var host in new[]{HostType.Excel,HostType.Word,HostType.PowerPoint})
            Check(OMNIX.Core.Agent.OfficePlaybooks.Load(host,"gold shop").Contains("BUSINESS TASK GUIDE"), "Selective embedded playbook missing");
        plan.SaveCheckpoint = text => { throw new IOException("disk unavailable"); };
        plan.VerifyAll(probe);
        Check(plan.Complete, "Checkpoint failure changed native verification result");
    }

    static void NativeGatewayRegression() {
        var settings=SettingsManager.Instance.Settings;
        var oldProvider=settings.SelectedProviderId; var oldPrivacy=settings.Privacy; bool oldLocal=settings.PreferLocalWhenAvailable;
        try {
            settings.SelectedProviderId="custom"; settings.Privacy=PrivacyMode.CloudAllowed; settings.PreferLocalWhenAvailable=false;
            var registry=new ProviderRegistry(); var provider=new NativeWriteProvider();
            var providers=(System.Collections.Generic.List<IProviderAdapter>)typeof(ProviderRegistry).GetField("_providers",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(registry);
            providers.Clear(); providers.Add(provider);
            var gateway=new OMNIX.Core.AiGateway.AiGateway(registry);
            var host=new FakeHost { AllowWrites=true };
            int confirmations=0;
            var executor=new ToolExecutor { WriteConfirmation=preview=> { confirmations++; return Task.FromResult(true); } };
            var result=gateway.ChatAsync(new ChatRequest { UserTurn=new ChatTurn { Role=ChatRole.User,Text="Create a test table" } },
                host,part=>{},executor,CancellationToken.None).GetAwaiter().GetResult();
            Check(provider.Calls==3,"Native tool response did not require read-back before final provider turn");
            Check(confirmations==1 && host.Writes==1,"Native tool call with empty text was not executed through confirmation");
            Check(host.Reads>=2,"Latest write was not read back after execution");
            Check(result.Text=="Verified completed","Native tool loop did not return final answer after read-back");
        } finally { settings.SelectedProviderId=oldProvider; settings.Privacy=oldPrivacy; settings.PreferLocalWhenAvailable=oldLocal; }
    }

    static void VerificationEnforcementRegression() {
        var settings=SettingsManager.Instance.Settings;
        var oldProvider=settings.SelectedProviderId; var oldPrivacy=settings.Privacy; bool oldLocal=settings.PreferLocalWhenAvailable;
        try {
            settings.SelectedProviderId="custom"; settings.Privacy=PrivacyMode.CloudAllowed; settings.PreferLocalWhenAvailable=false;
            var registry=new ProviderRegistry(); var provider=new NoReadbackProvider();
            var providers=(System.Collections.Generic.List<IProviderAdapter>)typeof(ProviderRegistry).GetField("_providers",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(registry);
            providers.Clear(); providers.Add(provider);
            var gateway=new OMNIX.Core.AiGateway.AiGateway(registry);
            var host=new FakeHost { AllowWrites=true }; int confirmations=0;
            var executor=new ToolExecutor { WriteConfirmation=preview=> { confirmations++; return Task.FromResult(true); } };
            var result=gateway.ChatAsync(new ChatRequest { UserTurn=new ChatTurn { Role=ChatRole.User,Text="Write done into A1 in this Excel file" } },host,part=>{},executor,CancellationToken.None).GetAwaiter().GetResult();
            Check(confirmations==1 && host.Writes==1,"Verification enforcement repeated or skipped the approved write");
            Check(provider.Calls==4,"Verification enforcement did not perform bounded repair attempts");
            Check(result.Text.Contains("could not be verified") || result.Text.Contains("runtime stopped safely"),"Unverified write was incorrectly reported as completed");
        } finally { settings.SelectedProviderId=oldProvider; settings.Privacy=oldPrivacy; settings.PreferLocalWhenAvailable=oldLocal; }
    }

    static void AccessRecoveryRegression() {
        var settings=SettingsManager.Instance.Settings;
        var oldProvider=settings.SelectedProviderId; var oldPrivacy=settings.Privacy; bool oldLocal=settings.PreferLocalWhenAvailable;
        try {
            settings.SelectedProviderId="custom"; settings.Privacy=PrivacyMode.CloudAllowed; settings.PreferLocalWhenAvailable=false;
            foreach(bool cancel in new[]{false,true}) {
                var registry=new ProviderRegistry(); var provider=new AccessProvider { CancelScenario=cancel };
                var providers=(System.Collections.Generic.List<IProviderAdapter>)typeof(ProviderRegistry).GetField("_providers",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(registry);
                providers.Clear(); providers.Add(provider);
                var gateway=new OMNIX.Core.AiGateway.AiGateway(registry);
                var host=new FakeHost { AllowWrites=true }; int confirmations=0;
                var executor=new ToolExecutor { WriteConfirmation=preview=> { confirmations++; return Task.FromResult(!cancel); } };
                var result=gateway.ChatAsync(new ChatRequest { UserTurn=new ChatTurn { Role=ChatRole.User,Text="Create a test table" } },host,part=>{},executor,CancellationToken.None).GetAwaiter().GetResult();
                Check(confirmations==1 && host.Writes==(cancel ? 0 : 1),"Recovery bypassed confirmation or failed to execute approved write");
                Check(provider.Calls==(cancel ? 2 : 4),"Recovery retried cancelled write or skipped required read-back");
                if(!cancel) {
                    Check(host.Reads>=2,"Recovery path did not read back the successful write");
                    Check(result.Text=="Verified completed","Gateway did not return corrected final answer after verification");
                }
            }
        } finally { settings.SelectedProviderId=oldProvider; settings.Privacy=oldPrivacy; settings.PreferLocalWhenAvailable=oldLocal; }
    }
    static void CapabilityRegression() {
        var host=new FakeHost(); var executor=new ToolExecutor();
        var accessCall = new ToolCall { Name=ToolNames.ReadOfficeAccess, ArgumentsJson="{}" };
        var accessResult = executor.ExecuteAsync(accessCall,host).GetAwaiter().GetResult();
        Check(accessResult.Success && accessResult.ContentForModel.Contains("confirmationHandlerAvailable=False") && host.AccessReads==1,"Access probe must disclose missing confirmation handler");
        executor.WriteConfirmation = preview => Task.FromResult(true);
        Check(executor.ExecuteAsync(accessCall,host).GetAwaiter().GetResult().ContentForModel.Contains("confirmationHandlerAvailable=True"),"Access probe did not reflect available confirmation");
        executor.RequestScopeValidator = () => false;
        try { executor.ExecuteAsync(accessCall,host).GetAwaiter().GetResult(); throw new Exception("Stale access probe allowed"); } catch(OperationCanceledException) {}
        Check(host.AccessReads==2,"Access probe crossed stale document boundary");
        executor.RequestScopeValidator = () => true;
        var claim = typeof(OMNIX.Core.AiGateway.AiGateway).GetMethod("IsUnsupportedAccessClaim",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Static);
        Check((bool)claim.Invoke(null,new object[]{"\u062f\u0633\u062a\u0631\u0633\u06cc \u0646\u0648\u0634\u062a\u0646 \u0628\u0647 \u0641\u0627\u06cc\u0644 \u0641\u0639\u0627\u0644 \u062f\u0631 \u0627\u06cc\u0646 \u0646\u0634\u0633\u062a \u062f\u0631 \u062f\u0633\u062a\u0631\u0633 \u0646\u06cc\u0633\u062a"}),"Reported Persian access denial not recognized");
        Check((bool)claim.Invoke(null,new object[]{"Write access is unavailable"}),"English access denial not recognized");
        Check(!(bool)claim.Invoke(null,new object[]{"The table was created"}),"Normal answer misclassified as denial");
        var detector=typeof(OMNIX.Core.AiGateway.AiGateway).Assembly.GetType("OMNIX.Core.AiGateway.MutationIntentDetector",true);
        var detect=detector.GetMethod("IsLikelyMutation",BindingFlags.Public|BindingFlags.Static);
        Check((bool)detect.Invoke(null,new object[]{"در فایل اکسل شیت محصولات را بساز ولی اطلاعات قبلی را حذف نکن"}),"Persian build intent was not enforced");
        Check(!(bool)detect.Invoke(null,new object[]{"فقط بررسی کن، هیچ تغییری نده"}),"Read-only Persian request misclassified as mutation");
        Check(ToolNames.Normalize("omnix.write_to_cell")==ToolNames.WriteToCell && ToolNames.Normalize("write-to-cell")==ToolNames.WriteToCell,"Tool namespace normalization failed");
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
        string typed="{\"sheet\":\"Report\",\"uniqueName\":true,\"headers\":[\"Formula\",\"Date\"],\"rows\":[[{\"formula\":\"=1+1\",\"numberFormat\":\"0\"},{\"date\":\"2026-09-21\"}]]}";
        Check(ExcelTableBuilder.ValidatePlan(typed)!=null,"Typed formula/date table rejected");
        bool badTypedRejected=false; try { ExcelTableBuilder.ValidatePlan(typed.Replace("=1+1","1+1")); } catch { badTypedRejected=true; }
        Check(badTypedRejected,"Non-formula typed cell accepted");
        foreach(string invalid in new[]{valid.Replace("Products","Bad/Name"),valid.Replace("Price","ID"),valid.Replace("12.5]","12.5,4]"),"{}",new string('x',32001)}) {
            bool rejected=false; try { ExcelTableBuilder.ValidatePlan(invalid); } catch { rejected=true; }
            Check(rejected,"Invalid table plan accepted");
        }
        string prompt=SystemPromptBuilder.Build(host,host.ReadContext());
        Check(prompt.Contains("RUNTIME IDENTITY") && prompt.Contains("ACTIVE OFFICE HOST: Excel") &&
              prompt.Contains("read_document_section") && prompt.Contains("create_data_table") &&
              prompt.Contains("format_range"),"Professional host capabilities/runtime identity missing from prompt");
        Check(ToolNames.IsWhitelisted(ToolNames.FormatRange) && ToolNames.IsWriteTool(ToolNames.FormatRange),"format_range must remain inside confirmed write boundary");
        Check(ToolNames.IsWhitelisted(ToolNames.ListOfficeCapabilities) && !ToolNames.IsWriteTool(ToolNames.ListOfficeCapabilities),"capability discovery must be read-only");
        Check(ToolNames.IsWhitelisted(ToolNames.ExecuteOfficeCapability) && ToolNames.IsWriteTool(ToolNames.ExecuteOfficeCapability),"capability execution must remain inside confirmed write boundary");
        Check(OfficeCapabilityRegistry.ForHost(HostType.Excel).Count >= 85,"Excel capability catalog is unexpectedly narrow");
        Check(OfficeCapabilityRegistry.ForHost(HostType.Word).Count >= 64,"Word capability catalog is unexpectedly narrow");
        Check(OfficeCapabilityRegistry.ForHost(HostType.PowerPoint).Count >= 51,"PowerPoint capability catalog is unexpectedly narrow");
        Check(OfficeCapabilityRegistry.Search(HostType.Excel,"chart",0).Contains("chart.create") &&
              OfficeCapabilityRegistry.Search(HostType.Excel,"chart",0).Contains("chart.source") &&
              OfficeCapabilityRegistry.Search(HostType.Excel,"conditional",0).Contains("conditional.formula"),"Excel advanced capability discovery failed");
        Check(OfficeCapabilityRegistry.Search(HostType.Word,"review",0).Contains("review.track_changes") &&
              OfficeCapabilityRegistry.Search(HostType.Word,"content",0).Contains("content_control.add") &&
              OfficeCapabilityRegistry.Search(HostType.Word,"highlight",0).Contains("selection.highlight"),"Word advanced capability discovery failed");
        Check(OfficeCapabilityRegistry.Search(HostType.PowerPoint,"animation",0).Contains("animation.fade") &&
              OfficeCapabilityRegistry.Search(HostType.PowerPoint,"table",0).Contains("table.cell_format") &&
              OfficeCapabilityRegistry.Search(HostType.PowerPoint,"margins",0).Contains("text.margins"),"PowerPoint advanced capability discovery failed");
        Check(prompt.Contains("list_office_capabilities") && prompt.Contains("execute_office_capability"),"Capability engine missing from model prompt");
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
        foreach(bool anthropic in new[]{false,true}) foreach(int responseMode in new[]{0,1,2}) {
            bool streaming=responseMode != 0;
            bool serverStreams=responseMode == 1;
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
                    string reply=serverStreams
                        ? (anthropic ? "data: {\"type\":\"content_block_delta\",\"delta\":{\"type\":\"text_delta\",\"text\":\"OK\"}}\n\ndata: {\"type\":\"message_stop\"}\n\n" : "data: {\"choices\":[{\"delta\":{\"content\":\"OK\"}}]}\n\ndata: [DONE]\n\n")
                        : (anthropic ? "{\"content\":[{\"type\":\"text\",\"text\":\"OK\"}]}" : "{\"choices\":[{\"message\":{\"content\":\"OK\"}}]}");
                    byte[] bytes=Encoding.UTF8.GetBytes(reply); context.Response.ContentType=serverStreams?"text/event-stream":"application/json";
                    context.Response.ContentLength64=bytes.Length; await context.Response.OutputStream.WriteAsync(bytes,0,bytes.Length); context.Response.Close();
                });
                var client=new CustomOpenAiCompatibleAdapter();
                SettingsManager.Instance.Settings.EndpointConfig("custom").ApiType=anthropic?"OpenAI":"Anthropic";
                client.Configure(new ProviderCredentials {BaseUrl=origin+"/v1",ApiKey="fixture-key",Model="fixture-model",ApiType=anthropic?"Anthropic":"OpenAI"});
                var text=new StringBuilder(); Action<string> delta=streaming ? new Action<string>(x=>text.Append(x)) : null;
                var answer=client.SendAsync(new ChatRequest{UserTurn=new ChatTurn{Role=ChatRole.User,Text="Reply with OK."}},delta,timeout.Token).GetAwaiter().GetResult();
                serving.GetAwaiter().GetResult();
                Check(answer.Text=="OK" && (!streaming || text.ToString()=="OK"),"Protocol response parsing failed");
            }
        }

        // Native tool transport contract: real tools/functionDeclarations must cross the provider
        // boundary and come back as structured ProviderToolCall objects, never only prompt text.
        foreach(bool anthropic in new[]{false,true}) foreach(int responseMode in new[]{0,1,2}) {
            bool streaming=responseMode != 0;
            bool serverStreams=responseMode == 1;
            var portPicker=new TcpListener(IPAddress.Loopback,0); portPicker.Start();
            int port=((IPEndPoint)portPicker.LocalEndpoint).Port; portPicker.Stop();
            using(var server=new HttpListener()) using(var timeout=new CancellationTokenSource(5000)) {
                string origin="http://127.0.0.1:"+port; server.Prefixes.Add(origin+"/"); server.Start();
                var serving=Task.Run(async ()=> {
                    var context=await server.GetContextAsync();
                    using(var reader=new StreamReader(context.Request.InputStream)) {
                        string body=await reader.ReadToEndAsync();
                        Check(body.Contains("\"tools\"") && body.Contains("omnix_tool") && body.Contains("write_to_cell"),"Native tool schema missing from provider request");
                    }

                    string reply;
                    if(!serverStreams && !anthropic)
                        reply="{\"choices\":[{\"message\":{\"content\":null,\"tool_calls\":[{\"id\":\"call1\",\"type\":\"function\",\"function\":{\"name\":\"omnix_tool\",\"arguments\":\"{\\\"tool\\\":\\\"write_to_cell\\\",\\\"args\\\":{\\\"sheet\\\":\\\"Sheet1\\\",\\\"address\\\":\\\"B2\\\",\\\"value\\\":42}}\"}}]}}]}";
                    else if(serverStreams && !anthropic)
                        reply="data: {\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"call1\",\"function\":{\"name\":\"omnix_tool\",\"arguments\":\"{\\\"tool\\\":\\\"write_to_cell\\\",\\\"args\\\":{\\\"sheet\\\":\\\"Sheet1\\\",\\\"address\\\":\\\"B2\\\",\\\"value\\\":42}}\"}}]}}]}\n\ndata: [DONE]\n\n";
                    else if(!serverStreams)
                        reply="{\"content\":[{\"type\":\"tool_use\",\"id\":\"tool1\",\"name\":\"omnix_tool\",\"input\":{\"tool\":\"write_to_cell\",\"args\":{\"sheet\":\"Sheet1\",\"address\":\"B2\",\"value\":42}}}]}";
                    else
                        reply="data: {\"type\":\"content_block_start\",\"index\":0,\"content_block\":{\"type\":\"tool_use\",\"id\":\"tool1\",\"name\":\"omnix_tool\",\"input\":{}}}\n\ndata: {\"type\":\"content_block_delta\",\"index\":0,\"delta\":{\"type\":\"input_json_delta\",\"partial_json\":\"{\\\"tool\\\":\\\"write_to_cell\\\",\\\"args\\\":{\\\"sheet\\\":\\\"Sheet1\\\",\\\"address\\\":\\\"B2\\\",\\\"value\\\":42}}\"}}\n\ndata: {\"type\":\"message_stop\"}\n\n";

                    byte[] bytes=Encoding.UTF8.GetBytes(reply);
                    context.Response.ContentType=serverStreams?"text/event-stream":"application/json";
                    context.Response.ContentLength64=bytes.Length;
                    await context.Response.OutputStream.WriteAsync(bytes,0,bytes.Length);
                    context.Response.Close();
                });

                var client=new OpenAiCompatibleClient(origin+"/v1","Fixture",null,anthropic);
                Action<string> delta=streaming ? new Action<string>(x=>{}) : null;
                var answer=client.SendAsync(
                    new ChatRequest{UseNativeTools=true,UserTurn=new ChatTurn{Role=ChatRole.User,Text="Create it."}},
                    "fixture-key","fixture-model",delta,timeout.Token).GetAwaiter().GetResult();
                serving.GetAwaiter().GetResult();
                Check(answer.HasToolCalls && answer.ToolCalls.Count==1,"Native provider tool call was not materialized");
                Check(answer.ToolCalls[0].Name=="write_to_cell" && answer.ToolCalls[0].ArgumentsJson.Contains("\"value\":42"),"Native tool call decoded incorrectly");
            }
        }
    }
    sealed class SlowProvider : IProviderAdapter {
        public int TransportThread;
        public ProviderInfo Info { get { return new ProviderInfo {Id="custom",Kind=ProviderKind.Cloud,DisplayName="Slow fixture"}; } }
        public void Configure(ProviderCredentials c) {}
        public bool SupportsVisionNow() { return false; }
        public Task<IReadOnlyList<string>> ListModelsAsync(CancellationToken ct) { return Task.FromResult<IReadOnlyList<string>>(new string[0]); }
        public Task<bool> TestConnectionAsync(CancellationToken ct) { return Task.FromResult(true); }
        public Task<ChatResponse> SendAsync(ChatRequest request,Action<string> delta,CancellationToken ct) {
            TransportThread=Thread.CurrentThread.ManagedThreadId;
            Thread.Sleep(250); // Deliberately synchronous provider setup must not freeze Office.
            for(int i=0;i<2000;i++) if(delta!=null) delta("x");
            return Task.FromResult(new ChatResponse {Text="Done"});
        }
    }
    static void ResponsiveGatewayRegression() {
        var settings=SettingsManager.Instance.Settings;
        string old=settings.SelectedProviderId; var privacy=settings.Privacy;
        settings.SelectedProviderId="custom"; settings.Privacy=PrivacyMode.CloudAllowed;
        int owner=Thread.CurrentThread.ManagedThreadId, ticks=0;
        var heartbeat=new DispatcherTimer { Interval=TimeSpan.FromMilliseconds(10) };
        heartbeat.Tick+=(sender,args)=>ticks++; heartbeat.Start();
        try {
            var registry=new ProviderRegistry(); var provider=new SlowProvider();
            var list=(List<IProviderAdapter>)typeof(ProviderRegistry).GetField("_providers",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(registry);
            list.Clear(); list.Add(provider);
            var gateway=new OMNIX.Core.AiGateway.AiGateway(registry);
            var runner=typeof(WorkspaceController).Assembly.GetType("OMNIX.Core.Ui.OfficeUi").GetMethod("RunAsync",BindingFlags.Public|BindingFlags.Static);
            Func<Task> work=async ()=> {
                await gateway.ChatAsync(new ChatRequest {UserTurn=new ChatTurn {Role=ChatRole.User,Text="Hello"}},new FakeHost(),part=>{},new ToolExecutor(),CancellationToken.None);
                Check(Thread.CurrentThread.ManagedThreadId==owner,"Gateway continuation left Office STA");
            };
            var task=(Task)runner.Invoke(null,new object[]{Dispatcher.CurrentDispatcher,work});
            var deadline=Stopwatch.StartNew();
            while(!task.IsCompleted && deadline.ElapsedMilliseconds<5000) {
                var frame=new DispatcherFrame();
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,new Action(()=>frame.Continue=false));
                Dispatcher.PushFrame(frame);
            }
            Check(task.IsCompleted,"Responsive gateway timed out"); task.GetAwaiter().GetResult();
            Check(ticks>=2 && provider.TransportThread!=owner,"Provider work blocked Office heartbeat");
        } finally {heartbeat.Stop(); settings.SelectedProviderId=old; settings.Privacy=privacy;}
    }
    static void CatalogRoutesRegression() {
        var portListener=new TcpListener(IPAddress.Loopback,0); portListener.Start();
        int port=((IPEndPoint)portListener.LocalEndpoint).Port; portListener.Stop();
        using(var server=new HttpListener()) using(var timeout=new CancellationTokenSource(5000)) {
            string origin="http://127.0.0.1:"+port; server.Prefixes.Add(origin+"/"); server.Start();
            var serving=Task.Run(async ()=> {
                var context=await server.GetContextAsync();
                Check(context.Request.Url.AbsolutePath=="/accounts/test/ai/models/search","Cloudflare discovery used wrong route");
                Check(context.Request.QueryString["format"]=="openrouter" && context.Request.QueryString["page"]=="1","Cloudflare query missing");
                Check(context.Request.Headers["Authorization"]=="Bearer fixture-key","Catalog auth missing");
                byte[] bytes=Encoding.UTF8.GetBytes("{\"data\":[{\"id\":\"@cf/test\"}]}");
                context.Response.ContentType="application/json"; context.Response.ContentLength64=bytes.Length;
                await context.Response.OutputStream.WriteAsync(bytes,0,bytes.Length); context.Response.Close();
            });
            var client=new OpenAiCompatibleClient(origin+"/accounts/test/ai/v1","Cloudflare fixture",catalogPath:"../models/search?format=openrouter&per_page=100",pagedCatalog:true);
            var models=client.ListModelsAsync("fixture-key",timeout.Token).GetAwaiter().GetResult();
            serving.GetAwaiter().GetResult(); Check(models.Count==1 && models[0]=="@cf/test","Cloudflare catalog decode failed");
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
        string nativeAnthropic=anthropic.BuildPayload(new ChatRequest{UseNativeTools=true,SystemPrompt="Office context",UserTurn=new ChatTurn{Role=ChatRole.User,Text="Create"}},false);
        Check(nativeAnthropic.Contains("\"tools\"") && nativeAnthropic.Contains("omnix_tool") && nativeAnthropic.Contains("write_to_cell"),"Anthropic native tool schema missing");
        var gemini=new GeminiAdapter();
        string geminiPayload=gemini.BuildPayload(new ChatRequest{UseNativeTools=true,UserTurn=new ChatTurn{Role=ChatRole.User,Text="Create"}},false);
        Check(geminiPayload.Contains("functionDeclarations") && geminiPayload.Contains("omnix_tool") && geminiPayload.Contains("args_json"),"Gemini function declaration missing");
        var ollama=new OllamaAdapter();
        string ollamaPayload=ollama.BuildPayload(new ChatRequest{UseNativeTools=true,UserTurn=new ChatTurn{Role=ChatRole.User,Text="Create"}});
        Check(ollamaPayload.Contains("\"tools\"") && ollamaPayload.Contains("omnix_tool"),"Ollama native tool schema missing");
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
            var namespacedFallback = OMNIX.Core.Tools.ToolCallParser.Parse("[omnix.write_to_cell(sheet='Sheet1', address='A1', value='x')]");
            Check(namespacedFallback != null && namespacedFallback.Name == "write_to_cell", "Namespaced text-fallback tool name was not normalized before whitelist");
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
            string faReference = OMNIX.Core.Reference.OfficeReference.Search("Excel", "SUM", 0, "fa");
            string enReference = OMNIX.Core.Reference.OfficeReference.Search("Excel", "SUM", 0, "en");
            Check(!string.IsNullOrWhiteSpace(faReference) && faReference != enReference && faReference.Contains("Excel") && faReference.Contains("Microsoft"), "Persian reference mode missing");
            Check(OmnixSettings.CreateDefaults().Privacy==PrivacyMode.CloudAllowed,"Fresh install cloud default incorrect");
            var titled=ExcelTableBuilder.ValidatePlan("{\"sheet\":\"Gold\",\"title\":\"Shop\",\"headers\":[\"Weight\",\"Total\"],\"rows\":[[5,{\"formula\":\"=A5*2\"}]]}");
            Check(ExcelTableBuilder.HeaderRow(titled)==4,"Separate heading did not reserve rows above table");
            try { ExcelTableBuilder.ValidatePlan("{\"sheet\":\"Gold\",\"title\":\"Shop\",\"startRow\":1,\"headers\":[\"A\"],\"rows\":[]}"); throw new Exception("Overlapping heading accepted"); } catch(ArgumentException) {}
            try { ExcelTableBuilder.ValidatePlan("{\"sheet\":\"Gold\",\"headers\":[\"A\"],\"rows\":[[\"=A2*2\"]]}"); throw new Exception("Unevaluated formula string accepted"); } catch(ArgumentException) {}
            Check(OfficeCapabilityRegistry.Exists(HostType.Excel,"sheet.heading"),"Native heading capability unavailable");
            TransportRegression();
            AsyncContextRegression();
            ResponsiveGatewayRegression();
            CatalogRoutesRegression();
            ExecutionPlanRegression();
            CapabilityRegression();
            AccessRecoveryRegression();
            NativeGatewayRegression();
            VerificationEnforcementRegression();
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
