$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $root 'src\OMNIX.Core\bin\Release'
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$wpf = Join-Path $framework 'WPF'
$source = @'
using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Controls.Primitives;
using OMNIX.Core.Settings;
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
            var provider=(ComboBox)settings.FindName("ProviderCombo");
            provider.ItemsSource=new ProviderRegistry().All;
            provider.DisplayMemberPath="Info.DisplayName";
            provider.SelectedIndex=0;
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
@{TestId='WORKSPACE-STARTUP-WPF-001';OverallPass=$true;Cycles=3;ColdBackgroundLocalization=$true;CredentialConstructionBounded=$true;DarkAndLightDropdownContrastPass=$true;EditableModelBindingPass=$true;RealOfficeTested=$false} | ConvertTo-Json | Set-Content (Join-Path $root 'build\artifact\workspace-startup-acceptance.json')

Remove-Item -LiteralPath $file,$exe,$stdout,$stderr -Force
