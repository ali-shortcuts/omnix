using System;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using OMNIX.Core.AiGateway;
using OMNIX.Core.Context;
using OMNIX.Core.Logging;
using OMNIX.Core.Settings;
using OMNIX.Core.Storage;

namespace OMNIX.PowerPoint
{
    /// <summary>
    /// Thin PowerPoint host (spec Section 3, Layer 1). Static ctor logs first (spec 4.3).
    /// </summary>
    public partial class ThisAddIn
    {
        static ThisAddIn()
        {
            Logger.Startup("=== OMNIX.PowerPoint ThisAddIn: static ctor reached (pid " + System.Diagnostics.Process.GetCurrentProcess().Id + ") ===");
        }

        internal PowerPointHostAdapter Adapter { get; private set; }
        internal PowerPointTaskPaneService Panes { get; private set; }
        internal OmnixRibbon Ribbon { get; private set; }

        internal static AiGateway SharedGateway;
        internal static ChatHistoryStore SharedHistory;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Logger.Startup("ThisAddIn_Startup begin — PowerPoint version: " + SafeVersion());

            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            System.Threading.Tasks.TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

            if (Application == null)
            {
                Logger.Startup("FATAL: PowerPoint Application object is null — add-in stays passive.");
                return;
            }

            Adapter = new PowerPointHostAdapter(Application,
                () => SettingsManager.Instance.Settings.ContextMaxChars);

            if (Ribbon != null) Ribbon.BindAdapterIfReady();

            EnsureSharedServices();

            Panes = new PowerPointTaskPaneService(this, Adapter);
            Panes.AttachEvents();
            Logger.Startup("ThisAddIn_Startup complete");
        }

        internal static void EnsureSharedServices()
        {
            if (SharedGateway == null)
            {
                SharedHistory = new ChatHistoryStore();
                var registry = new ProviderRegistry();
                SharedGateway = new AiGateway(registry);
                System.Threading.Tasks.Task.Run(() => SharedGateway.ProbeLocalAsync()).ContinueWith(
                    task => Logger.Error("gateway", "Initial local discovery failed", task.Exception.GetBaseException()),
                    System.Threading.CancellationToken.None,
                    System.Threading.Tasks.TaskContinuationOptions.OnlyOnFaulted,
                    System.Threading.Tasks.TaskScheduler.Default);
            }
        }

        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Logger.Error("startup-debug", "AppDomain.UnhandledException", e.ExceptionObject as Exception);
        }

        private static void OnUnobservedTaskException(object sender, System.Threading.Tasks.UnobservedTaskExceptionEventArgs e)
        {
            Logger.Error("startup-debug", "TaskScheduler.UnobservedTaskException", e.Exception);
            e.SetObserved();
        }

        private string SafeVersion()
        {
            try { return Application.Version; } catch { return "?"; }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Logger.Startup("ThisAddIn_Shutdown begin");
            try
            {
                if (Panes != null)
                {
                    Panes.DetachEvents();
                    Panes.DisposeAll();
                    Panes = null;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("startup-debug", "PowerPoint pane shutdown cleanup failed", ex);
            }

            AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
            System.Threading.Tasks.TaskScheduler.UnobservedTaskException -= OnUnobservedTaskException;
            Logger.Startup("ThisAddIn_Shutdown complete");
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Logger.Startup("CreateRibbonExtensibilityObject -> OmnixRibbon");
            Ribbon = new OmnixRibbon(this);
            return Ribbon;
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
