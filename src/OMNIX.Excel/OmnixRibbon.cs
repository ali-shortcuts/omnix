using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using OMNIX.Core;

namespace OMNIX.Excel
{
    /// <summary>
    /// Ribbon XML extensibility class (spec Layer 2). Hosts register it through
    /// CreateRibbonExtensibilityObject(); all callbacks delegate to the task pane service.
    /// </summary>
    [ComVisible(true)]
    public class OmnixRibbon : Office.IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;
        private Office.IRibbonUI _ribbonUi;

        public OmnixRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonID)
        {
            OMNIX.Core.Logging.Logger.Startup("GetCustomUI ribbonID=" + ribbonID);
            if (!string.Equals(ribbonID, "Microsoft.Excel.Workbook", StringComparison.Ordinal))
                return null;
            return LoadRibbonXml();
        }

        private static string LoadRibbonXml()
        {
            Assembly asm = typeof(OmnixRibbon).Assembly;
            using (var stream = asm.GetManifestResourceStream("OMNIX.Excel.OmnixRibbon.xml"))
            {
                if (stream == null)
                {
                    OMNIX.Core.Logging.Logger.Startup("OmnixRibbon.xml embedded resource NOT FOUND — manifest resources: " + string.Join(", ", asm.GetManifestResourceNames()));
                    return null;
                }
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        // ------------------------------------------------------------ callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            _ribbonUi = ribbonUi;
            BindAdapterIfReady();
        }

        internal void BindAdapterIfReady()
        {
            if (_ribbonUi == null || _addIn.Adapter == null) return;
            var visibleHost = _addIn.Adapter as OMNIX.Core.Context.IVisibleOfficeExecutionHost;
            if (visibleHost != null) visibleHost.BindRibbon(_ribbonUi);
        }

        public bool GetVisible(Office.IRibbonControl control)
        {
            return true;
        }

        public string GetScreentipOpen(Office.IRibbonControl control)
        {
            return "Open the OMNIX AI workspace next to this workbook";
        }

        public void OnOpenWorkspace(Office.IRibbonControl control)
        {
            try
            {
                OMNIX.Core.Logging.Logger.Startup("Ribbon: OnOpenWorkspace");
                _addIn.Panes.ToggleActive();
            }
            catch (Exception ex)
            {
                OMNIX.Core.Logging.Logger.Error("ui", "OnOpenWorkspace failed", ex);
            }
        }

        public void OnOpenSettings(Office.IRibbonControl control)
        {
            try
            {
                OMNIX.Core.Logging.Logger.Startup("Ribbon: OnOpenSettings");
                _addIn.Panes.ShowSettings();
            }
            catch (Exception ex)
            {
                OMNIX.Core.Logging.Logger.Error("ui", "OnOpenSettings failed", ex);
            }
        }
    }
}
