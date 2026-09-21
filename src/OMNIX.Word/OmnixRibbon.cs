using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using OMNIX.Core;

namespace OMNIX.Word
{
    /// <summary>Ribbon XML extensibility for Word (spec Layer 2).</summary>
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
            if (!string.Equals(ribbonID, "Microsoft.Word.Document", StringComparison.Ordinal))
                return null;
            return LoadRibbonXml();
        }

        private static string LoadRibbonXml()
        {
            Assembly asm = typeof(OmnixRibbon).Assembly;
            using (var stream = asm.GetManifestResourceStream("OMNIX.Word.OmnixRibbon.xml"))
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

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            _ribbonUi = ribbonUi;
            var visibleHost = _addIn.Adapter as OMNIX.Core.Context.IVisibleOfficeExecutionHost;
            if (visibleHost != null) visibleHost.BindRibbon(ribbonUi);
        }
        public bool GetVisible(Office.IRibbonControl control) { return true; }
        public string GetScreentipOpen(Office.IRibbonControl control) { return "Open the OMNIX AI workspace next to this document"; }

        public void OnOpenWorkspace(Office.IRibbonControl control)
        {
            try { _addIn.Panes.ToggleActive(); }
            catch (Exception ex) { OMNIX.Core.Logging.Logger.Error("ui", "OnOpenWorkspace failed", ex); }
        }

        public void OnOpenSettings(Office.IRibbonControl control)
        {
            try { _addIn.Panes.ShowSettings(); }
            catch (Exception ex) { OMNIX.Core.Logging.Logger.Error("ui", "OnOpenSettings failed", ex); }
        }
    }
}
