using System;
using System.Collections.Generic;
using System.Windows.Forms;
using XL = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using OMNIX.Core.Context;
using OMNIX.Core.Logging;
using OMNIX.Core.Ui;

namespace OMNIX.Excel
{
    /// <summary>
    /// Per-window task pane management for Excel (spec Section 5):
    /// every open workbook window gets its OWN pane + its own chat/context.
    /// Docked RIGHT (msoCTPDockPositionRight), default width 360px, user-resizable.
    /// </summary>
    public sealed class ExcelTaskPaneService
    {
        private const int DefaultWidth = 360;
        private const int MaxWidth = 640;

        private readonly ThisAddIn _addIn;
        private readonly IHostAdapter _adapter;
        private readonly Dictionary<IntPtr, CustomTaskPane> _panes = new Dictionary<IntPtr, CustomTaskPane>();
        private readonly Dictionary<IntPtr, WorkspaceController> _controllers = new Dictionary<IntPtr, WorkspaceController>();
        private bool _clamping;
        private bool _disposed;
        private System.Windows.Forms.Timer _startupTimer;

        public ExcelTaskPaneService(ThisAddIn addIn, IHostAdapter adapter)
        {
            _addIn = addIn;
            _adapter = adapter;
        }

        public void AttachEvents()
        {
            if (_disposed) return;
            _addIn.Application.WindowActivate += OnWindowActivate;
            _addIn.Application.WindowDeactivate += OnWindowDeactivate;
            _addIn.Application.SheetSelectionChange += OnSheetSelectionChange;
            _addIn.Application.WorkbookBeforeClose += OnWorkbookBeforeClose;
            _startupTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            _startupTimer.Tick += OnStartupTick;
            _startupTimer.Start();
            Logger.Startup("ExcelTaskPaneService events attached");
        }

        public void DetachEvents()
        {
            if (_startupTimer != null) { _startupTimer.Stop(); _startupTimer.Dispose(); _startupTimer = null; }
            try
            {
                _addIn.Application.WindowActivate -= OnWindowActivate;
                _addIn.Application.WindowDeactivate -= OnWindowDeactivate;
                _addIn.Application.SheetSelectionChange -= OnSheetSelectionChange;
                _addIn.Application.WorkbookBeforeClose -= OnWorkbookBeforeClose;
            }
            catch { }
        }

        private IntPtr KeyOf(XL.Window window)
        {
            try { return (IntPtr)window.Hwnd; }
            catch { return IntPtr.Zero; }
        }

        private void OnStartupTick(object sender, EventArgs args)
        {
            if (_disposed) return;
            try {
                var window = _addIn.Application.ActiveWindow;
                IntPtr key = KeyOf(window);
                if (key == IntPtr.Zero) return;
                ShowNewWindow(key, window);
                _startupTimer.Stop();
            } catch (Exception ex) { Logger.Error("ui", "Deferred startup workspace failed", ex); }
        }

        private void ShowNewWindow(IntPtr key, object window)
        {
            if (_disposed || key == IntPtr.Zero || _panes.ContainsKey(key)) return;
            var pane = EnsurePane(key, window);
            if (pane != null) pane.Visible = true;
        }

        private void OnWindowActivate(XL.Workbook wb, XL.Window wn)
        {
            try
            {
                IntPtr key = KeyOf(wn);
                ShowNewWindow(key, wn);
                WorkspaceController controller;
                if (_controllers.TryGetValue(key, out controller) && controller != null)
                    controller.RefreshContextBar();
            }
            catch (Exception ex)
            {
                Logger.Error("ui", "Excel OnWindowActivate failed", ex);
            }
        }

        private void OnWindowDeactivate(XL.Workbook wb, XL.Window wn) { }

        private void OnSheetSelectionChange(object sh, XL.Range target)
        {
            try
            {
                IntPtr key = KeyOf(_addIn.Application.ActiveWindow);
                WorkspaceController controller;
                if (key != IntPtr.Zero && _controllers.TryGetValue(key, out controller))
                    controller.RefreshContextBar();
            }
            catch { }
        }

        private void OnWorkbookBeforeClose(XL.Workbook wb, ref bool cancel)
        {
            // Excel can still cancel this event, so cleanup must not happen here. VSTO disposes
            // the window-bound host control only after the window is actually torn down; the
            // host-control Disposed handler then releases the matching controller/pane state.
        }

        private CustomTaskPane EnsurePane(IntPtr key, object ownerWindow = null)
        {
            if (_disposed || key == IntPtr.Zero) return null;

            CustomTaskPane pane;
            if (_panes.TryGetValue(key, out pane) && pane != null) return pane;

            object window = ownerWindow ?? ActiveWindowObject();
            if (window == null) return null;

            var controller = new WorkspaceController(_adapter, ThisAddIn.SharedHistory);
            _controllers[key] = controller;

            var hostControl = new TaskPaneHostControl(controller.View);
            try { pane = _addIn.CustomTaskPanes.Add(hostControl, "OMNIX", window); }
            catch {
                _controllers.Remove(key);
                try { controller.Dispose(); } finally { hostControl.Dispose(); }
                throw;
            }
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            try { pane.Width = DefaultWidth; } catch { }
            pane.VisibleChanged += OnPaneVisibleChanged;
            _panes[key] = pane;

            // A VSTO CustomTaskPane is window-bound. When Office destroys the underlying
            // window it disposes our WinForms host control; use that post-close signal to
            // release the per-window controller immediately instead of retaining stale
            // chat/provider/cancellation/COM state until the entire add-in shuts down.
            CustomTaskPane capturedPane = pane;
            hostControl.SizeChanged += delegate { ClampPaneWidth(capturedPane); };
            hostControl.Disposed += delegate
            {
                if (!_disposed)
                    ReleaseWindow(key, capturedPane, "task-pane host disposed");
            };

            Logger.Startup("Excel task pane created for window " + key + " (width " + DefaultWidth + ", docked right)");
            return pane;
        }

        private object ActiveWindowObject()
        {
            try { return _addIn.Application.ActiveWindow; }
            catch { return null; }
        }

        private void ClampPaneWidth(CustomTaskPane pane)
        {
            if (_clamping || pane == null) return;
            try
            {
                if (pane.Width > MaxWidth)
                {
                    _clamping = true;
                    pane.Width = MaxWidth;
                }
            }
            catch { }
            finally { _clamping = false; }
        }

        private void OnPaneVisibleChanged(object sender, EventArgs e)
        {
            var pane = sender as CustomTaskPane;
            if (pane != null && !pane.Visible)
            {
                foreach (var kv in _panes)
                {
                    if (ReferenceEquals(kv.Value, pane))
                    {
                        WorkspaceController controller;
                        if (_controllers.TryGetValue(kv.Key, out controller))
                            controller.OnPaneClosing();
                        break;
                    }
                }
            }
        }

        private void ReleaseWindow(IntPtr key, CustomTaskPane expectedPane, string reason)
        {
            if (key == IntPtr.Zero) return;

            CustomTaskPane pane;
            if (_panes.TryGetValue(key, out pane))
            {
                // HWND values can be reused by Windows. A late dispose from an old host must
                // never tear down a newly-created OMNIX pane that happens to have the same key.
                if (expectedPane != null && !ReferenceEquals(pane, expectedPane))
                    return;

                try { pane.VisibleChanged -= OnPaneVisibleChanged; } catch { }
                _panes.Remove(key);
            }
            else if (expectedPane != null)
            {
                return;
            }

            WorkspaceController controller;
            if (_controllers.TryGetValue(key, out controller))
            {
                _controllers.Remove(key);
                try { if (controller != null) controller.OnPaneClosing(); } catch { }
                try { if (controller != null) controller.Dispose(); } catch { }
            }

            Logger.Startup("Excel task pane released for window " + key + " (" + reason + ")");
        }

        public void ToggleActive()
        {
            if (_disposed) return;
            IntPtr key = KeyOf(_addIn.Application.ActiveWindow);
            if (key == IntPtr.Zero) return;
            CustomTaskPane pane = EnsurePane(key);
            if (pane == null) return;
            // Open Workspace is idempotent; closing uses the pane close button.
            pane.Visible = true;
        }

        public void ShowSettings()
        {
            if (_disposed) return;
            IntPtr key = KeyOf(_addIn.Application.ActiveWindow);
            CustomTaskPane pane = EnsurePane(key);
            if (pane == null) return;
            pane.Visible = true;
            WorkspaceController controller;
            if (_controllers.TryGetValue(key, out controller))
                controller.View.ShowSettingsTab();
        }

        public void DisposeAll()
        {
            if (_disposed) return;
            _disposed = true;
            DetachEvents();

            foreach (var controller in _controllers.Values)
            {
                try { if (controller != null) controller.Dispose(); } catch { }
            }
            _controllers.Clear();

            foreach (var pane in _panes.Values)
            {
                if (pane == null) continue;
                try { pane.VisibleChanged -= OnPaneVisibleChanged; } catch { }
                try { _addIn.CustomTaskPanes.Remove(pane); } catch { }
            }
            _panes.Clear();
            Logger.Startup("ExcelTaskPaneService disposed all panes/controllers");
        }
    }
}
