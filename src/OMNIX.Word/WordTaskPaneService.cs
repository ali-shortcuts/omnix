using System;
using System.Collections.Generic;
using Wd = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using OMNIX.Core.Context;
using OMNIX.Core.Logging;
using OMNIX.Core.Ui;

namespace OMNIX.Word
{
    /// <summary>
    /// Per-window task pane management for Word: each document window gets its own
    /// pane/chat/context/provider state. Docked right at a compact default width.
    /// </summary>
    public sealed class WordTaskPaneService
    {
        private const int DefaultWidth = 360;
        private const int MaxWidth = 640;

        private readonly ThisAddIn _addIn;
        private readonly IHostAdapter _adapter;
        private readonly Dictionary<IntPtr, CustomTaskPane> _panes = new Dictionary<IntPtr, CustomTaskPane>();
        private readonly Dictionary<IntPtr, WorkspaceController> _controllers = new Dictionary<IntPtr, WorkspaceController>();
        private bool _clamping;
        private bool _disposed;
        private readonly HashSet<IntPtr> _automaticAttempts = new HashSet<IntPtr>();
        private readonly HashSet<IntPtr> _creating = new HashSet<IntPtr>();
        private System.Windows.Forms.Timer _startupTimer;

        public WordTaskPaneService(ThisAddIn addIn, IHostAdapter adapter)
        {
            _addIn = addIn;
            _adapter = adapter;
        }

        public void AttachEvents()
        {
            if (_disposed) return;
            _addIn.Application.WindowActivate += OnWindowActivate;
            _addIn.Application.WindowSelectionChange += OnWindowSelectionChange;
            _startupTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            _startupTimer.Tick += OnStartupTick;
            _startupTimer.Start();
            Logger.Startup("WordTaskPaneService events attached");
        }

        public void DetachEvents()
        {
            if (_startupTimer != null) { _startupTimer.Stop(); _startupTimer.Dispose(); _startupTimer = null; }
            try
            {
                _addIn.Application.WindowActivate -= OnWindowActivate;
                _addIn.Application.WindowSelectionChange -= OnWindowSelectionChange;
            }
            catch { }
        }

        private IntPtr KeyOf(Wd.Window window)
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
                _startupTimer.Stop(); // Stop before construction, including its failure path.
                ShowNewWindow(key, window);
            } catch (Exception ex) { Logger.Error("ui", "Deferred startup workspace failed", ex); }
        }

        private void ShowNewWindow(IntPtr key, object window)
        {
            if (_disposed || key == IntPtr.Zero || _panes.ContainsKey(key)) return;
            if (!_automaticAttempts.Add(key)) return; // One automatic attempt per window.
            var pane = EnsurePane(key, window);
            if (pane != null) pane.Visible = true;
        }

        private void OnWindowActivate(Wd.Document doc, Wd.Window wn)
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
                Logger.Error("ui", "Word OnWindowActivate failed", ex);
            }
        }

        private void OnWindowSelectionChange(Wd.Selection sel)
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

        private CustomTaskPane EnsurePane(IntPtr key, object ownerWindow = null)
        {
            if (_disposed || key == IntPtr.Zero || !_creating.Add(key)) return null;
            try { return CreatePane(key, ownerWindow); }
            finally { _creating.Remove(key); }
        }

        private CustomTaskPane CreatePane(IntPtr key, object ownerWindow)
        {
            if (_disposed || key == IntPtr.Zero) return null;

            CustomTaskPane pane;
            if (_panes.TryGetValue(key, out pane) && pane != null) return pane;

            object window = ownerWindow ?? _addIn.Application.ActiveWindow;
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
            Logger.Startup("Word task pane created for window " + key);
            return pane;
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

            _automaticAttempts.Remove(key);
            WorkspaceController controller;
            if (_controllers.TryGetValue(key, out controller))
            {
                _controllers.Remove(key);
                try { if (controller != null) controller.OnPaneClosing(); } catch { }
                try { if (controller != null) controller.Dispose(); } catch { }
            }

            Logger.Startup("Word task pane released for window " + key + " (" + reason + ")");
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
            Logger.Startup("WordTaskPaneService disposed all panes/controllers");
        }
    }
}
