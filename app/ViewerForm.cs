// ============================================================================
//  ViewerForm — ventana propia (WebView2) que muestra el editor de hojas.
//
//  Arranca MAXIMIZADA y SIN BORDES, ocupando el área de trabajo de Windows
//  (la barra de tareas permanece visible). Un WebView2 llena la ventana y
//  carga el editor local.
// ============================================================================

using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace XlsView
{
    internal sealed class ViewerForm : Form
    {
        private readonly WebView2 _web;
        private readonly string _initialUrl;
        private bool _trueFullscreen = false;
        private SplashForm _splash;

        public ViewerForm(string initialUrl, SplashForm splash = null)
        {
            _initialUrl = initialUrl;
            _splash = splash;

            Text = "Visor / Editor de hojas";
            try { Icon = new Icon(Path.Combine(AppDir(), "app.ico")); } catch { }
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = true;
            BackColor = Color.FromArgb(30, 30, 30);
            KeyPreview = true;

            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);

            Load += OnLoad;
            KeyDown += OnKeyDown;
            var screenTimer = new Timer { Interval = 1000 };
            screenTimer.Tick += (s, e) => { if (!_trueFullscreen) ApplyWorkingAreaBounds(); };
            screenTimer.Start();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            ApplyWorkingAreaBounds();
        }

        private static string AppDir()
        {
            return Path.GetDirectoryName(Application.ExecutablePath);
        }

        private void ApplyWorkingAreaBounds()
        {
            var screen = Screen.FromHandle(Handle) ?? Screen.PrimaryScreen;
            var wa = screen.WorkingArea;
            WindowState = FormWindowState.Normal;
            MinimumSize = System.Drawing.Size.Empty;
            Bounds = wa;
        }

        // ------------------------------------------------------------------
        //  Arrastre de la ventana (asa de la toolbar) — seguimiento manual.
        // ------------------------------------------------------------------
        private bool _dragging = false;
        private Size _dragOffset;

        private void HandleDrag(string msg)
        {
            if (_trueFullscreen) return;

            if (msg.StartsWith("dragstart", StringComparison.OrdinalIgnoreCase))
            {
                _dragging = true;
                var c = Cursor.Position;
                _dragOffset = new Size(c.X - Location.X, c.Y - Location.Y);
            }
            else if (msg.StartsWith("dragmove", StringComparison.OrdinalIgnoreCase))
            {
                if (!_dragging) return;
                var c = Cursor.Position;
                Location = new Point(c.X - _dragOffset.Width, c.Y - _dragOffset.Height);
            }
            else if (msg.StartsWith("dragend", StringComparison.OrdinalIgnoreCase))
            {
                if (!_dragging) return;
                _dragging = false;
                var screen = Screen.FromPoint(Cursor.Position) ?? Screen.PrimaryScreen;
                Bounds = screen.WorkingArea;
            }
        }

        private async void OnLoad(object sender, EventArgs e)
        {
            await InitWebView();
            if (!string.IsNullOrEmpty(_initialUrl))
                _web.CoreWebView2.Navigate(_initialUrl);
        }

        private async Task InitWebView()
        {
            string userData = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XlsView", "WebView2");
            Directory.CreateDirectory(userData);

            // Puerto de depuración remota opcional (solo si XLSVIEW_DEBUG_PORT):
            // permite inspeccionar/automatizar la ventana con Chrome DevTools.
            string dbgPort = Environment.GetEnvironmentVariable("XLSVIEW_DEBUG_PORT");
            CoreWebView2EnvironmentOptions opts = null;
            if (!string.IsNullOrEmpty(dbgPort))
                opts = new CoreWebView2EnvironmentOptions(
                    "--remote-debugging-port=" + dbgPort + " --remote-allow-origins=*");

            var env = await CoreWebView2Environment.CreateAsync(null, userData, opts);
            await _web.EnsureCoreWebView2Async(env);

            var s = _web.CoreWebView2.Settings;
            s.AreDefaultContextMenusEnabled = true;
            s.IsStatusBarEnabled = false;
            s.AreBrowserAcceleratorKeysEnabled = true;
            s.IsZoomControlEnabled = true;
            s.IsSwipeNavigationEnabled = false;

            _web.CoreWebView2.NewWindowRequested += (o, a) =>
            {
                a.Handled = true;
                _web.CoreWebView2.Navigate(a.Uri);
            };

            _web.CoreWebView2.WebMessageReceived += (o, a) =>
            {
                string msg = null;
                try { msg = a.TryGetWebMessageAsString(); } catch { }
                if (msg == null) { try { msg = a.WebMessageAsJson; } catch { } }
                if (msg == null) return;
                if (msg.StartsWith("drag", StringComparison.OrdinalIgnoreCase))
                    BeginInvoke(new Action<string>(HandleDrag), msg);
                else if (msg.IndexOf("rendered", StringComparison.OrdinalIgnoreCase) >= 0)
                    BeginInvoke(new Action(DismissSplash));
                else if (msg.IndexOf("close", StringComparison.OrdinalIgnoreCase) >= 0)
                    BeginInvoke(new Action(Close));
                else if (msg.IndexOf("fullscreen", StringComparison.OrdinalIgnoreCase) >= 0)
                    BeginInvoke(new Action(ToggleTrueFullscreen));
                else if (msg.IndexOf("minimize", StringComparison.OrdinalIgnoreCase) >= 0)
                    BeginInvoke(new Action(() => WindowState = FormWindowState.Minimized));
            };
        }

        private void DismissSplash()
        {
            if (_splash != null)
            {
                var s = _splash;
                _splash = null;
                try { s.CloseSplash(); } catch { }
            }
        }

        public void NavigateTo(string url)
        {
            if (_web != null && _web.CoreWebView2 != null && !string.IsNullOrEmpty(url))
                _web.CoreWebView2.Navigate(url);
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            // Esc: cerrar SOLO si no estamos editando (evita cerrar al cancelar
            // la edición de una celda). Como no podemos saberlo aquí, dejamos
            // que Esc lo maneje el editor; F11 y Ctrl+M sí actúan.
            if (e.Alt && e.KeyCode == Keys.F4)
            {
                Close(); e.Handled = true; return;
            }
            if (e.KeyCode == Keys.F11)
            {
                ToggleTrueFullscreen();
                e.Handled = true;
                return;
            }
            if (e.Control && e.KeyCode == Keys.M)
            {
                WindowState = FormWindowState.Minimized;
                e.Handled = true;
            }
        }

        private void ToggleTrueFullscreen()
        {
            var screen = Screen.FromHandle(Handle) ?? Screen.PrimaryScreen;
            _trueFullscreen = !_trueFullscreen;
            if (_trueFullscreen)
            {
                WindowState = FormWindowState.Normal;
                MaximizedBounds = screen.Bounds;
                Bounds = screen.Bounds;
                TopMost = true;
            }
            else
            {
                TopMost = false;
                ApplyWorkingAreaBounds();
            }
        }
    }
}
