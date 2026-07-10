// ============================================================================
//  SplashForm — pantalla de bienvenida / progreso para la velocidad percibida.
//
//  Fase 1 (0 → ~4.5s):  "Abriendo hoja…" con spinner.
//  Fase 2 (> ~4.5s):    si aún no se pintó (archivo grande), pasa a
//                       "Procesando hoja… XX.Xs" con los datos "Acerca de".
//  El splash se cierra cuando el editor avisa que pintó ("rendered").
// ============================================================================

using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace XlsView
{
    internal sealed class SplashForm : Form
    {
        private const string AppTitle = "XlsView";
        private const string AppVersion = "v1.0.0";
        private const string Author = "Fabian Alaniz";
        private const string TechLine = "SheetJS · Univer · WebView2 · .NET";

        private readonly Timer _spin = new Timer { Interval = 33 };
        private float _angle = 0f;
        private readonly Stopwatch _clock = Stopwatch.StartNew();
        private readonly int _renderThresholdMs;
        private Image _icon;

        public SplashForm(int renderThresholdMs)
        {
            _renderThresholdMs = Math.Max(0, renderThresholdMs);

            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            TopMost = true;
            Size = new Size(400, 260);
            BackColor = Color.FromArgb(28, 28, 28);
            DoubleBuffered = true;

            try
            {
                string ico = Path.Combine(
                    Path.GetDirectoryName(Application.ExecutablePath), "app.ico");
                if (File.Exists(ico)) _icon = new Icon(ico, 48, 48).ToBitmap();
            }
            catch { }

            _spin.Tick += (s, e) =>
            {
                _angle = (_angle + 12f) % 360f;
                Invalidate();
                KeepOnTop();
            };
            _spin.Start();
        }

        private bool Rendering { get { return _clock.ElapsedMilliseconds >= _renderThresholdMs; } }

        private const uint SWP_NOSIZE = 0x0001, SWP_NOMOVE = 0x0002,
                           SWP_NOACTIVATE = 0x0010, SWP_SHOWWINDOW = 0x0040;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr after,
            int x, int y, int cx, int cy, uint flags);

        private void KeepOnTop()
        {
            if (!IsHandleCreated || IsDisposed) return;
            SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0,
                SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        }

        public void CloseSplash()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(CloseSplash)); } catch { } return; }
            _spin.Stop();
            Close();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            using (var path = RoundedRect(new Rectangle(0, 0, Width, Height), 14))
                Region = new Region(path);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            bool rendering = Rendering;

            if (_icon != null)
                g.DrawImage(_icon, (Width - 48) / 2, 30, 48, 48);

            using (var f = new Font("Segoe UI", 13f, FontStyle.Regular))
            using (var br = new SolidBrush(Color.FromArgb(235, 235, 235)))
                DrawCentered(g, AppTitle + "  " + AppVersion, f, br, 90);

            using (var f = new Font("Segoe UI", 9.5f, FontStyle.Regular))
            using (var br = new SolidBrush(Color.FromArgb(160, 160, 160)))
            {
                string msg = rendering
                    ? string.Format("Procesando hoja… {0:0.0}s", _clock.Elapsed.TotalSeconds)
                    : "Abriendo hoja…";
                DrawCentered(g, msg, f, br, 118);
            }

            int d = 24, cx = Width / 2, cy = 152;
            var rect = new Rectangle(cx - d / 2, cy - d / 2, d, d);
            using (var pen = new Pen(Color.FromArgb(68, 68, 68), 3)) g.DrawEllipse(pen, rect);
            using (var pen = new Pen(Color.FromArgb(61, 126, 255), 3))
            {
                pen.StartCap = LineCap.Round; pen.EndCap = LineCap.Round;
                g.DrawArc(pen, rect, _angle, 90);
            }

            using (var sep = new Pen(Color.FromArgb(50, 50, 50)))
                g.DrawLine(sep, 40, 188, Width - 40, 188);
            using (var fa = new Font("Segoe UI", 8.5f, FontStyle.Regular))
            using (var brA = new SolidBrush(Color.FromArgb(150, 150, 150)))
            using (var brB = new SolidBrush(Color.FromArgb(110, 110, 110)))
            {
                DrawCentered(g, "Desarrollado por " + Author, fa, brA, 200);
                DrawCentered(g, TechLine, fa, brB, 220);
            }
        }

        private void DrawCentered(Graphics g, string text, Font f, Brush br, int y)
        {
            var size = g.MeasureString(text, f);
            g.DrawString(text, f, br, (Width - size.Width) / 2, y);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            int d = radius * 2;
            var p = new GraphicsPath();
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { _spin.Dispose(); if (_icon != null) _icon.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
