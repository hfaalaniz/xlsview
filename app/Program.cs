// ============================================================================
//  XlsViewApp — Browser propietario (WebView2) para el Visor/Editor de hojas.
//
//  Todo-en-uno: servidor HTTP local + instancia única + asociación de formatos
//  (.xlsx .xlsm .xls .csv) + ventana propia WebView2 (maximizada sin bordes,
//  barra de tareas visible).
//
//  Uso:
//     XlsViewApp.exe "hoja.xlsx"   -> abre ese archivo en el editor
//     XlsViewApp.exe                -> abre el editor vacío
//     XlsViewApp.exe --install      -> asocia formatos y añade al PATH
//     XlsViewApp.exe --uninstall    -> revierte
//     XlsViewApp.exe --stop         -> cierra la instancia en ejecución
// ============================================================================

using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace XlsView
{
    internal static class Program
    {
        private const string PipeName = "XlsViewApp.pipe";
        private const string MutexName = "Global\\XlsViewApp.singleton";

        private static Server _server;
        private static ViewerForm _form;
        private static SynchronizationContext _ui;

        [STAThread]
        private static int Main(string[] args)
        {
            try
            {
                if (args.Length > 0)
                {
                    string cmd = args[0].ToLowerInvariant();
                    if (cmd == "--install")   { Installer.Install();   return 0; }
                    if (cmd == "--uninstall") { Installer.Uninstall(); return 0; }
                    if (cmd == "--stop")      { SendToServer("STOP");  return 0; }
                    if (cmd == "--help" || cmd == "/?")
                    {
                        MessageBox.Show(
                            "XlsViewApp — Visor / Editor de hojas de cálculo\n\n" +
                            "  XlsViewApp.exe \"hoja.xlsx\"   Abrir una hoja\n" +
                            "  XlsViewApp.exe --install       Asociar formatos y añadir al PATH\n" +
                            "  XlsViewApp.exe --uninstall     Revertir\n" +
                            "  XlsViewApp.exe --stop          Cerrar la ventana\n\n" +
                            "En la ventana:  Ctrl+S = guardar   Ctrl+Shift+S = guardar como\n" +
                            "  Esc = cerrar   F11 = pantalla completa   Ctrl+M = minimizar",
                            "Visor de hojas", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return 0;
                    }
                }

                string filePath = null;
                if (args.Length > 0 && !args[0].StartsWith("--"))
                {
                    filePath = Path.GetFullPath(args[0]);
                    if (!File.Exists(filePath))
                    {
                        MessageBox.Show("No se encontró el archivo:\n" + filePath,
                            "Visor de hojas", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return 1;
                    }
                }

                bool createdNew;
                using (var mutex = new Mutex(true, MutexName, out createdNew))
                {
                    if (!createdNew)
                    {
                        string payload = filePath == null ? "OPEN\n" : "OPEN\n" + filePath;
                        if (SendToServer(payload))
                            return 0;
                    }

                    string webRoot = ResolveWebRoot();
                    if (webRoot == null)
                    {
                        MessageBox.Show(
                            "No se encontró index.html del visor.\n\n" +
                            "Coloca XlsViewApp.exe en la carpeta del visor (junto a index.html) " +
                            "o en su subcarpeta \\app\\.",
                            "Visor de hojas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return 2;
                    }

                    _server = new Server(webRoot);
                    _server.Start();

                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    _ui = new WindowsFormsSynchronizationContext();
                    SynchronizationContext.SetSynchronizationContext(_ui);

                    SplashForm splash = null;
                    if (filePath != null)
                    {
                        splash = new SplashForm(4500);
                        splash.Show();
                        splash.Refresh();
                    }

                    _form = new ViewerForm(_server.BuildViewerUrl(filePath), splash);
                    _server.UiContext = _ui;   // el servidor usa la UI para "Guardar como"

                    var pipeThread = new Thread(RunPipeServer) { IsBackground = true };
                    pipeThread.Start();

                    Application.Run(_form);
                    _server.Stop();
                }
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Visor de hojas",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 99;
            }
        }

        // Localiza la carpeta del visor (donde está index.html).
        private static string ResolveWebRoot()
        {
            var dir = new DirectoryInfo(Path.GetDirectoryName(Application.ExecutablePath));
            for (int i = 0; i < 8 && dir != null; i++)
            {
                if (File.Exists(Path.Combine(dir.FullName, "index.html")))
                    return dir.FullName;
                dir = dir.Parent;
            }
            return null;
        }

        // ------------------------------------------------------------------
        //  Instancia única (Named Pipe)
        // ------------------------------------------------------------------
        private static void RunPipeServer()
        {
            while (true)
            {
                try
                {
                    using (var server = new NamedPipeServerStream(
                        PipeName, PipeDirection.In, 1, PipeTransmissionMode.Byte, PipeOptions.None))
                    {
                        server.WaitForConnection();
                        string msg;
                        using (var reader = new StreamReader(server, Encoding.UTF8))
                            msg = reader.ReadToEnd();

                        if (msg != null && msg.StartsWith("STOP"))
                        {
                            _ui.Post(_ => { if (_form != null) _form.Close(); }, null);
                            return;
                        }
                        if (msg != null && msg.StartsWith("OPEN"))
                        {
                            string[] lines = msg.Split(new[] { '\n' }, 2);
                            string file = lines.Length > 1 ? lines[1].Trim() : null;
                            if (string.IsNullOrEmpty(file)) file = null;

                            if (file != null)
                                _server.Enqueue(file);

                            _ui.Post(_ => BringToFront(), null);
                        }
                    }
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }

        private static void BringToFront()
        {
            if (_form == null) return;
            if (_form.WindowState == FormWindowState.Minimized)
                _form.WindowState = FormWindowState.Normal;
            _form.Activate();
            _form.BringToFront();
        }

        private static bool SendToServer(string message)
        {
            try
            {
                using (var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                {
                    client.Connect(1500);
                    byte[] bytes = Encoding.UTF8.GetBytes(message);
                    client.Write(bytes, 0, bytes.Length);
                    client.Flush();
                }
                return true;
            }
            catch { return false; }
        }
    }
}
