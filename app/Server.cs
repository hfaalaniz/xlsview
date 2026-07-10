// ============================================================================
//  Server — servidor HTTP local que sirve el editor y las hojas de cálculo.
//
//  Rutas:
//    GET  /                 -> index.html (y demás estáticos del visor)
//    GET  /xls/<token>      -> devuelve el archivo real (para abrirlo)
//    POST /save/<token>     -> sobrescribe en disco el archivo del token
//    POST /saveas?name=...  -> muestra "Guardar como" nativo y escribe
//    GET  /pending          -> cola de archivos a abrir (instancias nuevas)
//
//  HttpListener en 127.0.0.1, protección contra path traversal.
// ============================================================================

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace XlsView
{
    internal sealed class Server
    {
        private HttpListener _listener;
        private readonly string _webRoot;
        private int _port;

        // token -> ruta absoluta del archivo que se puede servir/guardar.
        private readonly ConcurrentDictionary<string, string> _served =
            new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly ConcurrentQueue<string> _pending = new ConcurrentQueue<string>();

        // Contexto de UI para invocar diálogos (Guardar como) en el hilo correcto.
        public SynchronizationContext UiContext { get; set; }

        public int Port { get { return _port; } }
        public string BaseUrl { get { return "http://127.0.0.1:" + _port + "/"; } }

        public Server(string webRoot) { _webRoot = webRoot; }

        public void Start()
        {
            for (int p = 8799; p < 8799 + 50; p++)
            {
                try
                {
                    var l = new HttpListener();
                    l.Prefixes.Add("http://127.0.0.1:" + p + "/");
                    l.Start();
                    _listener = l;
                    _port = p;
                    break;
                }
                catch (HttpListenerException) { /* puerto ocupado */ }
            }
            if (_listener == null)
                throw new Exception("No se pudo abrir un puerto local para el servidor.");

            var t = new Thread(ListenLoop) { IsBackground = true };
            t.Start();
        }

        public void Stop()
        {
            try { if (_listener != null) _listener.Stop(); } catch { }
        }

        public string BuildViewerUrl(string filePath)
        {
            if (filePath == null) return BaseUrl;
            return BaseUrl + "?file=/xls/" + RegisterFile(filePath);
        }

        public void Enqueue(string filePath)
        {
            if (filePath == null) return;
            _pending.Enqueue("/xls/" + RegisterFile(filePath));
        }

        // -------------------- interno --------------------
        private string RegisterFile(string filePath)
        {
            string id = Guid.NewGuid().ToString("N").Substring(0, 12);
            string fileName = Path.GetFileName(filePath);
            _served[id + "/" + fileName] = filePath;
            return id + "/" + Uri.EscapeDataString(fileName);
        }

        private static readonly Dictionary<string, string> Mime =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { ".html", "text/html; charset=utf-8" },
            { ".mjs",  "text/javascript; charset=utf-8" },
            { ".js",   "text/javascript; charset=utf-8" },
            { ".css",  "text/css; charset=utf-8" },
            { ".svg",  "image/svg+xml" },
            { ".json", "application/json" },
            { ".woff", "font/woff" },
            { ".woff2","font/woff2" },
            { ".ttf",  "font/ttf" },
            { ".png",  "image/png" },
            { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
            { ".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12" },
            { ".xls",  "application/vnd.ms-excel" },
            { ".csv",  "text/csv" },
        };

        private void ListenLoop()
        {
            while (_listener != null && _listener.IsListening)
            {
                HttpListenerContext ctx;
                try { ctx = _listener.GetContext(); }
                catch { break; }
                ThreadPool.QueueUserWorkItem(_ => HandleRequest(ctx));
            }
        }

        private void HandleRequest(HttpListenerContext ctx)
        {
            try
            {
                string path = Uri.UnescapeDataString(ctx.Request.Url.AbsolutePath);
                string method = ctx.Request.HttpMethod;

                // ---- /pending ----
                if (path.Equals("/pending", StringComparison.OrdinalIgnoreCase))
                {
                    var sb = new StringBuilder("[");
                    string item; bool first = true;
                    while (_pending.TryDequeue(out item))
                    {
                        if (!first) sb.Append(",");
                        sb.Append("\"").Append(item.Replace("\\", "\\\\").Replace("\"", "\\\"")).Append("\"");
                        first = false;
                    }
                    sb.Append("]");
                    WriteJson(ctx, sb.ToString());
                    return;
                }

                // ---- POST /save/<token> ----
                if (method == "POST" && path.StartsWith("/save/", StringComparison.OrdinalIgnoreCase))
                {
                    string token = path.Substring("/save/".Length);
                    string real;
                    if (_served.TryGetValue(token, out real))
                    {
                        byte[] body = ReadBody(ctx.Request);
                        File.WriteAllBytes(real, body);
                        WriteJson(ctx, "{\"saved\":true}");
                    }
                    else
                    {
                        ctx.Response.StatusCode = 404; ctx.Response.Close();
                    }
                    return;
                }

                // ---- POST /saveas?name=... ----
                if (method == "POST" && path.Equals("/saveas", StringComparison.OrdinalIgnoreCase))
                {
                    string suggested = ctx.Request.QueryString["name"] ?? "hoja.xlsx";
                    byte[] body = ReadBody(ctx.Request);
                    string chosen = AskSaveAs(suggested);
                    if (chosen == null)
                    {
                        WriteJson(ctx, "{\"saved\":false}");
                        return;
                    }
                    File.WriteAllBytes(chosen, body);
                    string token = RegisterFile(chosen);
                    string name = Path.GetFileName(chosen);
                    WriteJson(ctx, "{\"saved\":true,\"token\":\"" + JsonEsc("/xls/" + token) +
                                   "\",\"name\":\"" + JsonEsc(name) + "\"}");
                    return;
                }

                // ---- POST /openfile ---- (diálogo Abrir nativo; registra token)
                if (method == "POST" && path.Equals("/openfile", StringComparison.OrdinalIgnoreCase))
                {
                    string chosen = AskOpen();
                    if (chosen == null || !File.Exists(chosen))
                    {
                        WriteJson(ctx, "{\"opened\":false}");
                        return;
                    }
                    string token = RegisterFile(chosen);
                    string name = Path.GetFileName(chosen);
                    WriteJson(ctx, "{\"opened\":true,\"token\":\"" + JsonEsc("/xls/" + token) +
                                   "\",\"name\":\"" + JsonEsc(name) + "\"}");
                    return;
                }

                // ---- GET /xls/<token> ----
                if (path.StartsWith("/xls/", StringComparison.OrdinalIgnoreCase))
                {
                    string token = path.Substring("/xls/".Length);
                    string real;
                    if (_served.TryGetValue(token, out real) && File.Exists(real))
                    {
                        string ext = Path.GetExtension(real);
                        string mime = Mime.ContainsKey(ext) ? Mime[ext] : "application/octet-stream";
                        WriteFile(ctx, real, mime);
                        return;
                    }
                    ctx.Response.StatusCode = 404; ctx.Response.Close();
                    return;
                }

                // ---- estáticos del visor ----
                if (path == "/") path = "/index.html";
                string full = Path.GetFullPath(Path.Combine(_webRoot, path.TrimStart('/').Replace('/', '\\')));
                if (!full.StartsWith(_webRoot, StringComparison.OrdinalIgnoreCase))
                {
                    ctx.Response.StatusCode = 403; ctx.Response.Close();
                    return;
                }
                if (!File.Exists(full))
                {
                    ctx.Response.StatusCode = 404; ctx.Response.Close();
                    return;
                }
                string ext2 = Path.GetExtension(full);
                string mime2 = Mime.ContainsKey(ext2) ? Mime[ext2] : "application/octet-stream";
                WriteFile(ctx, full, mime2);
            }
            catch
            {
                try { ctx.Response.StatusCode = 500; ctx.Response.Close(); } catch { }
            }
        }

        // Diálogo "Guardar como" nativo, ejecutado en el hilo de UI.
        private string AskSaveAs(string suggested)
        {
            if (UiContext == null) return null;
            string result = null;
            var done = new ManualResetEventSlim(false);
            UiContext.Post(_ =>
            {
                try
                {
                    using (var dlg = new SaveFileDialog())
                    {
                        dlg.FileName = suggested;
                        string ext = Path.GetExtension(suggested).TrimStart('.').ToLowerInvariant();
                        dlg.Filter =
                            "Libro de Excel (*.xlsx)|*.xlsx|" +
                            "Libro con macros (*.xlsm)|*.xlsm|" +
                            "Excel 97-2003 (*.xls)|*.xls|" +
                            "CSV (*.csv)|*.csv|" +
                            "Todos los archivos (*.*)|*.*";
                        dlg.FilterIndex = ext == "xlsm" ? 2 : ext == "xls" ? 3 : ext == "csv" ? 4 : 1;
                        dlg.OverwritePrompt = true;
                        if (dlg.ShowDialog() == DialogResult.OK)
                            result = dlg.FileName;
                    }
                }
                catch { }
                finally { done.Set(); }
            }, null);
            done.Wait(120000);
            return result;
        }

        // Diálogo "Abrir" nativo, ejecutado en el hilo de UI. Devuelve la ruta
        // real elegida (para registrarla y poder guardar de vuelta en disco).
        private string AskOpen()
        {
            if (UiContext == null) return null;
            string result = null;
            var done = new ManualResetEventSlim(false);
            UiContext.Post(_ =>
            {
                try
                {
                    using (var dlg = new OpenFileDialog())
                    {
                        dlg.Filter =
                            "Hojas de cálculo (*.xlsx;*.xlsm;*.xls;*.csv)|*.xlsx;*.xlsm;*.xls;*.csv|" +
                            "Libro de Excel (*.xlsx)|*.xlsx|" +
                            "Libro con macros (*.xlsm)|*.xlsm|" +
                            "Excel 97-2003 (*.xls)|*.xls|" +
                            "CSV (*.csv)|*.csv|" +
                            "Todos los archivos (*.*)|*.*";
                        dlg.CheckFileExists = true;
                        if (dlg.ShowDialog() == DialogResult.OK)
                            result = dlg.FileName;
                    }
                }
                catch { }
                finally { done.Set(); }
            }, null);
            done.Wait(120000);
            return result;
        }

        // -------------------- helpers --------------------
        private static byte[] ReadBody(HttpListenerRequest req)
        {
            using (var ms = new MemoryStream())
            {
                req.InputStream.CopyTo(ms);
                return ms.ToArray();
            }
        }

        private static void WriteJson(HttpListenerContext ctx, string json)
        {
            byte[] data = Encoding.UTF8.GetBytes(json);
            ctx.Response.ContentType = "application/json; charset=utf-8";
            ctx.Response.Headers["Cache-Control"] = "no-store";
            ctx.Response.ContentLength64 = data.Length;
            ctx.Response.OutputStream.Write(data, 0, data.Length);
            ctx.Response.OutputStream.Close();
        }

        private static void WriteFile(HttpListenerContext ctx, string file, string mime)
        {
            byte[] data = File.ReadAllBytes(file);
            ctx.Response.ContentType = mime;
            ctx.Response.Headers["Cache-Control"] = "no-store";
            ctx.Response.ContentLength64 = data.Length;
            ctx.Response.OutputStream.Write(data, 0, data.Length);
            ctx.Response.OutputStream.Close();
        }

        private static string JsonEsc(string s)
        {
            return (s ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"");
        }
    }
}
