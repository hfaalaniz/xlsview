// ============================================================================
//  Installer — asocia .xlsx/.xlsm/.xls/.csv al visor y lo añade al PATH.
//
//  Todo a nivel de USUARIO (HKCU): NO requiere permisos de administrador.
//   - Asociación:  ProgID propio (XlsView.Document) + OpenWithProgids en cada
//     extensión; fija el ProgID por defecto solo si la extensión no tenía uno.
//   - PATH:        añade la carpeta del .exe y crea alias "xlsview.cmd".
// ============================================================================

using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace XlsView
{
    internal static class Installer
    {
        private const string ProgId = "XlsView.Document";
        private const string AppName = "XlsView";
        private static readonly string[] Exts = { ".xlsx", ".xlsm", ".xls", ".csv" };

        private static string ExePath { get { return Application.ExecutablePath; } }
        private static string ExeDir  { get { return Path.GetDirectoryName(Application.ExecutablePath); } }

        public static void Install()
        {
            try
            {
                RegisterProgId();
                foreach (var ext in Exts) AssociateExt(ext);
                AddToPath();
                CreateCmdAlias();
                NotifyShellChanged();

                MessageBox.Show(
                    "Instalación completada:\n\n" +
                    "• Las extensiones .xlsx .xlsm .xls .csv ofrecen abrir con «" + AppName + "».\n" +
                    "• La carpeta se añadió al PATH del usuario.\n" +
                    "• Puedes invocarlo desde una terminal nueva con: xlsview hoja.xlsx\n\n" +
                    "Si Windows no lo usa por defecto, clic derecho en un archivo →\n" +
                    "«Abrir con» → «Elegir otra aplicación» → " + AppName + " → «Siempre».",
                    AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo completar la instalación:\n" + ex.Message,
                    AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void Uninstall()
        {
            try
            {
                using (var classes = Registry.CurrentUser.OpenSubKey(@"Software\Classes", true))
                {
                    if (classes != null)
                    {
                        try { classes.DeleteSubKeyTree(ProgId, false); } catch { }
                        try { classes.DeleteSubKeyTree(@"Applications\XlsViewApp.exe", false); } catch { }
                        foreach (var ext in Exts)
                        {
                            using (var k = classes.OpenSubKey(ext, true))
                            {
                                if (k == null) continue;
                                using (var owp = k.OpenSubKey("OpenWithProgids", true))
                                    if (owp != null) { try { owp.DeleteValue(ProgId, false); } catch { } }
                                object def = k.GetValue("");
                                if (def != null && def.ToString() == ProgId)
                                    k.SetValue("", "");
                            }
                        }
                    }
                }

                RemoveFromPath();
                RemoveCmdAlias();
                NotifyShellChanged();

                MessageBox.Show("Se revirtió la instalación (asociaciones y PATH).",
                    AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo desinstalar del todo:\n" + ex.Message,
                    AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void RegisterProgId()
        {
            using (var classes = Registry.CurrentUser.CreateSubKey(@"Software\Classes"))
            using (var prog = classes.CreateSubKey(ProgId))
            {
                prog.SetValue("", AppName);
                prog.SetValue("FriendlyTypeName", AppName);
                using (var icon = prog.CreateSubKey("DefaultIcon"))
                    icon.SetValue("", "\"" + ExePath + "\",0");
                using (var cmd = prog.CreateSubKey(@"shell\open\command"))
                    cmd.SetValue("", "\"" + ExePath + "\" \"%1\"");
            }

            using (var appPaths = Registry.CurrentUser.CreateSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\App Paths\XlsViewApp.exe"))
            {
                appPaths.SetValue("", ExePath);
                appPaths.SetValue("Path", ExeDir);
            }

            using (var classes = Registry.CurrentUser.CreateSubKey(@"Software\Classes"))
            using (var app = classes.CreateSubKey(@"Applications\XlsViewApp.exe"))
            {
                app.SetValue("FriendlyAppName", AppName);
                using (var icon = app.CreateSubKey("DefaultIcon"))
                    icon.SetValue("", "\"" + ExePath + "\",0");
                using (var cmd = app.CreateSubKey(@"shell\open\command"))
                    cmd.SetValue("", "\"" + ExePath + "\" \"%1\"");
                using (var supported = app.CreateSubKey("SupportedTypes"))
                    foreach (var ext in Exts) supported.SetValue(ext, "");
            }
        }

        private static void AssociateExt(string ext)
        {
            using (var classes = Registry.CurrentUser.CreateSubKey(@"Software\Classes"))
            using (var k = classes.CreateSubKey(ext))
            {
                using (var owp = k.CreateSubKey("OpenWithProgids"))
                    owp.SetValue(ProgId, new byte[0], RegistryValueKind.None);

                object current = k.GetValue("");
                if (current == null || string.IsNullOrEmpty(current.ToString()))
                    k.SetValue("", ProgId);
            }
        }

        private static void AddToPath()
        {
            string dir = ExeDir;
            string path = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User) ?? "";
            foreach (var part in path.Split(';'))
                if (part.Trim().TrimEnd('\\').Equals(dir.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                    return;
            string updated = string.IsNullOrEmpty(path) ? dir : path.TrimEnd(';') + ";" + dir;
            Environment.SetEnvironmentVariable("PATH", updated, EnvironmentVariableTarget.User);
        }

        private static void RemoveFromPath()
        {
            string dir = ExeDir.TrimEnd('\\');
            string path = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User) ?? "";
            var kept = new System.Collections.Generic.List<string>();
            foreach (var part in path.Split(';'))
            {
                if (string.IsNullOrWhiteSpace(part)) continue;
                if (part.Trim().TrimEnd('\\').Equals(dir, StringComparison.OrdinalIgnoreCase)) continue;
                kept.Add(part);
            }
            Environment.SetEnvironmentVariable("PATH", string.Join(";", kept), EnvironmentVariableTarget.User);
        }

        private static void CreateCmdAlias()
        {
            string alias = Path.Combine(ExeDir, "xlsview.cmd");
            File.WriteAllText(alias,
                "@echo off\r\n" +
                "\"" + ExePath + "\" %*\r\n");
        }

        private static void RemoveCmdAlias()
        {
            string alias = Path.Combine(ExeDir, "xlsview.cmd");
            try { if (File.Exists(alias)) File.Delete(alias); } catch { }
        }

        private static void NotifyShellChanged()
        {
            SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero);
            IntPtr result;
            SendMessageTimeout((IntPtr)0xFFFF, 0x001A, IntPtr.Zero, "Environment", 0x0002, 3000, out result);
        }

        [System.Runtime.InteropServices.DllImport("shell32.dll")]
        private static extern void SHChangeNotify(int eventId, uint flags, IntPtr item1, IntPtr item2);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam,
            string lParam, uint flags, uint timeout, out IntPtr result);
    }
}
