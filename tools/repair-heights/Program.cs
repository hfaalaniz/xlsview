using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

// ============================================================================
//  repair-heights  —  Reparador de alturas de fila infladas de XlsView
// ----------------------------------------------------------------------------
//  Contexto (ver memoria xlsview-gotchas, bug de guardado #2):
//  Antes del fix, XlsView leia la altura en HPT (puntos), la convertia a
//  pixeles (x4/3) para el render, y al guardar escribia ese valor en pixeles
//  DIRECTAMENTE como hpt, sin dividir x3/4. Resultado: cada ciclo guardar/
//  reabrir inflaba las alturas x4/3, y como escribia Math.round(px), TODAS las
//  alturas danadas quedaron como ENTEROS redondos (p.ej. 15pt default -> 20).
//
//  Este reparador desinfla (x3/4) las alturas de las hojas cuya firma indica
//  que pasaron por ese bug, y deja intactos los archivos originales de Excel
//  (que usan alturas fraccionarias .25/.5/.75, jamas producidas por el bug).
//
//  Uso:
//    repair-heights <archivo-o-carpeta> [<mas...>] [--apply] [--recurse]
//                                       [--force] [--verbose]
//
//    (por defecto es DRY-RUN: solo informa, no escribe nada)
//    --apply     escribe los cambios (crea copia .bak antes de tocar)
//    --recurse   busca .xlsx/.xlsm dentro de las carpetas dadas
//    --force     desinfla TODA hoja con alturas enteras aunque no aparezca el
//                valor firma 20 (usar con cuidado)
//    --verbose   detalla fila por fila
// ============================================================================

internal static class Program
{
    // Factor de inflacion del bug: 1 punto se guardaba como 4/3 (px por pt).
    private const double DeflateFactor = 3.0 / 4.0;

    // Valor firma: altura default de Excel (15pt) inflada -> round(15*4/3)=20.
    private const int SignatureValue = 20;

    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    private static int Main(string[] args)
    {
        var paths = new List<string>();
        bool apply = false, recurse = false, force = false, verbose = false;

        foreach (var a in args)
        {
            switch (a.ToLowerInvariant())
            {
                case "--apply": apply = true; break;
                case "--recurse": case "-r": recurse = true; break;
                case "--force": force = true; break;
                case "--verbose": case "-v": verbose = true; break;
                case "--help": case "-h": case "/?": PrintHelp(); return 0;
                default:
                    if (a.StartsWith("--"))
                    {
                        Console.Error.WriteLine("Opcion desconocida: " + a);
                        return 2;
                    }
                    paths.Add(a);
                    break;
            }
        }

        if (paths.Count == 0) { PrintHelp(); return 2; }

        var files = new List<string>();
        foreach (var p in paths)
        {
            if (Directory.Exists(p))
            {
                var opt = recurse ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                foreach (var f in Directory.EnumerateFiles(p, "*.*", opt))
                {
                    var ext = Path.GetExtension(f).ToLowerInvariant();
                    if ((ext == ".xlsx" || ext == ".xlsm") && !f.EndsWith(".bak", StringComparison.OrdinalIgnoreCase))
                        files.Add(f);
                }
            }
            else if (File.Exists(p))
            {
                files.Add(p);
            }
            else
            {
                Console.Error.WriteLine("No existe: " + p);
            }
        }

        if (files.Count == 0) { Console.Error.WriteLine("Sin archivos que procesar."); return 1; }

        Console.WriteLine(apply ? "== MODO APLICAR (se crearan copias .bak) ==" : "== DRY-RUN (no se escribe nada; use --apply para reparar) ==");
        Console.WriteLine();

        int repaired = 0, skippedClean = 0, errors = 0;
        foreach (var file in files.Distinct())
        {
            try
            {
                var res = ProcessFile(file, apply, force, verbose);
                if (res == Outcome.Repaired) repaired++;
                else if (res == Outcome.Clean) skippedClean++;
            }
            catch (Exception ex)
            {
                errors++;
                Console.WriteLine($"[ERROR] {file}: {ex.Message}");
            }
        }

        Console.WriteLine();
        Console.WriteLine($"Resumen: {repaired} con cambios, {skippedClean} sin tocar (sanos), {errors} con error.");
        if (!apply && repaired > 0)
            Console.WriteLine("Nada se escribio. Vuelva a ejecutar con --apply para aplicar los cambios.");
        return errors > 0 ? 1 : 0;
    }

    private enum Outcome { Repaired, Clean, Error }

    private static Outcome ProcessFile(string file, bool apply, bool force, bool verbose)
    {
        // Leemos todas las entradas del ZIP en memoria.
        var entries = new List<(string Name, byte[] Data)>();
        using (var zip = ZipFile.OpenRead(file))
        {
            foreach (var e in zip.Entries)
            {
                if (string.IsNullOrEmpty(e.Name) && e.FullName.EndsWith("/")) continue; // carpeta
                using var s = e.Open();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                entries.Add((e.FullName, ms.ToArray()));
            }
        }

        var sheetEntries = entries
            .Where(en => Regex.IsMatch(en.Name, @"^xl/worksheets/sheet\d+\.xml$", RegexOptions.IgnoreCase))
            .ToList();

        int totalChanged = 0;
        var report = new StringBuilder();
        var newEntries = new List<(string Name, byte[] Data)>();

        foreach (var en in entries)
        {
            bool isSheet = sheetEntries.Any(se => se.Name == en.Name);
            if (!isSheet) { newEntries.Add(en); continue; }

            string xml = DecodeUtf8(en.Data, out var hadBom);
            var (fixedXml, changed, decision) = RepairSheet(xml, force);
            if (verbose && decision.Length > 0)
                report.Append("    " + en.Name + ": " + decision + "\n");

            if (changed > 0)
            {
                totalChanged += changed;
                newEntries.Add((en.Name, EncodeUtf8(fixedXml, hadBom)));
            }
            else
            {
                newEntries.Add(en);
            }
        }

        if (totalChanged == 0)
        {
            Console.WriteLine($"[OK sano]  {file}");
            if (verbose && report.Length > 0) Console.Write(report.ToString());
            return Outcome.Clean;
        }

        Console.WriteLine($"[REPARAR]  {file}  ({totalChanged} filas a desinflar)");
        if (verbose && report.Length > 0) Console.Write(report.ToString());

        if (apply)
        {
            var bak = file + ".bak";
            if (!File.Exists(bak)) File.Copy(file, bak);
            WriteZip(file, newEntries);
            Console.WriteLine($"           -> escrito. Backup: {bak}");
        }

        return Outcome.Repaired;
    }

    // Repara UNA hoja. Devuelve (xml corregido, filas cambiadas, mensaje).
    private static (string Xml, int Changed, string Decision) RepairSheet(string xml, bool force)
    {
        // Extraer todas las etiquetas <row ...> (self-closing o con contenido).
        var rowTagRe = new Regex(@"<row\b[^>]*>", RegexOptions.IgnoreCase);
        var htRe = new Regex(@"\bht\s*=\s*""([^""]*)""", RegexOptions.IgnoreCase);
        var customRe = new Regex(@"\bcustomHeight\s*=\s*""1""", RegexOptions.IgnoreCase);

        var customHeights = new List<double>();
        foreach (Match m in rowTagRe.Matches(xml))
        {
            if (!customRe.IsMatch(m.Value)) continue;
            var htm = htRe.Match(m.Value);
            if (!htm.Success) continue;
            if (double.TryParse(htm.Groups[1].Value, NumberStyles.Float, Inv, out var h))
                customHeights.Add(h);
        }

        if (customHeights.Count == 0)
            return (xml, 0, "sin alturas custom");

        // Firma del dano: TODAS las alturas custom son enteras. El bug hacia
        // Math.round(px), asi que jamas produjo fraccionarios. Una sola altura
        // fraccionaria (.25/.5/.75) delata que el archivo es original de Excel.
        bool allInteger = customHeights.All(IsInteger);
        bool hasSignature = customHeights.Any(h => (int)Math.Round(h) == SignatureValue);

        bool damaged = allInteger && (hasSignature || force);
        if (!damaged)
        {
            string why = !allInteger ? "alturas fraccionarias (original de Excel)"
                       : "sin valor firma 20 (use --force si esta seguro)";
            return (xml, 0, "no se toca: " + why);
        }

        // Desinflar: reescribir el ht de cada <row customHeight="1">.
        int changed = 0;
        string outXml = rowTagRe.Replace(xml, m =>
        {
            if (!customRe.IsMatch(m.Value)) return m.Value;
            return htRe.Replace(m.Value, htm =>
            {
                if (!double.TryParse(htm.Groups[1].Value, NumberStyles.Float, Inv, out var h))
                    return htm.Value;
                double deflated = Math.Round(h * DeflateFactor, 2);
                changed++;
                return "ht=\"" + deflated.ToString(Inv) + "\"";
            });
        });

        return (outXml, changed, $"desinflado x3/4, {changed} filas");
    }

    private static bool IsInteger(double v) => Math.Abs(v - Math.Round(v)) < 0.001;

    // ---- ZIP / encoding helpers ------------------------------------------

    private static void WriteZip(string path, List<(string Name, byte[] Data)> entries)
    {
        var tmp = path + ".tmp";
        using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.Write))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            foreach (var (name, data) in entries)
            {
                var e = zip.CreateEntry(name, CompressionLevel.Optimal);
                using var s = e.Open();
                s.Write(data, 0, data.Length);
            }
        }
        File.Delete(path);
        File.Move(tmp, path);
    }

    private static string DecodeUtf8(byte[] data, out bool hadBom)
    {
        hadBom = data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF;
        int off = hadBom ? 3 : 0;
        return Encoding.UTF8.GetString(data, off, data.Length - off);
    }

    private static byte[] EncodeUtf8(string s, bool withBom)
    {
        var body = Encoding.UTF8.GetBytes(s);
        if (!withBom) return body;
        var bom = new byte[] { 0xEF, 0xBB, 0xBF };
        return bom.Concat(body).ToArray();
    }

    private static void PrintHelp()
    {
        Console.WriteLine(@"repair-heights - Reparador de alturas de fila infladas de XlsView

  Uso:
    repair-heights <archivo|carpeta> [mas...] [opciones]

  Por defecto hace DRY-RUN: informa que reparara sin escribir nada.

  Opciones:
    --apply      aplica los cambios (crea copia .bak antes de tocar)
    --recurse    -r   busca .xlsx/.xlsm dentro de las carpetas
    --force      desinfla toda hoja con alturas enteras (aun sin la firma 20)
    --verbose    -v   detalla la decision por hoja
    --help       esta ayuda

  Detecta hojas cuyas alturas fueron infladas x4/3 por el bug de guardado
  previo (todas enteras + presencia del valor 20 = 15pt default inflado) y
  las desinfla x3/4. Los archivos originales de Excel (alturas .25/.5/.75) se
  dejan intactos.");
    }
}
