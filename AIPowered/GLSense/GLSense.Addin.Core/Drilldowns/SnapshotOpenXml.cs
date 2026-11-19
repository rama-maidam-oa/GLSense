// SnapshotOpenXml.cs - Sheet/Book Snapshot written straight into the saved copy's .xlsx package (no Excel).
//
// Replaces the hidden-Excel step (open the copy in a second Excel, write values, save, quit): ~4x faster on large
// books, and no second Excel (which switched GLSense off - CLAUDE.md 120 - and lingered - 119). Research, the
// corruption causes and measurements: CLAUDE.md 122/123, prototype docs/perf/OxmlSnapshotPrototype.cs.txt.
//
// What it does to the copy:
//   - GetBalance cells: a cell that is one GetBalance call -> the balance (number, or text when the balance is a
//     message); a cell with more around the call(s) keeps its formula with each call replaced by its balance
//     (Excel calculates it on open); a cell without a balance -> "Click Refresh...".
//   - Other GLSense formulas (GetPeriod, GetSegmentDesc...) -> their current value (number stays number, text
//     stays text, error -> "#GETTING_DATA"), as the hidden-Excel step did.
//   - Sheet snapshot: keeps only that sheet and deletes the others (a kept other sheet would still hold live GLSense
//     formulas). Defined names that refer to a deleted sheet become #REF!.
//   - Always removes calcChain.xml (Excel rebuilds it; a stale one makes Excel repair/refuse the file).
// Ordinary formulas are never touched, so nothing is recalculated except the mixed GetBalance cells.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace GLSense.Addin.Core.Drilldowns
{
#nullable enable
    internal static class SnapshotOpenXml
    {
        private static readonly XNamespace M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PR = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace CT = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string BalanceName = "GLSense_GetBalance(";
        private const string UnresolvedText = "Click Refresh..."; // what the refreshed cell would show

        internal sealed class Result
        {
            public readonly Dictionary<string, int> Counts = new Dictionary<string, int>();
            public List<string> KeptSheets = new List<string>();
            public List<string> DeletedSheets = new List<string>();
            public void Count(string key, int n = 1) { Counts.TryGetValue(key, out int c); Counts[key] = c + n; }
            public override string ToString() => string.Join(", ", Counts.Select(k => $"{k.Key} {k.Value:N0}"));
        }

        /// <param name="balances">Per sheet, per (row, col) of a GetBalance cell: the balance of each call left to right,
        /// or null when the cell got no balance.</param>
        /// <param name="keepSheet">Sheet snapshot: the sheet to keep (plus the sheets it refers to). Null = book snapshot.</param>
        internal static Result Apply(string path, IReadOnlyDictionary<string, Dictionary<(int row, int col), List<string>?>> balances,
            string? keepSheet, CancellationToken token)
        {
            var result = new Result();
            string tempPath = path + ".oxml.tmp";
            using (var source = ZipFile.OpenRead(path))
            {
                XDocument wb = Load(source, "xl/workbook.xml");
                XDocument rels = Load(source, "xl/_rels/workbook.xml.rels");
                XDocument ct = Load(source, "[Content_Types].xml");

                var relTargets = rels.Root!.Elements(PR + "Relationship").ToDictionary(r => (string)r.Attribute("Id")!, r => PartPath("xl/", (string)r.Attribute("Target")!));
                var sheets = wb.Root!.Element(M + "sheets")!.Elements(M + "sheet")
                    .Select((s, i) => new SheetInfo(s, i, (string)s.Attribute("name")!, relTargets[(string)s.Attribute(R + "id")!]))
                    .ToList();

                // Sheet snapshot keeps only its own sheet (user, 18-Nov: a kept referenced sheet would still hold live
                // GLSense formulas, which defeats the snapshot).
                HashSet<string> kept = keepSheet == null
                    ? new HashSet<string>(sheets.Select(s => s.Name), StringComparer.OrdinalIgnoreCase)
                    : new HashSet<string>(StringComparer.OrdinalIgnoreCase) { keepSheet };
                var deleted = sheets.Where(s => !kept.Contains(s.Name)).ToList();
                result.KeptSheets = sheets.Where(s => kept.Contains(s.Name)).Select(s => s.Name).ToList();
                result.DeletedSheets = deleted.Select(s => s.Name).ToList();

                var skipParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "xl/calcChain.xml" };
                foreach (var d in deleted) foreach (string p in PartsOfSheet(source, d.Part)) skipParts.Add(p);

                // Package-level XML: workbook, its relationships, content types.
                rels.Root.Elements(PR + "Relationship").Where(r => ((string)r.Attribute("Type")!).EndsWith("/calcChain", StringComparison.Ordinal)).Remove();
                foreach (var d in deleted)
                {
                    string rid = (string)d.Element.Attribute(R + "id")!;
                    rels.Root.Elements(PR + "Relationship").Where(r => (string)r.Attribute("Id")! == rid).Remove();
                    d.Element.Remove();
                }
                ct.Root!.Elements(CT + "Override").Where(o => skipParts.Contains(((string)o.Attribute("PartName")!).TrimStart('/'))).Remove();
                if (deleted.Count > 0) FixWorkbookAfterDelete(wb, sheets, deleted, keepSheet!, result);
                result.Count("calcChain removed", source.GetEntry("xl/calcChain.xml") != null ? 1 : 0);

                var sheetByPart = sheets.Where(s => kept.Contains(s.Name)).ToDictionary(s => s.Part, StringComparer.OrdinalIgnoreCase);
                Regex? deletedRefs = DeletedSheetReferences(deleted.Select(d => d.Name).ToList());
                using (var target = ZipFile.Open(tempPath, ZipArchiveMode.Create))
                {
                    foreach (var entry in source.Entries)
                    {
                        token.ThrowIfCancellationRequested();
                        string name = entry.FullName;
                        if (skipParts.Contains(name)) continue;
                        if (name == "xl/workbook.xml") { Save(target, name, wb); continue; }
                        if (name == "xl/_rels/workbook.xml.rels") { Save(target, name, rels); continue; }
                        if (name == "[Content_Types].xml") { Save(target, name, ct); continue; }
                        if (sheetByPart.TryGetValue(name, out SheetInfo? sheet))
                        {
                            string xml;
                            using (var rd = new StreamReader(entry.Open(), Encoding.UTF8)) xml = rd.ReadToEnd();
                            balances.TryGetValue(sheet.Name, out var sheetBalances);
                            bool? selected = keepSheet == null ? (bool?)null : string.Equals(sheet.Name, keepSheet, StringComparison.OrdinalIgnoreCase);
                            xml = RewriteSheet(xml, sheetBalances, selected, deletedRefs, result);
                            using (var w = new StreamWriter(target.CreateEntry(name, CompressionLevel.Fastest).Open(), new UTF8Encoding(false))) w.Write(xml);
                            continue;
                        }
                        using (var input = entry.Open())
                        using (var output = target.CreateEntry(name, CompressionLevel.Fastest).Open())
                            input.CopyTo(output);
                    }
                }
            }
            File.Delete(path);
            File.Move(tempPath, path);
            return result;
        }

        private sealed class SheetInfo
        {
            public readonly XElement Element; public readonly int Index; public readonly string Name; public readonly string Part;
            public SheetInfo(XElement e, int i, string n, string p) { Element = e; Index = i; Name = n; Part = p; }
        }

        /// <summary>Formula text refers to the sheet: 'Name'!  or, for a name that needs no quotes, Name! not preceded by
        /// another name character (so "Sheet1!" doesn't match inside "MySheet1!").</summary>
        internal static bool RefersToSheet(string formula, string sheet)
        {
            if (formula.IndexOf("'" + sheet.Replace("'", "''") + "'!", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            string plain = sheet + "!";
            for (int i = formula.IndexOf(plain, StringComparison.OrdinalIgnoreCase); i >= 0; i = formula.IndexOf(plain, i + 1, StringComparison.OrdinalIgnoreCase))
                if (i == 0 || !IsNameChar(formula[i - 1])) return true;
            return false;
        }

        private static bool IsNameChar(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\';

        /// <summary>The sheet part plus the parts only it uses (drawings, comments, tables, printer settings...).</summary>
        private static IEnumerable<string> PartsOfSheet(ZipArchive source, string part)
        {
            yield return part;
            string relsPath = Path.GetDirectoryName(part)!.Replace('\\', '/') + "/_rels/" + Path.GetFileName(part) + ".rels";
            var relsEntry = source.GetEntry(relsPath);
            if (relsEntry == null) yield break;
            yield return relsPath;
            XDocument rels;
            using (var s = relsEntry.Open()) rels = XDocument.Load(s);
            string dir = Path.GetDirectoryName(part)!.Replace('\\', '/') + "/";
            foreach (var r in rels.Root!.Elements(PR + "Relationship"))
                if ((string?)r.Attribute("TargetMode") != "External")
                    yield return PartPath(dir, (string)r.Attribute("Target")!);
        }

        private static void FixWorkbookAfterDelete(XDocument wb, List<SheetInfo> sheets, List<SheetInfo> deleted, string keepSheet, Result result)
        {
            var deletedIdx = new HashSet<int>(deleted.Select(d => d.Index));
            var names = wb.Root!.Element(M + "definedNames");
            if (names != null)
            {
                foreach (var n in names.Elements(M + "definedName").ToList())
                {
                    var lsi = n.Attribute("localSheetId");
                    if (lsi != null)
                    {
                        int idx = (int)lsi;
                        if (deletedIdx.Contains(idx)) { n.Remove(); continue; }
                        lsi.Value = (idx - deletedIdx.Count(x => x < idx)).ToString(CultureInfo.InvariantCulture);
                    }
                    // A name pointing at a deleted sheet would make Excel repair the file; Excel's own delete gives #REF!.
                    if (deleted.Any(d => RefersToSheet(n.Value, d.Name))) { n.Value = "#REF!"; result.Count("defined names -> #REF!"); }
                }
                if (!names.HasElements) names.Remove();
            }
            int keepIndex = sheets.Where(s => !deletedIdx.Contains(s.Index)).Select(s => s.Name).ToList()
                .FindIndex(n => n.Equals(keepSheet, StringComparison.OrdinalIgnoreCase));
            foreach (var v in wb.Root.Descendants(M + "workbookView"))
            {
                v.SetAttributeValue("activeTab", Math.Max(0, keepIndex).ToString(CultureInfo.InvariantCulture));
                v.Attribute("firstSheet")?.Remove();
            }
            result.Count("sheets deleted", deleted.Count);
        }

        // ---------- worksheet rewrite (text scan; cells without a formula are copied as they are) ----------

        /// <summary>A reference to one of the deleted sheets ('Name'!A1, Name!$A$1:$B$2, Name!A:A, Name!Table...), as Excel
        /// writes it. Excel's own sheet delete turns these into #REF!; left as is, Excel would take the unknown sheet
        /// for another workbook and ask for it on the next recalculation.</summary>
        private static Regex? DeletedSheetReferences(List<string> deletedSheets)
        {
            if (deletedSheets.Count == 0) return null;
            var names = deletedSheets.Select(s => "'" + Regex.Escape(s.Replace("'", "''")) + "'!|(?<![\\w.\\\\])" + Regex.Escape(s) + "!");
            return new Regex("(?:" + string.Join("|", names) + ")[A-Za-z0-9_$:.]*", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static string RewriteSheet(string xml, Dictionary<(int row, int col), List<string>?>? balances, bool? selected, Regex? deletedRefs, Result result)
        {
            if (selected.HasValue) xml = SetTabSelected(xml, selected.Value);
            SharedMasterIsSingleCall.Clear(); // shared formula ids are per sheet
            var sb = new StringBuilder(xml.Length);
            var glsenseSharedIds = new HashSet<string>(); // shared formulas whose master is GLSense (masters come first)
            var refSharedIds = new HashSet<string>();     // shared formulas whose master referred to a deleted sheet
            int pos = 0;
            while (true)
            {
                int start = xml.IndexOf("<c ", pos, StringComparison.Ordinal);
                if (start < 0) break;
                int gt = xml.IndexOf('>', start);
                bool selfClosing = xml[gt - 1] == '/';
                int end = selfClosing ? gt + 1 : xml.IndexOf("</c>", gt, StringComparison.Ordinal) + 4;
                if (selfClosing || xml.IndexOf("<f", gt, end - gt, StringComparison.Ordinal) < 0)
                    sb.Append(xml, pos, end - pos);
                else
                {
                    sb.Append(xml, pos, start - pos);
                    sb.Append(RewriteCell(xml.Substring(start, end - start), balances, glsenseSharedIds, deletedRefs, refSharedIds, result));
                }
                pos = end;
            }
            sb.Append(xml, pos, xml.Length - pos);
            return sb.ToString();
        }

        // Only the kept sheet is selected (a deleted sheet may have been the selected one; several selected = grouped).
        private static string SetTabSelected(string xml, bool selected)
        {
            int i = xml.IndexOf("<sheetView ", StringComparison.Ordinal);
            if (i < 0) return xml;
            int gt = xml.IndexOf('>', i);
            string tag = xml.Substring(i, gt - i);
            string cleaned = Regex.Replace(tag, "\\s+tabSelected=\"[^\"]*\"", "");
            if (selected) cleaned = cleaned.Insert("<sheetView".Length, " tabSelected=\"1\"");
            return xml.Substring(0, i) + cleaned + xml.Substring(gt);
        }

        private sealed class Cell
        {
            public string Raw = "", OpenAttrs = "", FAttrs = "", FText = "";
            public string? V;
            public bool Simple = true;
            private List<KeyValuePair<string, string>>? attrs, fattrs;
            public List<KeyValuePair<string, string>> Attrs => attrs ??= ParseAttrs(OpenAttrs);
            public string? Attr(string name) => Find(Attrs, name);
            public string? FAttr(string name) => Find(fattrs ??= ParseAttrs(FAttrs), name);
        }

        private static Cell ParseCell(string raw)
        {
            var c = new Cell { Raw = raw };
            int gt = raw.IndexOf('>');
            c.OpenAttrs = raw.Substring(3, gt - 3);
            int p = gt + 1, close = raw.Length - 4; // before "</c>"
            while (p < close)
            {
                if (string.CompareOrdinal(raw, p, "<f", 0, 2) == 0)
                {
                    int fgt = raw.IndexOf('>', p);
                    bool selfClosing = raw[fgt - 1] == '/';
                    c.FAttrs = raw.Substring(p + 2, (selfClosing ? fgt - 1 : fgt) - p - 2);
                    if (selfClosing) { p = fgt + 1; continue; }
                    int fend = raw.IndexOf("</f>", fgt, StringComparison.Ordinal);
                    c.FText = XmlUnescape(raw.Substring(fgt + 1, fend - fgt - 1));
                    p = fend + 4;
                }
                else if (string.CompareOrdinal(raw, p, "<v>", 0, 3) == 0)
                {
                    int vend = raw.IndexOf("</v>", p, StringComparison.Ordinal);
                    c.V = XmlUnescape(raw.Substring(p + 3, vend - p - 3));
                    p = vend + 4;
                }
                else if (string.CompareOrdinal(raw, p, "<v/>", 0, 4) == 0) p += 4;
                else { c.Simple = false; break; }
            }
            return c;
        }

        // Per sheet (reset in RewriteSheet): shared GetBalance formula id -> its master is exactly one GetBalance call.
        [ThreadStatic] private static Dictionary<string, bool>? sharedMasterIsSingleCall;
        private static Dictionary<string, bool> SharedMasterIsSingleCall => sharedMasterIsSingleCall ??= new Dictionary<string, bool>();

        private static string RewriteCell(string raw, Dictionary<(int row, int col), List<string>?>? balances, HashSet<string> glsenseSharedIds,
            Regex? deletedRefs, HashSet<string> refSharedIds, Result result)
        {
            Cell c = ParseCell(raw);
            string? fType = c.FAttr("t"), si = c.FAttr("si");
            bool isGlsense = c.FText.IndexOf("GLSENSE", StringComparison.OrdinalIgnoreCase) >= 0;
            if (isGlsense && fType == "shared" && si != null) glsenseSharedIds.Add(si);
            if (!isGlsense && fType == "shared" && si != null && glsenseSharedIds.Contains(si)) isGlsense = true; // dependent of a GLSense master
            if (!isGlsense) return deletedRefs == null || !c.Simple ? raw : PointAwayFromDeletedSheets(c, fType, si, deletedRefs, refSharedIds, result);
            if (!c.Simple) { result.Count("GLSense cells left unchanged (unexpected layout)"); return raw; }

            bool isBalance = c.FText.IndexOf("GLSENSE_GETBALANCE", StringComparison.OrdinalIgnoreCase) >= 0
                             || (fType == "shared" && c.FText.Length == 0 && balances != null && balances.ContainsKey(CellPosition(c.Attr("r"))));
            if (!isBalance)
            {
                result.Count("other GLSense formulas -> value");
                return ValueCell(c, CachedValue(c));
            }

            List<string>? values = null;
            bool known = balances != null && balances.TryGetValue(CellPosition(c.Attr("r")), out values);
            if (!known)
            {
                // Not part of this snapshot's request (a sheet kept only because the snapshot sheet refers to it).
                result.Count("balance cells on referenced sheets -> current value");
                return ValueCell(c, CachedValue(c));
            }
            if (values == null) { result.Count("balance cells without a balance"); return ValueCell(c, UnresolvedText); }
            if (c.FText.Length == 0)
            {
                // Dependent of a shared GetBalance formula: same shape as its master. A single call takes its own balance;
                // a mixed shape can't be rebuilt without shifting the master's references, so it keeps its current value.
                bool single = si != null && SharedMasterIsSingleCall.TryGetValue(si, out bool s) && s;
                if (single && values.Count == 1 && double.TryParse(values[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double dep))
                { result.Count("balance cells -> number"); return ValueCell(c, dep); }
                if (single && values.Count == 1) { result.Count("balance cells -> text"); return ValueCell(c, values[0]); }
                result.Count("shared mixed balance cells -> current value");
                return ValueCell(c, CachedValue(c));
            }

            var spans = BalanceCallSpans(c.FText);
            if (fType == "shared" && si != null)
                SharedMasterIsSingleCall[si] = spans.Count == 1 && spans[0].start == 0 && spans[0].end == c.FText.Length;
            if (spans.Count != values.Count) { result.Count("balance cells without a balance"); return ValueCell(c, UnresolvedText); }

            var numbers = new double[values.Count];
            for (int i = 0; i < values.Count; i++)
            {
                if (!double.TryParse(values[i], NumberStyles.Float, CultureInfo.InvariantCulture, out numbers[i]))
                {
                    result.Count("balance cells -> text"); // no-data option or a message: shown as text, as the refreshed cell shows it
                    return ValueCell(c, values[i]);
                }
            }
            if (spans.Count == 1 && spans[0].start == 0 && spans[0].end == c.FText.Length)
            {
                result.Count("balance cells -> number");
                return ValueCell(c, numbers[0]);
            }
            // More than the call(s): keep the file's own formula text (its _xlfn./_xll. prefixes stay valid), with
            // only the GetBalance calls replaced - right to left so the earlier positions stay valid.
            var sb = new StringBuilder(c.FText);
            for (int i = spans.Count - 1; i >= 0; i--)
            {
                string literal = numbers[i].ToString("R", CultureInfo.InvariantCulture);
                if (numbers[i] < 0) literal = "(" + literal + ")";
                sb.Remove(spans[i].start, spans[i].end - spans[i].start).Insert(spans[i].start, literal);
            }
            result.Count("balance cells -> formula with values");
            return $"<c{KeptAttrs(c)}><f>{Escape(sb.ToString())}</f></c>"; // no cached <v>: Excel calculates it on open
        }

        /// <summary>An ordinary formula on the kept sheet: references to deleted sheets become #REF! and the cell shows the
        /// #REF! error - what Excel's own sheet delete produced on the old hidden-Excel path. Anything else is untouched.</summary>
        private static string PointAwayFromDeletedSheets(Cell c, string? fType, string? si, Regex deletedRefs, HashSet<string> refSharedIds, Result result)
        {
            string fAttrs = c.FAttrs.TrimEnd('/', ' ');
            if (c.FText.Length == 0)
            {
                // Dependent of a shared formula: it follows its master's (rewritten) formula; only its stored value changes.
                if (fType != "shared" || si == null || !refSharedIds.Contains(si)) return c.Raw;
                return $"<c{KeptAttrs(c)} t=\"e\"><f{fAttrs}/><v>#REF!</v></c>";
            }
            if (!deletedRefs.IsMatch(c.FText)) return c.Raw;
            if (fType == "shared" && si != null) refSharedIds.Add(si);
            result.Count("formulas referring to deleted sheets -> #REF!");
            return $"<c{KeptAttrs(c)} t=\"e\"><f{fAttrs}>{Escape(deletedRefs.Replace(c.FText, "#REF!"))}</f><v>#REF!</v></c>";
        }

        /// <summary>[start, end) of each GLSense_GetBalance(...) call (with an _xll. prefix if present), left to right.</summary>
        internal static List<(int start, int end)> BalanceCallSpans(string f)
        {
            var spans = new List<(int, int)>();
            int from = 0;
            while (true)
            {
                int name = f.IndexOf(BalanceName, from, StringComparison.OrdinalIgnoreCase);
                if (name < 0) break;
                int start = name >= 5 && string.Compare(f, name - 5, "_xll.", 0, 5, StringComparison.OrdinalIgnoreCase) == 0 ? name - 5 : name;
                int depth = 0, end = -1;
                bool quoted = false;
                for (int i = name + BalanceName.Length - 1; i < f.Length; i++)
                {
                    char ch = f[i];
                    if (ch == '"') quoted = !quoted;
                    else if (!quoted && ch == '(') depth++;
                    else if (!quoted && ch == ')' && --depth == 0) { end = i + 1; break; }
                }
                if (end < 0) break;
                spans.Add((start, end));
                from = end;
            }
            return spans;
        }

        private static (int row, int col) CellPosition(string? r)
        {
            int col = 0, i = 0;
            if (r == null) return (0, 0);
            for (; i < r.Length && char.IsLetter(r[i]); i++) col = col * 26 + (char.ToUpperInvariant(r[i]) - 'A' + 1);
            int.TryParse(r.Substring(i), NumberStyles.Integer, CultureInfo.InvariantCulture, out int row);
            return (row, col);
        }

        private static object CachedValue(Cell c)
        {
            switch (c.Attr("t"))
            {
                case "str": return c.V ?? "";
                case "e": return "#GETTING_DATA"; // as the hidden-Excel step wrote errors
                case "b": return c.V == "1";
                default: return double.TryParse(c.V, NumberStyles.Float, CultureInfo.InvariantCulture, out double d) ? (object)d : "";
            }
        }

        // Every attribute of <c> except t (set again) and cm (dynamic-array metadata of the removed formula).
        private static string KeptAttrs(Cell c)
        {
            var sb = new StringBuilder();
            foreach (var kv in c.Attrs) if (kv.Key != "t" && kv.Key != "cm") sb.Append(' ').Append(kv.Key).Append("=\"").Append(kv.Value).Append('"');
            return sb.ToString();
        }

        private static string ValueCell(Cell c, object value)
        {
            string attrs = KeptAttrs(c);
            switch (value)
            {
                case double d: return $"<c{attrs}><v>{d.ToString("R", CultureInfo.InvariantCulture)}</v></c>";
                case bool b: return $"<c{attrs} t=\"b\"><v>{(b ? 1 : 0)}</v></c>";
                default: return $"<c{attrs} t=\"inlineStr\"><is><t xml:space=\"preserve\">{Escape(value?.ToString() ?? "")}</t></is></c>";
            }
        }

        // ---------- small helpers ----------

        private static string? Find(List<KeyValuePair<string, string>> list, string name)
        {
            foreach (var kv in list) if (kv.Key == name) return kv.Value;
            return null;
        }

        // name="value" pairs; values stay XML-escaped (they are written back as they were).
        private static List<KeyValuePair<string, string>> ParseAttrs(string s)
        {
            var list = new List<KeyValuePair<string, string>>(4);
            int i = 0;
            while (i < s.Length)
            {
                while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
                int eq = s.IndexOf('=', i); if (eq < 0) break;
                int q1 = s.IndexOf('"', eq); if (q1 < 0) break;
                int q2 = s.IndexOf('"', q1 + 1); if (q2 < 0) break;
                list.Add(new KeyValuePair<string, string>(s.Substring(i, eq - i).Trim(), s.Substring(q1 + 1, q2 - q1 - 1)));
                i = q2 + 1;
            }
            return list;
        }

        private static string Escape(string s) => System.Security.SecurityElement.Escape(s) ?? string.Empty;

        private static string XmlUnescape(string s)
        {
            if (s.IndexOf('&') < 0) return s;
            var sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] != '&') { sb.Append(s[i]); continue; }
                int semi = s.IndexOf(';', i);
                string ent = semi > i ? s.Substring(i + 1, semi - i - 1) : "";
                switch (ent)
                {
                    case "quot": sb.Append('"'); break;
                    case "amp": sb.Append('&'); break;
                    case "lt": sb.Append('<'); break;
                    case "gt": sb.Append('>'); break;
                    case "apos": sb.Append('\''); break;
                    default:
                        if (ent.StartsWith("#x", StringComparison.Ordinal)) sb.Append(char.ConvertFromUtf32(Convert.ToInt32(ent.Substring(2), 16)));
                        else if (ent.StartsWith("#", StringComparison.Ordinal)) sb.Append(char.ConvertFromUtf32(int.Parse(ent.Substring(1), CultureInfo.InvariantCulture)));
                        else { sb.Append('&'); continue; }
                        break;
                }
                i = semi;
            }
            return sb.ToString();
        }

        private static string PartPath(string baseDir, string target)
        {
            if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
            var parts = new List<string>(baseDir.TrimEnd('/').Split('/'));
            foreach (string seg in target.Split('/'))
            {
                if (seg == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
                else if (seg != "." && seg.Length > 0) parts.Add(seg);
            }
            return string.Join("/", parts);
        }

        private static XDocument Load(ZipArchive zip, string name)
        {
            var entry = zip.GetEntry(name) ?? throw new InvalidDataException($"'{name}' is missing from the workbook package.");
            using (var s = entry.Open()) return XDocument.Load(s);
        }

        private static void Save(ZipArchive zip, string name, XDocument doc)
        {
            using (var s = zip.CreateEntry(name, CompressionLevel.Fastest).Open())
            using (var w = XmlWriter.Create(s, new XmlWriterSettings { Encoding = new UTF8Encoding(false) })) doc.Save(w);
        }
    }
#nullable restore
}
