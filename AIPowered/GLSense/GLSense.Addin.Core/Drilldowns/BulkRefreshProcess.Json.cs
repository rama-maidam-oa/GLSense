// BulkRefreshProcess.Json.cs - JSON Sheet/Book Refresh (branch ribbon-compact).
//
// Replaces steps 1-4 of the file path (SaveCopyAs + zip + upload + extract + DataTable) with:
//   extract GLSense_GetBalance calls from the live workbook -> resolve their arguments to values (bulk reads, no
//   recalculation) -> build the UDF's own cache key (UdfDispatcher.BuildGetBalanceKey) and the same request DTO the
//   single-cell refresh sends (BalanceDto.CreateFromXllParameters) -> POST one request per worksheet (in parallel for
//   Book Refresh) to /rest/secure/finance/sheet-balances-json -> fill CalculatedBalances + PreComputedBalances
//   exactly like the file path.
// Step 5 (write-back recalculation, FormulaCache, SQLite on save) is unchanged and shared. Snapshot keeps the file path.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Udf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
#nullable enable
    public partial class BulkRefreshProcess
    {
        private const string GetBalanceName = "GLSENSE_GETBALANCE";
        // Book Refresh: worksheets sent at the same time (each holds its JSON body and response in memory).
        private const int MaxParallelSheetRequests = 4;
        // Larger worksheets are split so one JSON body stays far below .NET's string size limit.
        private const int MaxRequestsPerPost = 100000;
        // OISR-22975: request bodies are sent gzip-compressed until a server rejects that (see SendJsonBatchAsync).
        private volatile bool _gzipRequests = true;

        // Cell reference as written in a formula argument: A1, $A$1, Sheet1!A1, 'My Sheet'!$A$1.
        private static readonly Regex CellRefRegex = new Regex(
            @"^(?:(?:'(?<q>(?:[^']|'')+)'|(?<s>[^'!\s()=,&""+\-*/^<>]+))!)?\$?(?<c>[A-Za-z]{1,3})\$?(?<r>\d{1,7})$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex QuotedLiteralRegex = new Regex(@"^""(?:[^""]|"""")*""$", RegexOptions.Compiled);

        /// <summary>One GLSense_GetBalance call in one cell, after argument resolution.</summary>
        private sealed class JsonCall
        {
            public string Sheet = "";
            public string Address = "";   // $C$5 style, like the file path's excelCell
            public string CallText = "";
            public object?[] Args = Array.Empty<object?>(); // 32 values exactly as the UDF would receive them
        }

        /// <summary>Snapshot only: one cell holding GetBalance calls, and its calls (indexes into the call list, left to right).</summary>
        private sealed class SnapshotCell
        {
            public string Sheet = "";
            public int Row, Col;
            public readonly List<int> Calls = new List<int>();
            public bool Unresolved; // a call whose arguments couldn't be resolved (the refreshed cell wouldn't get a balance either)
        }

        private List<SnapshotCell>? _snapshotCells;
        // Snapshot: per worksheet, per GetBalance cell, the balance of each call (null = none) - see BuildSnapshotBalances.
        private Dictionary<string, Dictionary<(int row, int col), List<string>?>>? _snapshotBalances;

        private Task<Exception?> Step_JsonBalances(CancellationToken token)
            => RunLongRunningTask(token, JsonBalancesAsync);

        private async Task JsonBalancesAsync(CancellationToken token)
        {
            LogDebug("RefreshPerf: mode = JSON (one request per worksheet)");
            await MessageProgressWindowAsync("Reading balance formulas from the workbook.");

            // 1. Extract + resolve on the Excel thread (bulk reads only; nothing is recalculated).
            _snapshotCells = _isSnapshot ? new List<SnapshotCell>() : null;
            List<JsonCall> calls;
            using (PerformanceHelper.MeasureExecutionTime("RefreshPerf: J1 extract + resolve arguments (Excel thread)"))
                calls = await RunExcelAsync(() => ExtractAndResolveCalls(token));
            token.ThrowIfCancellationRequested();

            // 2. Keys (the UDF's own key builder) + de-duplication.
            await MessageProgressWindowAsync($"Preparing {calls.Count:N0} balance requests.");
            var keyOfCall = new string[calls.Count];
            var uniqueKeys = new List<string>();
            var uniqueParams = new List<object[]>();
            var uniqueSheet = new List<string>(); // first worksheet using the key: a key shared by sheets is requested once
            var indexOfKey = new Dictionary<string, int>(StringComparer.Ordinal);
            using (PerformanceHelper.MeasureExecutionTime("RefreshPerf: J2 build cache keys + de-duplicate"))
            {
                for (int i = 0; i < calls.Count; i++)
                {
                    string key = UdfDispatcher.BuildGetBalanceKey(calls[i].Args!, out object[] parameters, out _);
                    keyOfCall[i] = key;
                    if (string.IsNullOrEmpty(key) || indexOfKey.ContainsKey(key)) continue;
                    indexOfKey[key] = uniqueKeys.Count;
                    uniqueKeys.Add(key);
                    uniqueParams.Add(parameters);
                    uniqueSheet.Add(calls[i].Sheet);
                }
            }
            LogDebug($"RefreshPerf: JSON calls = {calls.Count:N0}, distinct requests = {uniqueKeys.Count:N0}");

            // 3. Request DTOs - same builder as the single-cell refresh (POST /balance).
            var dtos = new List<Balances>(uniqueKeys.Count);
            var dtoSheet = new List<string>(uniqueKeys.Count);
            var dtoUnique = new List<int>(uniqueKeys.Count); // index into uniqueKeys, for the progress text's formula count
            int dtoFailures = 0;
            using (PerformanceHelper.MeasureExecutionTime("RefreshPerf: J3 build request DTOs"))
            using (BalanceDto.BeginBatchLookupCache()) // ledger/segment/activity lookups once per refresh, not per request (no await inside: same thread)
            {
                for (int u = 0; u < uniqueKeys.Count; u++)
                {
                    token.ThrowIfCancellationRequested();
                    var stringArgs = uniqueParams[u].Select(p => p?.ToString() ?? string.Empty).ToList();
                    try
                    {
                        if (BalanceDto.CreateFromXllParameters("J" + u.ToString(CultureInfo.InvariantCulture), stringArgs) is Balances dto)
                        {
                            dto.excelSheet = "GLSENSE_JSON";
                            // Adaptive memory keys on inputFormula and fails without it; the resolved call text is unique per key.
                            dto.inputFormula = "=" + GetBalanceName + "(" + string.Join(",", stringArgs.Select(a => "\"" + a.Trim() + "\"")) + ")";
                            dtos.Add(dto);
                            dtoSheet.Add(uniqueSheet[u]);
                            dtoUnique.Add(u);
                        }
                        else dtoFailures++;
                    }
                    catch (Exception ex)
                    {
                        if (dtoFailures++ < 5) LogWarn($"RefreshPerf: request DTO build failed ({ex.GetType().Name}: {ex.Message})");
                    }
                }
            }
            LogDebug($"RefreshPerf: request DTOs = {dtos.Count:N0}, build failures = {dtoFailures:N0}");

            // 4. One request per worksheet, up to MaxParallelSheetRequests at a time; results are correlated by the "J<n>" id
            //    carried in excelCell (response order is not preserved). The first failing request cancels the rest.
            // Formulas served by each distinct balance (a balance repeated on several sheets serves all of them).
            var formulasOfUnique = new int[uniqueKeys.Count];
            foreach (string key in keyOfCall)
                if (!string.IsNullOrEmpty(key) && indexOfKey.TryGetValue(key, out int ku)) formulasOfUnique[ku]++;
            var sheetOrder = calls.Select(c => c.Sheet).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            int sheetsInScope = sheetOrder.Count;
            // Worksheets (1-based positions) whose formulas each distinct balance serves: a request built under the first
            // sheet using its balances also fills the other sheets repeating them, so the progress text names them all.
            var sheetsOfUnique = new SortedSet<int>[uniqueKeys.Count];
            var sheetNoOf = sheetOrder.Select((s, i) => (s, i)).ToDictionary(x => x.s, x => x.i + 1, StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < calls.Count; i++)
                if (!string.IsNullOrEmpty(keyOfCall[i]) && indexOfKey.TryGetValue(keyOfCall[i], out int su))
                    (sheetsOfUnique[su] ??= new SortedSet<int>()).Add(sheetNoOf[calls[i].Sheet]);

            var batches = new List<(string sheet, List<Balances> dtos, int formulas, string progress)>();
            foreach (var group in dtos.Select((d, i) => (d, sheet: dtoSheet[i], u: dtoUnique[i])).GroupBy(x => x.sheet, StringComparer.OrdinalIgnoreCase))
            {
                var all = group.ToList();
                int parts = (all.Count + MaxRequestsPerPost - 1) / MaxRequestsPerPost;
                int sheetNo = sheetOrder.FindIndex(s => string.Equals(s, group.Key, StringComparison.OrdinalIgnoreCase)) + 1;
                for (int s = 0, part = 1; s < all.Count; s += MaxRequestsPerPost, part++)
                {
                    var slice = all.GetRange(s, Math.Min(MaxRequestsPerPost, all.Count - s));
                    int formulas = slice.Sum(x => formulasOfUnique[x.u]);
                    string partText = parts > 1 ? $" (part {part} of {parts})" : string.Empty;
                    var served = new SortedSet<int> { sheetNo };
                    foreach (var x in slice) if (sheetsOfUnique[x.u] != null) served.UnionWith(sheetsOfUnique[x.u]);
                    batches.Add((group.Key, slice.Select(x => x.d).ToList(), formulas,
                        $"{SheetsText(served)} of {sheetsInScope:N0}{partText}: waiting for the server to return {slice.Count:N0} balances for {formulas:N0} formulas."));
                }
            }

            var valueOfKey = new Dictionary<string, (string value, bool cache)>(StringComparer.Ordinal);
            int done = 0;
            using (PerformanceHelper.MeasureExecutionTime($"RefreshPerf: J4 send {batches.Count:N0} worksheet request(s) to /sheet-balances-json"))
            using (var client = CreateJsonHttpClient())
            using (var gate = new SemaphoreSlim(MaxParallelSheetRequests))
            using (var failFast = CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                string url = $"{_loginUrl}/rest/secure/finance/sheet-balances-json?cubeId={_state.SelectedCube.CubeId}&coaid={_state.SelectedLedger.CoaId}";
                // Requests group each distinct balance under the FIRST worksheet using it, so a book whose sheets repeat
                // the same formulas sends fewer requests than it has worksheets (e.g. one request: "Worksheet 1 of 3").
                LogDebug($"RefreshPerf: J4 scope = {sheetsInScope:N0} worksheet(s), {calls.Count:N0} formulas, {dtos.Count:N0} distinct balances, {batches.Count:N0} request(s)");
                var tasks = batches.Select(async (b, n) =>
                {
                    await gate.WaitAsync(failFast.Token).ConfigureAwait(false);
                    try
                    {
                        await MessageProgressWindowAsync(b.progress);
                        await SendJsonBatchAsync(client, url, b.dtos, $"{n + 1}/{batches.Count} '{b.sheet}'", uniqueKeys, valueOfKey, failFast.Token).ConfigureAwait(false);
                        int finished = Interlocked.Increment(ref done);
                        await MessageProgressWindowAsync(batches.Count == 1
                            ? $"Balances received for {b.formulas:N0} formulas."
                            : $"Balances received for {finished:N0} of {batches.Count:N0} requests (last: worksheet '{b.sheet}', {b.formulas:N0} formulas).");
                    }
                    catch { failFast.Cancel(); throw; }
                    finally { gate.Release(); }
                }).ToList();

                try { await Task.WhenAll(tasks); }
                catch
                {
                    token.ThrowIfCancellationRequested(); // a user cancel is reported as a cancel
                    // Report the request that actually failed, not the ones cancelled because of it.
                    Exception? first = tasks.Where(t => t.IsFaulted).Select(t => t.Exception!.InnerException!)
                                            .FirstOrDefault(e => e is not OperationCanceledException);
                    if (first != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(first).Throw();
                    throw;
                }
            }
            int cached = valueOfKey.Values.Count(v => v.cache);
            LogDebug($"RefreshPerf: balances returned = {valueOfKey.Count:N0} of {dtos.Count:N0} requests, adaptive-memory hits = {cached:N0}");

            if (_isSnapshot)
            {
                // Snapshot: the live workbook isn't recalculated, so no CalculatedBalances/PreComputedBalances (as before).
                using (PerformanceHelper.MeasureExecutionTime("RefreshPerf: J5 build snapshot cell contents"))
                    BuildSnapshotBalances(keyOfCall, valueOfKey);
                return;
            }

            // 5. Store exactly like the file path: CalculatedBalances (MemCells/Reset) + PrepareBatchCache (PreComputedBalances).
            using (PerformanceHelper.MeasureExecutionTime("RefreshPerf: J5 fill CalculatedBalances + PreComputedBalances"))
            {
                var dt = new DataTable();
                dt.Columns.Add("excelSheet", typeof(string));
                dt.Columns.Add("excelCell", typeof(string));
                dt.Columns.Add("inputFormula", typeof(string));
                dt.Columns.Add("formulaKey", typeof(string));
                dt.Columns.Add("balanceValue", typeof(string));
                dt.Columns.Add("cache", typeof(bool));
                for (int i = 0; i < calls.Count; i++)
                {
                    if (!valueOfKey.TryGetValue(keyOfCall[i] ?? string.Empty, out var v)) continue; // miss -> "Click Refresh...", as in the file path
                    dt.Rows.Add(calls[i].Sheet, calls[i].Address, calls[i].CallText, keyOfCall[i], v.value, v.cache);
                }
                _state.CalculatedBalances = dt;
                if (dt.Rows.Count > 0) PrepareBatchCache();
                LogDebug($"RefreshPerf: CalculatedBalances rows = {dt.Rows.Count:N0}, PreComputedBalances = {_state.PreComputedBalances?.Count ?? 0:N0}");
            }
        }

        /// <summary>"Worksheet 2", "Worksheets 1-3", "Worksheets 1, 3, 5" (runs collapsed; at most 6 items, then "+n").</summary>
        internal static string SheetsText(SortedSet<int> sheets)
        {
            if (sheets.Count == 1) return $"Worksheet {sheets.Min:N0}";
            var parts = new List<string>();
            int? runStart = null, prev = null;
            foreach (int n in sheets.Concat(new[] { int.MaxValue }))
            {
                if (prev.HasValue && n == prev.Value + 1) { prev = n; continue; }
                if (runStart.HasValue) parts.Add(prev == runStart ? $"{runStart:N0}" : $"{runStart:N0}-{prev:N0}");
                runStart = prev = n;
            }
            string list = parts.Count <= 6 ? string.Join(", ", parts) : string.Join(", ", parts.Take(6)) + $" +{parts.Count - 6:N0} more";
            return $"Worksheets {list}";
        }

        // ---------- snapshot (JSON path) ----------

        /// <summary>Per GetBalance cell: the balance of each of its calls, left to right (null = no balance). The
        /// OpenXML step (SnapshotOpenXml) turns these into the cell's value or its formula with values.</summary>
        private void BuildSnapshotBalances(string[] keyOfCall, Dictionary<string, (string value, bool cache)> valueOfKey)
        {
            var result = new Dictionary<string, Dictionary<(int row, int col), List<string>?>>(StringComparer.OrdinalIgnoreCase);
            int complete = 0, unresolved = 0;
            var unresolvedSamples = new List<string>();

            foreach (SnapshotCell cell in _snapshotCells ?? new List<SnapshotCell>())
            {
                List<string>? values = null;
                if (!cell.Unresolved)
                {
                    values = new List<string>(cell.Calls.Count);
                    foreach (int i in cell.Calls)
                    {
                        if (!valueOfKey.TryGetValue(keyOfCall[i] ?? string.Empty, out var v)) { values = null; break; }
                        values.Add(v.value);
                    }
                }
                if (values == null) { if (unresolved++ < 5) unresolvedSamples.Add($"'{cell.Sheet}'!{A1(cell.Row, cell.Col)}"); }
                else complete++;

                if (!result.TryGetValue(cell.Sheet, out var sheetCells))
                    result[cell.Sheet] = sheetCells = new Dictionary<(int row, int col), List<string>?>();
                sheetCells[(cell.Row, cell.Col)] = values;
            }

            _snapshotBalances = result;
            _snapshotCells = null;
            LogDebug($"RefreshPerf: snapshot balance cells = {complete + unresolved:N0} (with balances {complete:N0}, without {unresolved:N0})");
            if (unresolved > 0)
                LogWarn($"Snapshot: {unresolved:N0} balance cell(s) got no balance and show \"Click Refresh...\" in the snapshot, e.g. {string.Join(", ", unresolvedSamples)}.");
        }

        // ---------- extraction + argument resolution (Excel thread) ----------

        private List<JsonCall> ExtractAndResolveCalls(CancellationToken token)
        {
            var calls = new List<JsonCall>();
            var valueCache = new Dictionary<string, (object?[,]? data, int row0, int col0, object? single)>(StringComparer.OrdinalIgnoreCase);
            var evalCache = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            int formulaCells = 0, skipped = 0, evaluated = 0;

            foreach (Excel.Worksheet sheet in TargetSheets())
            {
                token.ThrowIfCancellationRequested();
                string sheetName = sheet.Name;
                Excel.Range? formulas = null;
                try
                {
                    try { formulas = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas); }
                    catch (COMException) { continue; } // no formulas on this sheet

                    foreach (Excel.Range area in formulas.Areas)
                    {
                        try
                        {
                            int row0 = area.Row, col0 = area.Column;
                            object raw = area.Formula;
                            if (raw is string single)
                            {
                                formulaCells++;
                                AddCalls(sheet, sheetName, row0, col0, single);
                            }
                            else if (raw is object[,] grid)
                            {
                                int rows = grid.GetLength(0), cols = grid.GetLength(1);
                                for (int r = 1; r <= rows; r++)
                                    for (int c = 1; c <= cols; c++)
                                        if (grid[r, c] is string f && f.Length > 0)
                                        {
                                            formulaCells++;
                                            AddCalls(sheet, sheetName, row0 + r - 1, col0 + c - 1, f);
                                        }
                            }
                        }
                        finally { Marshal.ReleaseComObject(area); }
                    }
                }
                finally
                {
                    if (formulas != null) Marshal.ReleaseComObject(formulas);
                    if (!ReferenceEquals(sheet, _worksheet)) Marshal.ReleaseComObject(sheet); // step 5 still needs _worksheet
                }
            }

            LogDebug($"RefreshPerf: JSON extract: formula cells = {formulaCells:N0}, GetBalance calls = {calls.Count:N0}, skipped calls = {skipped:N0}, Evaluate() fallbacks = {evaluated:N0}, sheets read for references = {valueCache.Count:N0}");
            return calls;

            void AddCalls(Excel.Worksheet sheet, string sheetName, int row, int col, string formula)
            {
                if (formula.IndexOf(GetBalanceName, StringComparison.OrdinalIgnoreCase) < 0) return;
                SnapshotCell? cell = _snapshotCells == null ? null : new SnapshotCell { Sheet = sheetName, Row = row, Col = col };
                foreach (string call in BalanceNormalizer.ExtractUdfCalls(formula, GetBalanceName))
                {
                    var args = ResolveArgs(sheet, sheetName, call);
                    if (args == null) { skipped++; if (cell != null) cell.Unresolved = true; continue; }
                    args[31] = $"R{row}C{col}";
                    calls.Add(new JsonCall { Sheet = sheetName, Address = A1(row, col), CallText = call, Args = args });
                    cell?.Calls.Add(calls.Count - 1);
                }
                if (cell != null) _snapshotCells!.Add(cell);
            }

            object?[]? ResolveArgs(Excel.Worksheet sheet, string sheetName, string call)
            {
                int open = call.IndexOf('('), close = call.LastIndexOf(')');
                if (open < 0 || close <= open) return null;
                List<string> tokens = SplitTopLevelArgs(call.Substring(open + 1, close - open - 1));

                var args = new object?[32];
                for (int i = 0; i < 31; i++)
                {
                    object? v = i < tokens.Count ? ResolveToken(sheet, sheetName, tokens[i]) : null;
                    if (v is int) return null;                          // Excel error value (#N/A, #REF!...) - the UDF never looks it up
                    if (v == null)
                    {
                        if (i <= 10) return null;                       // missing mandatory argument: the host returns "#Error: Missing ..."
                        v = string.Empty;                               // missing segment: the host substitutes ""
                    }
                    args[i] = v;
                }
                return args;
            }

            object? ResolveToken(Excel.Worksheet sheet, string sheetName, string token)
            {
                string t = token.Trim();
                if (t.Length == 0) return null;
                if (QuotedLiteralRegex.IsMatch(t)) return t.Substring(1, t.Length - 2).Replace("\"\"", "\"");
                if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out double d)) return d;
                if (t.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return true;
                if (t.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return false;

                Match m = CellRefRegex.Match(t);
                if (m.Success)
                {
                    string target = m.Groups["q"].Success ? m.Groups["q"].Value.Replace("''", "'")
                                  : m.Groups["s"].Success ? m.Groups["s"].Value : sheetName;
                    return CellValue(target, ColumnNumber(m.Groups["c"].Value), int.Parse(m.Groups["r"].Value, CultureInfo.InvariantCulture));
                }

                // Names, & joins, functions: let Excel evaluate the expression once (no workbook recalculation).
                string evalKey = sheetName + "\u0001" + t;
                if (!evalCache.TryGetValue(evalKey, out object? value))
                {
                    evaluated++;
                    object result = sheet.Evaluate(t);
                    value = result is object[,] arr ? arr[arr.GetLowerBound(0), arr.GetLowerBound(1)] : result;
                    evalCache[evalKey] = value;
                }
                return value;
            }

            object? CellValue(string sheetName, int col, int row)
            {
                if (!valueCache.TryGetValue(sheetName, out var block))
                {
                    block = (null, 1, 1, null);
                    Excel.Worksheet? ws = null;
                    Excel.Range? used = null;
                    try
                    {
                        ws = (Excel.Worksheet)_workbook!.Worksheets[sheetName];
                        used = ws.UsedRange;
                        object v = used.Value2;
                        block = v is object[,] grid ? (grid, used.Row, used.Column, null) : (null, used.Row, used.Column, v);
                    }
                    catch (COMException) { /* unknown sheet: every reference to it resolves to null */ }
                    finally
                    {
                        if (used != null) Marshal.ReleaseComObject(used);
                        if (ws != null) Marshal.ReleaseComObject(ws);
                    }
                    valueCache[sheetName] = block;
                }

                if (block.data == null)
                    return row == block.row0 && col == block.col0 ? block.single : null;
                int r = row - block.row0 + 1, c = col - block.col0 + 1;
                return r >= 1 && c >= 1 && r <= block.data.GetLength(0) && c <= block.data.GetLength(1) ? block.data[r, c] : null;
            }
        }

        private IEnumerable<Excel.Worksheet> TargetSheets()
        {
            if (_isTypeSheet)
            {
                if (_worksheet != null) yield return _worksheet;
                yield break;
            }
            foreach (Excel.Worksheet ws in _workbook!.Worksheets) yield return ws;
        }

        /// <summary>Splits a formula's argument list on top-level commas, keeping empty arguments in place.</summary>
        internal static List<string> SplitTopLevelArgs(string inner)
        {
            var args = new List<string>();
            var cur = new StringBuilder();
            int depth = 0;
            bool quoted = false;
            for (int i = 0; i < inner.Length; i++)
            {
                char ch = inner[i];
                if (ch == '"') quoted = !quoted;                     // "" inside a literal toggles twice - net effect unchanged
                else if (!quoted && (ch == '(' || ch == '{')) depth++;
                else if (!quoted && (ch == ')' || ch == '}')) depth--;
                if (ch == ',' && !quoted && depth == 0) { args.Add(cur.ToString()); cur.Clear(); }
                else cur.Append(ch);
            }
            args.Add(cur.ToString());
            return args;
        }

        private static int ColumnNumber(string letters)
        {
            int n = 0;
            foreach (char ch in letters.ToUpperInvariant()) n = n * 26 + (ch - 'A' + 1);
            return n;
        }

        private static string A1(int row, int col)
        {
            var s = new StringBuilder();
            for (int n = col; n > 0; n = (n - 1) / 26) s.Insert(0, (char)('A' + (n - 1) % 26));
            return "$" + s + "$" + row.ToString(CultureInfo.InvariantCulture);
        }

        // ---------- HTTP ----------

        private HttpClient CreateJsonHttpClient()
        {
            var handler = new HttpClientHandler
            {
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12 |
                               System.Security.Authentication.SslProtocols.Tls13,
                ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate,
                MaxConnectionsPerServer = MaxParallelSheetRequests, // .NET Framework allows only 2 per server by default
                // OISR-22975: sends Accept-Encoding and unzips a gzip response; a plain response still works.
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };
            var client = new HttpClient(handler) { Timeout = Timeout.InfiniteTimeSpan };
            client.DefaultRequestHeaders.ExpectContinue = false;
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _authToken);
            return client;
        }

        private static byte[] Gzip(byte[] data)
        {
            using (var ms = new MemoryStream())
            {
                using (var gz = new GZipStream(ms, CompressionLevel.Fastest, leaveOpen: true))
                    gz.Write(data, 0, data.Length);
                return ms.ToArray();
            }
        }

        private static async Task<HttpResponseMessage> PostJsonAsync(
            HttpClient client, string url, byte[] body, bool gzip, string label, CancellationToken token)
        {
            using (var content = new ByteArrayContent(body))
            {
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json") { CharSet = "utf-8" };
                if (gzip) content.Headers.ContentEncoding.Add("gzip");
                using (PerformanceHelper.MeasureExecutionTime($"RefreshPerf: J4 {label} upload + server processing"))
                    return await client.PostAsync(url, content, token).ConfigureAwait(false);
            }
        }

        private async Task SendJsonBatchAsync(
            HttpClient client, string url, List<Balances> batch, string label, List<string> uniqueKeys,
            Dictionary<string, (string value, bool cache)> valueOfKey, CancellationToken token)
        {
            byte[]? body;
            using (PerformanceHelper.MeasureExecutionTime($"RefreshPerf: J4 {label} serialize"))
                body = JsonSerializer.SerializeToUtf8Bytes(new { balanceDtos = batch }, JsonGlobals.Options);

            HttpResponseMessage response;
            if (_gzipRequests)
            {
                byte[] zipped;
                using (PerformanceHelper.MeasureExecutionTime($"RefreshPerf: J4 {label} gzip request"))
                    zipped = Gzip(body);
                LogDebug($"RefreshPerf: request {label}: {batch.Count:N0} balances, {body.Length / (1024.0 * 1024.0):F2} MB JSON, {zipped.Length / (1024.0 * 1024.0):F2} MB gzip");
                response = await PostJsonAsync(client, url, zipped, gzip: true, label, token).ConfigureAwait(false);
                // A server without request decompression rejects the body: send plain JSON from now on.
                if (response.StatusCode == HttpStatusCode.BadRequest || response.StatusCode == HttpStatusCode.UnsupportedMediaType)
                {
                    LogWarn($"sheet-balances-json request {label}: server answered HTTP {(int)response.StatusCode} to a gzip request - resending uncompressed, and not compressing further requests.");
                    _gzipRequests = false;
                    response.Dispose();
                    response = await PostJsonAsync(client, url, body, gzip: false, label, token).ConfigureAwait(false);
                }
            }
            else
            {
                LogDebug($"RefreshPerf: request {label}: {batch.Count:N0} balances, {body.Length / (1024.0 * 1024.0):F2} MB JSON");
                response = await PostJsonAsync(client, url, body, gzip: false, label, token).ConfigureAwait(false);
            }
            body = null; // release the request bytes before the response is read

            string text;
            using (response)
            using (PerformanceHelper.MeasureExecutionTime($"RefreshPerf: J4 {label} download response"))
                text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            LogDebug($"RefreshPerf: request {label}: HTTP {(int)response.StatusCode}, response {text.Length / (1024.0 * 1024.0):F2} MB");
            // OISR-22830 pattern: full response to the Temp folder, preview + path in the log (debug only).
            string sheetHint = "sheet-balances-json_" + label.Replace('/', '-');
            if (_state.DebugLogs)
            {
                using (PerformanceHelper.MeasureExecutionTime($"RefreshPerf: J4 {label} save response to Temp (debug only)"))
                    LogDebug($"RefreshPerf: request {label} response: {GLSense.Shared.LargeLogText.ToLogText(text, ServiceLocator.Paths?.Temp, sheetHint)}");
            }

            JsonNode? root = null;
            try { root = JsonNode.Parse(text); } catch (JsonException) { /* reported below */ }
            string status = root?["status"]?.ToString() ?? string.Empty;
            if (!response.IsSuccessStatusCode || !status.Equals("success", StringComparison.OrdinalIgnoreCase) || root?["records"] is not JsonArray records)
            {
                // Always keep the evidence of a failed request, Debug on or off.
                if (!_state.DebugLogs)
                    LogWarn($"sheet-balances-json request {label} failed, response: {GLSense.Shared.LargeLogText.ToLogText(text, ServiceLocator.Paths?.Temp, sheetHint)}");
                string msg = root?["msg"]?.ToString() ?? (text.Length > 500 ? text.Substring(0, 500) : text);
                throw new Exception($"sheet-balances-json request {label} failed (HTTP {(int)response.StatusCode}, status '{status}'): {msg}");
            }

            using (PerformanceHelper.MeasureExecutionTime($"RefreshPerf: J4 {label} map results"))
            {
                foreach (JsonNode? rec in records)
                {
                    // excelCell carries "J<index into uniqueKeys>", set when the DTO was built.
                    string id = rec?["excelCell"]?.ToString() ?? string.Empty;
                    if (id.Length < 2 || id[0] != 'J' || !int.TryParse(id.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out int u) || u < 0 || u >= uniqueKeys.Count)
                        continue;
                    bool cache = (rec?["isCache"] ?? rec?["cache"]) is JsonValue cv && cv.TryGetValue(out bool b) && b;
                    lock (valueOfKey) valueOfKey[uniqueKeys[u]] = (ToInvariantText(rec?["balanceValue"]), cache);
                }
            }
        }

        /// <summary>Same text form the file path stores in CalculatedBalances.balanceValue.</summary>
        private static string ToInvariantText(JsonNode? node)
        {
            if (node is not JsonValue v) return node?.ToJsonString() ?? string.Empty;
            if (v.TryGetValue(out long l)) return l.ToString(CultureInfo.InvariantCulture);
            if (v.TryGetValue(out double d)) return d.ToString("R", CultureInfo.InvariantCulture);
            if (v.TryGetValue(out decimal m)) return m.ToString(CultureInfo.InvariantCulture);
            return v.TryGetValue(out string? s) ? s ?? string.Empty : v.ToJsonString();
        }

        // ---------- parity log (both paths) ----------

        /// <summary>Per sheet: sum of numeric formula results and count of "Click Refresh..." / errors / other text,
        /// so a file run and a JSON run on the same workbook can be compared line by line.</summary>
        private async Task LogRefreshParityAsync()
        {
            try
            {
                await RunExcelAsync(() =>
                {
                    foreach (Excel.Worksheet sheet in TargetSheets())
                    {
                        string name = sheet.Name;
                        Excel.Range? formulas = null;
                        double sum = 0; long numbers = 0, clickRefresh = 0, errors = 0, otherText = 0;
                        try
                        {
                            try { formulas = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas); }
                            catch (COMException) { continue; }
                            foreach (Excel.Range area in formulas.Areas)
                            {
                                try
                                {
                                    object raw = area.Value2;
                                    IEnumerable<object?> values = raw is object[,] g ? g.Cast<object?>() : new[] { raw };
                                    foreach (object? v in values)
                                    {
                                        switch (v)
                                        {
                                            case double d: sum += d; numbers++; break;
                                            case int: errors++; break;
                                            case string s when s.StartsWith("Click Refresh", StringComparison.OrdinalIgnoreCase): clickRefresh++; break;
                                            case string: otherText++; break;
                                        }
                                    }
                                }
                                finally { Marshal.ReleaseComObject(area); }
                            }
                            LogDebug($"RefreshPerf: PARITY [{name}] sum = {sum.ToString("R", CultureInfo.InvariantCulture)}, numbers = {numbers:N0}, Click Refresh = {clickRefresh:N0}, errors = {errors:N0}, other text = {otherText:N0}");
                        }
                        finally
                        {
                            if (formulas != null) Marshal.ReleaseComObject(formulas);
                            if (!ReferenceEquals(sheet, _worksheet)) Marshal.ReleaseComObject(sheet);
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                LogWarn($"RefreshPerf: parity log failed: {ex.Message}");
            }
        }
    }
}
