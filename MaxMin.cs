using System;
using System.IO;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace Hadi.Excel
{
    public static class Log
    {
        private static readonly string Path =
            System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Hadi.Excel.log");

        public static void Clear() { try { File.WriteAllText(Path, ""); } catch { } }
        public static void Write(string msg)
        {
            try { File.AppendAllText(Path, $"{DateTime.Now:HH:mm:ss.fff}  {msg}{Environment.NewLine}"); }
            catch { }
        }
    }

    public static class MaxMin
    {
        private const string HelperSheetName = "MaxMinHelper";
        private const int MaxHelperColumns = 40;

        private const int xlMarkerStyleCircle = 8;
        private const int xlNone = -4142;
        private const int msoFalse = 0;
        private const int xlLabelPositionAbove = 0;

        private const int ColorRed = 255;
        private const int ColorBlue = 16711680;

        public static void Toggle(string kind)
        {
            Log.Clear();
            Log.Write($"=== Toggle('{kind}') START ===");

            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic sel = app.Selection;

                dynamic series = TryGetSelectedSeries(sel);
                if (series == null)
                {
                    Log.Write("No series selected.");
                    System.Windows.Forms.MessageBox.Show(
                        "Please select a chart series first.\n\n" +
                        "Click directly on a line, marker, bar, or column that belongs to a series.",
                        "Hadi Tools",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    Log.Write("=== Toggle END (no series) ===");
                    return;
                }

                string sname = "";
                try { sname = (string)series.Name; } catch { }
                Log.Write("Selected series: " + sname);

                string formula = "";
                try { formula = (string)series.Formula; } catch { }
                Log.Write("Series formula: " + formula);

                string key = sname + "|" + kind;
                Log.Write("Helper key: " + key);

                dynamic chart = app.ActiveChart;
                if (chart == null) { Log.Write("No active chart."); return; }

                string[] sourceCells;
                string yRange;
                if (!TryParseSeriesFormula(formula, out sourceCells, out yRange))
                {
                    Log.Write("Parse failed.");
                    Log.Write("=== Toggle END (parse fail) ===");
                    return;
                }
                Log.Write("Parsed source cells: " + string.Join(",", sourceCells));
                Log.Write("Y-range: " + yRange);

                dynamic wb = app.ActiveWorkbook;
                dynamic originalSheet = app.ActiveSheet;

                dynamic helper = EnsureHelperSheet(wb);
                if (helper == null) { Log.Write("No helper sheet."); return; }

                int col = FindOrCreateColumn(helper, key);
                if (col <= 0) { Log.Write("No helper column."); return; }
                Log.Write("Helper column: " + col);

                if (HelperSeriesAtColumnExists(chart, col))
                {
                    Log.Write("Helper series at column present — toggling OFF.");
                    RemoveHelperSeriesAtColumn(chart, col);
                    Log.Write("=== Toggle END (toggled off) ===");
                    return;
                }

                double[] values = ReadSeriesValues(series);
                int extremeIdx = (kind == "Max") ? IndexOfMax(values) : IndexOfMin(values);
                Log.Write($"Extreme index for {kind}: {extremeIdx}");

                if (!WriteFormulaColumn(helper, col, sourceCells, yRange, kind))
                {
                    Log.Write("Write formula column failed.");
                    return;
                }

                dynamic newSeries = AddHelperSeries(chart, helper, col, sourceCells.Length, key);
                if (newSeries != null)
                {
                    FormatHelperSeries(chart, newSeries, kind);
                    AddDataLabelSimple(newSeries);
                }

                try { originalSheet.Activate(); } catch { }

                Log.Write("=== Toggle END (success) ===");
            }
            catch (Exception ex)
            {
                Log.Write("EXCEPTION: " + ex.GetType().FullName);
                Log.Write("MESSAGE : " + ex.Message);
                Log.Write("STACK   : " + ex.StackTrace);
                System.Windows.Forms.MessageBox.Show(
                    ex.Message, "Hadi Tools error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        // ---------- Ribbon state queries ----------

        public static bool HasSelectedSeries()
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic sel = app.Selection;
                if (sel == null) return false;
                try
                {
                    object raw = ((object)sel).GetType().InvokeMember(
                        "Values",
                        System.Reflection.BindingFlags.GetProperty,
                        null, sel, null);
                    return raw != null;
                }
                catch { return false; }
            }
            catch { return false; }
        }

        public static bool IsActive(string kind)
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic sel = app.Selection;
                if (sel == null) return false;

                object raw;
                try
                {
                    raw = ((object)sel).GetType().InvokeMember(
                        "Values",
                        System.Reflection.BindingFlags.GetProperty,
                        null, sel, null);
                }
                catch { return false; }
                if (raw == null) return false;

                string sname;
                try { sname = (string)sel.Name; } catch { return false; }

                dynamic chart = app.ActiveChart;
                if (chart == null) return false;

                dynamic wb = app.ActiveWorkbook;
                if (wb == null) return false;

                dynamic helper = FindHelperSheetNoCreate(wb);
                if (helper == null) return false;

                int col = FindColumnNoCreate(helper, sname + "|" + kind);
                if (col <= 0) return false;

                return HelperSeriesAtColumnExists(chart, col);
            }
            catch { return false; }
        }

        // ---------- Clear all highlights ----------

        public static void ClearAll()
        {
            Log.Clear();
            Log.Write("=== ClearAll START ===");

            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic chart = app.ActiveChart;

                if (chart != null)
                {
                    Log.Write("Active chart — clearing chart only.");
                    int removed = RemoveAllHelpersFromChart(chart);
                    Log.Write($"Removed {removed} helper series from chart.");
                }
                else
                {
                    var answer = System.Windows.Forms.MessageBox.Show(
                        "No chart is currently selected.\n\n" +
                        "Clear ALL highlights on every chart in the active sheet?",
                        "Hadi Tools — Clear All",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question);

                    if (answer != System.Windows.Forms.DialogResult.Yes)
                    {
                        Log.Write("User cancelled.");
                        Log.Write("=== ClearAll END (cancelled) ===");
                        return;
                    }

                    dynamic sheet = app.ActiveSheet;
                    int removed = RemoveAllHelpersFromSheet(sheet);
                    Log.Write($"Removed {removed} helper series from sheet.");
                }
            }
            catch (Exception ex)
            {
                Log.Write("EXCEPTION: " + ex.Message);
            }

            Log.Write("=== ClearAll END ===");
        }

        // ---------- private helpers ----------

        private static int IndexOfMax(double[] v)
        {
            int idx = -1; double best = double.NegativeInfinity;
            for (int i = 0; i < v.Length; i++)
            {
                if (double.IsNaN(v[i]) || double.IsInfinity(v[i])) continue;
                if (v[i] > best) { best = v[i]; idx = i; }
            }
            return idx;
        }

        private static int IndexOfMin(double[] v)
        {
            int idx = -1; double best = double.PositiveInfinity;
            for (int i = 0; i < v.Length; i++)
            {
                if (double.IsNaN(v[i]) || double.IsInfinity(v[i])) continue;
                if (v[i] < best) { best = v[i]; idx = i; }
            }
            return idx;
        }

        private static dynamic TryGetSelectedSeries(dynamic sel)
        {
            if (sel == null) return null;
            object raw = null;
            try
            {
                raw = ((object)sel).GetType().InvokeMember(
                    "Values", System.Reflection.BindingFlags.GetProperty, null, sel, null);
            }
            catch (Exception e) { Log.Write("sel.Values probe failed: " + e.Message); }
            return raw != null ? sel : (dynamic)null;
        }

        private static double[] ReadSeriesValues(dynamic series)
        {
            object raw = series.Values;
            Array arr = raw as Array;
            if (arr == null) return new double[0];
            var list = new List<double>(arr.Length);
            if (arr.Rank == 1)
            {
                int lo = arr.GetLowerBound(0), hi = arr.GetUpperBound(0);
                for (int i = lo; i <= hi; i++) list.Add(ToDouble(arr.GetValue(i)));
            }
            else if (arr.Rank == 2)
            {
                int rLo = arr.GetLowerBound(0), rHi = arr.GetUpperBound(0);
                int cLo = arr.GetLowerBound(1), cHi = arr.GetUpperBound(1);
                if (rHi - rLo >= cHi - cLo)
                    for (int r = rLo; r <= rHi; r++) list.Add(ToDouble(arr.GetValue(r, cLo)));
                else
                    for (int c = cLo; c <= cHi; c++) list.Add(ToDouble(arr.GetValue(rLo, c)));
            }
            return list.ToArray();
        }

        private static double ToDouble(object o)
        {
            if (o == null) return double.NaN;
            if (o is double d) return d;
            if (o is int i) return i;
            if (o is float f) return f;
            if (o is decimal m) return (double)m;
            try { return Convert.ToDouble(o); }
            catch { return double.NaN; }
        }

        private static bool TryParseSeriesFormula(string formula, out string[] sourceCells, out string yRange)
        {
            sourceCells = null; yRange = null;
            try
            {
                if (string.IsNullOrEmpty(formula)) return false;
                if (!formula.StartsWith("=SERIES(", StringComparison.OrdinalIgnoreCase)) return false;
                string inner = formula.Substring(8);
                if (inner.EndsWith(")")) inner = inner.Substring(0, inner.Length - 1);
                var parts = SplitTopLevel(inner, ',');
                if (parts.Count < 3) return false;
                yRange = parts[2].Trim();
                sourceCells = ExpandRange(yRange).ToArray();
                return sourceCells.Length > 0;
            }
            catch (Exception e) { Log.Write("Parse error: " + e.Message); return false; }
        }

        private static List<string> SplitTopLevel(string s, char sep)
        {
            var result = new List<string>();
            int depth = 0; bool inQuote = false; int start = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"') { inQuote = !inQuote; continue; }
                if (inQuote) continue;
                if (c == '(') depth++;
                else if (c == ')') depth--;
                else if (c == sep && depth == 0) { result.Add(s.Substring(start, i - start)); start = i + 1; }
            }
            result.Add(s.Substring(start));
            return result;
        }

        private static List<string> ExpandRange(string rangeRef)
        {
            var cells = new List<string>();
            int bang = rangeRef.LastIndexOf('!');
            if (bang < 0) throw new Exception("No '!' in range: " + rangeRef);
            string sheetRef = rangeRef.Substring(0, bang + 1);
            string rangePart = rangeRef.Substring(bang + 1);
            int colon = rangePart.IndexOf(':');
            string startRef, endRef;
            if (colon < 0) { startRef = rangePart; endRef = rangePart; }
            else { startRef = rangePart.Substring(0, colon); endRef = rangePart.Substring(colon + 1); }
            ParseCellRef(startRef, out int c1, out int r1);
            ParseCellRef(endRef, out int c2, out int r2);
            if (r1 == r2)
            {
                int cMin = Math.Min(c1, c2), cMax = Math.Max(c1, c2);
                for (int c = cMin; c <= cMax; c++) cells.Add(sheetRef + "$" + ColLetter(c) + "$" + r1);
            }
            else if (c1 == c2)
            {
                int rMin = Math.Min(r1, r2), rMax = Math.Max(r1, r2);
                for (int r = rMin; r <= rMax; r++) cells.Add(sheetRef + "$" + ColLetter(c1) + "$" + r);
            }
            else
            {
                int rMin = Math.Min(r1, r2), rMax = Math.Max(r1, r2);
                int cMin = Math.Min(c1, c2), cMax = Math.Max(c1, c2);
                for (int r = rMin; r <= rMax; r++)
                    for (int c = cMin; c <= cMax; c++)
                        cells.Add(sheetRef + "$" + ColLetter(c) + "$" + r);
            }
            return cells;
        }

        private static void ParseCellRef(string s, out int col, out int row)
        {
            s = s.Replace("$", "").ToUpperInvariant();
            int i = 0; while (i < s.Length && char.IsLetter(s[i])) i++;
            string letters = s.Substring(0, i);
            string digits = s.Substring(i);
            col = ColNumber(letters);
            row = int.Parse(digits);
        }

        private static int ColNumber(string letters)
        {
            int n = 0;
            foreach (char c in letters) n = n * 26 + (c - 'A' + 1);
            return n;
        }

        private static string ColLetter(int n)
        {
            string s = "";
            while (n > 0) { int r = (n - 1) % 26; s = (char)('A' + r) + s; n = (n - 1) / 26; }
            return s;
        }

        private static dynamic EnsureHelperSheet(dynamic wb)
        {
            dynamic sheet = null;
            try
            {
                dynamic sheets = wb.Worksheets;
                int count = (int)sheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic s = sheets.Item(i);
                    if ((string)s.Name == HelperSheetName) { sheet = s; break; }
                }
            }
            catch (Exception e) { Log.Write("Scan failed: " + e.Message); }

            if (sheet == null)
            {
                try
                {
                    dynamic sheets = wb.Worksheets;
                    int count = (int)sheets.Count;
                    sheet = sheets.Add(After: sheets.Item(count));
                    sheet.Name = HelperSheetName;
                    Log.Write("Created helper sheet.");
                }
                catch (Exception e) { Log.Write("Create failed: " + e.Message); return null; }
            }
            else Log.Write("Found existing helper sheet.");

            try { sheet.Visible = 2; } catch (Exception e) { Log.Write("Set Visible failed: " + e.Message); }
            return sheet;
        }

        private static dynamic FindHelperSheetNoCreate(dynamic wb)
        {
            try
            {
                dynamic sheets = wb.Worksheets;
                int count = (int)sheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic s = sheets.Item(i);
                    if ((string)s.Name == HelperSheetName) return s;
                }
            }
            catch { }
            return null;
        }

        private static int FindOrCreateColumn(dynamic helper, string key)
        {
            for (int c = 1; c <= MaxHelperColumns; c++)
            {
                string header = null;
                try
                {
                    object v = helper.Cells.Item(1, c).Value2;
                    if (v != null) header = v.ToString();
                }
                catch (Exception e) { Log.Write($"Probe col {c} failed: {e.Message}"); return -1; }

                if (header == key) { Log.Write($"Found existing column {c}."); return c; }
                if (string.IsNullOrEmpty(header))
                {
                    try { helper.Cells.Item(1, c).Value2 = key; Log.Write($"Claimed column {c}."); return c; }
                    catch (Exception e) { Log.Write($"Claim col {c} failed: {e.Message}"); return -1; }
                }
            }
            Log.Write("No free column.");
            return -1;
        }

        private static int FindColumnNoCreate(dynamic helper, string key)
        {
            for (int c = 1; c <= MaxHelperColumns; c++)
            {
                try
                {
                    object v = helper.Cells.Item(1, c).Value2;
                    if (v != null && v.ToString() == key) return c;
                }
                catch { return -1; }
            }
            return -1;
        }

        private static bool WriteFormulaColumn(dynamic helper, int col, string[] sourceCells, string yRange, string kind)
        {
            string func = (kind == "Max") ? "MAX" : "MIN";
            int n = sourceCells.Length;
            try
            {
                dynamic rng = helper.Range(helper.Cells.Item(2, col), helper.Cells.Item(n + 1, col));
                rng.ClearContents();
            }
            catch (Exception e) { Log.Write("Clear failed: " + e.Message); }

            for (int i = 0; i < n; i++)
            {
                int row = i + 2;
                string cell = sourceCells[i];
                string f = $"=IF({cell}={func}({yRange}), {cell}, NA())";
                try { helper.Cells.Item(row, col).Formula = f; }
                catch (Exception e)
                {
                    Log.Write($"Write [{row},{col}] failed: {e.Message}");
                    Log.Write("HRESULT: 0x" + e.HResult.ToString("X8"));
                    return false;
                }
            }
            Log.Write($"Wrote {n} live formulas to column {col}.");
            return true;
        }

        private static bool HelperSeriesAtColumnExists(dynamic chart, int col)
        {
            string colLetter = ColLetter(col);
            string needle = HelperSheetName + "!$" + colLetter + "$";
            try
            {
                dynamic coll = chart.SeriesCollection();
                int count = (int)coll.Count;
                for (int i = 1; i <= count; i++)
                {
                    string f = "";
                    try { f = (string)coll.Item(i).Formula; } catch { }
                    if (f.Contains(needle)) return true;
                }
            }
            catch (Exception e) { Log.Write("Exists check failed: " + e.Message); }
            return false;
        }

        private static void RemoveHelperSeriesAtColumn(dynamic chart, int col)
        {
            string colLetter = ColLetter(col);
            string needle = HelperSheetName + "!$" + colLetter + "$";
            try
            {
                dynamic coll = chart.SeriesCollection();
                int count = (int)coll.Count;
                for (int i = count; i >= 1; i--)
                {
                    string f = "";
                    try { f = (string)coll.Item(i).Formula; } catch { }
                    if (f.Contains(needle))
                    {
                        coll.Item(i).Delete();
                        Log.Write($"Removed helper series at index {i} (col {colLetter}).");
                    }
                }
            }
            catch (Exception e) { Log.Write("Remove failed: " + e.Message); }
        }

        private static int RemoveAllHelpersFromChart(dynamic chart)
        {
            int removed = 0;
            try
            {
                dynamic coll = chart.SeriesCollection();
                int count = (int)coll.Count;
                for (int i = count; i >= 1; i--)
                {
                    string f = "";
                    try { f = (string)coll.Item(i).Formula; } catch { }
                    if (f.Contains(HelperSheetName + "!"))
                    {
                        try { coll.Item(i).Delete(); removed++; }
                        catch (Exception e) { Log.Write($"Delete chart series {i} failed: {e.Message}"); }
                    }
                }
            }
            catch (Exception e) { Log.Write("Chart clear failed: " + e.Message); }
            return removed;
        }

        private static int RemoveAllHelpersFromSheet(dynamic sheet)
        {
            int total = 0;
            try
            {
                dynamic cos = sheet.ChartObjects();
                int count = (int)cos.Count;
                Log.Write($"ChartObjects on sheet: {count}");
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic co = cos.Item(i);
                        dynamic ch = co.Chart;
                        int r = RemoveAllHelpersFromChart(ch);
                        total += r;
                        Log.Write($"Chart {i}: removed {r}.");
                    }
                    catch (Exception e) { Log.Write($"Chart {i} failed: {e.Message}"); }
                }
            }
            catch (Exception e) { Log.Write("Sheet scan failed: " + e.Message); }
            return total;
        }

        private static dynamic AddHelperSeries(dynamic chart, dynamic helper, int col, int n, string key)
        {
            try
            {
                dynamic range = helper.Range(helper.Cells.Item(2, col), helper.Cells.Item(n + 1, col));
                dynamic coll = chart.SeriesCollection();
                dynamic ns = coll.NewSeries();
                ns.Values = range;
                Log.Write("Set Values.");

                try { ns.Name = " "; Log.Write("Set Name = ' '."); }
                catch (Exception e) { Log.Write("Set Name failed: " + e.Message); }

                try
                {
                    if ((bool)chart.HasLegend)
                    {
                        dynamic legendEntries = chart.Legend.LegendEntries();
                        int leCount = (int)legendEntries.Count;
                        if (leCount > 0)
                        {
                            legendEntries.Item(leCount).Delete();
                            Log.Write("Deleted legend entry.");
                        }
                    }
                    else Log.Write("No legend on chart.");
                }
                catch (Exception e) { Log.Write("Legend entry delete failed: " + e.Message); }

                Log.Write("Added helper series.");
                return ns;
            }
            catch (Exception e)
            {
                Log.Write("Add failed: " + e.Message);
                Log.Write("HRESULT: 0x" + e.HResult.ToString("X8"));
                return null;
            }
        }

        private static void FormatHelperSeries(dynamic chart, dynamic ns, string kind)
        {
            try { ns.Format.Line.Visible = msoFalse; Log.Write("Line hidden."); }
            catch (Exception e) { Log.Write("Line.Visible failed: " + e.Message); }

            try { ns.MarkerStyle = xlMarkerStyleCircle; Log.Write("MarkerStyle circle."); }
            catch (Exception e) { Log.Write("MarkerStyle failed: " + e.Message); }

            try { ns.MarkerSize = 10; Log.Write("MarkerSize 10."); }
            catch (Exception e) { Log.Write("MarkerSize failed: " + e.Message); }

            try { ns.MarkerBackgroundColorIndex = xlNone; Log.Write("MarkerBackground xlNone."); }
            catch (Exception e) { Log.Write("MarkerBackground failed: " + e.Message); }

            int color = (kind == "Max") ? ColorRed : ColorBlue;
            try { ns.MarkerForegroundColor = color; Log.Write("MarkerForeground " + color); }
            catch (Exception e) { Log.Write("MarkerForeground failed: " + e.Message); }

            try
            {
                int count = (int)chart.SeriesCollection().Count;
                ns.PlotOrder = count;
                Log.Write("PlotOrder = " + count);
            }
            catch (Exception e) { Log.Write("PlotOrder failed: " + e.Message); }
        }

        private static void AddDataLabelSimple(dynamic ns)
        {
            try { ns.HasDataLabels = true; Log.Write("HasDataLabels = true."); }
            catch (Exception e) { Log.Write("HasDataLabels failed: " + e.Message); return; }

            try { ns.DataLabels.Position = xlLabelPositionAbove; Log.Write("Label position above."); }
            catch (Exception e) { Log.Write("Label position failed: " + e.Message); }

            try { ns.DataLabels.NumberFormat = "General"; Log.Write("Label NumberFormat."); }
            catch (Exception e) { Log.Write("Label NumberFormat failed: " + e.Message); }

            try { ns.DataLabels.Font.Size = 9; Log.Write("Label font size 9."); }
            catch (Exception e) { Log.Write("Label font size failed: " + e.Message); }
        }
    }
}