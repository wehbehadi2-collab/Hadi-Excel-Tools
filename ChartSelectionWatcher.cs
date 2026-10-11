
using System;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace Hadi.Excel
{
    public static class ChartSelectionWatcher
    {
        private static Timer _timer;
        private static string _lastSelection = "";

        public static void Start()
        {
            if (_timer != null)
                return;

            _timer = new Timer();
            _timer.Interval = 300;
            _timer.Tick += CheckSelection;
            _timer.Start();
        }

        public static void Stop()
        {
            if (_timer == null)
                return;

            _timer.Stop();
            _timer.Tick -= CheckSelection;
            _timer.Dispose();
            _timer = null;
        }

        private static void CheckSelection(object sender, EventArgs e)
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic chart = app.ActiveChart;
                dynamic selection = app.Selection;

                string currentSelection = "NoChart";

                if (chart != null && selection != null)
                {
                    string workbookName = "";
                    string sheetName = "";
                    string chartName = "";
                    string seriesName = "";
                    string seriesFormula = "";

                    try { workbookName = app.ActiveWorkbook.Name; }
                    catch { }

                    try { sheetName = app.ActiveSheet.Name; }
                    catch { }

                    try { chartName = chart.Name; }
                    catch { }

                    try { seriesName = selection.Name; }
                    catch { }

                    try { seriesFormula = selection.Formula; }
                    catch { }

                    currentSelection =
                        workbookName + "|" +
                        sheetName + "|" +
                        chartName + "|" +
                        seriesName + "|" +
                        seriesFormula;
                }

                if (currentSelection != _lastSelection)
                {
                    _lastSelection = currentSelection;
                    HadiRibbon.InvalidateControls();
                }
            }
            catch
            {
                // Ignore temporary Excel selection errors.
            }
        }
    }
}
