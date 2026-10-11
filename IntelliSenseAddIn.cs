
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace Hadi.Excel
{
    public class IntelliSenseAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
            ChartSelectionWatcher.Start();
        }

        public void AutoClose()
        {
            ChartSelectionWatcher.Stop();
            IntelliSenseServer.Uninstall();
        }
    }
}
