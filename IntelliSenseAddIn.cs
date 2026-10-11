using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace Hadi.Excel
{
    public class IntelliSenseAddIn : IExcelAddIn
    {
        public void AutoOpen() => IntelliSenseServer.Install();
        public void AutoClose() => IntelliSenseServer.Uninstall();
    }
}