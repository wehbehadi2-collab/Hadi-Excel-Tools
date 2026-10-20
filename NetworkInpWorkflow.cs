
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace Hadi.Excel
{
    public static class NetworkInpWorkflow
    {
        public static void ReadNetworkTables()
        {
            try
            {
                dynamic excel = ExcelDnaUtil.Application;

                // Type 8 asks the user to select an Excel range.
                dynamic selectedRange = excel.InputBox(
                    "Select the Excel cells containing your network data, "
                    + "including the headers. You may select one rectangle "
                    + "or multiple areas.",
                    "Select Network Data",
                    Type: 8);

                if (selectedRange == null || selectedRange is bool)
                    return;

                var records = new List<NetworkRecord>();
                var extractionReport = new StringBuilder();

                dynamic areas = selectedRange.Areas;
                int areaCount = (int)areas.Count;

                for (int i = 1; i <= areaCount; i++)
                {
                    dynamic area = areas[i];

                    object[,] values = ToMatrix(area.Value2);

                    extractionReport.AppendLine(
                        $"Selection area {i}");
                    extractionReport.AppendLine(
                        NetworkDataExtractor.Extract(values));
                    extractionReport.AppendLine();

                    records.AddRange(
                        NetworkDataExtractor.ExtractRecords(values));
                }

                if (records.Count == 0)
                {
                    MessageBox.Show(
                        extractionReport
                            + Environment.NewLine
                            + "No network records were recognized. "
                            + "The INP file was not generated.",
                        "Network Data",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                string layoutReport =
                    NetworkAutoLayout.GenerateCoordinates(records);

                string validationReport =
                    NetworkDataValidator.Validate(records);

                if (!validationReport.Contains(
                        "VALIDATION PASSED",
                        StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(
                        extractionReport
                            + Environment.NewLine
                            + layoutReport
                            + Environment.NewLine
                            + validationReport
                            + Environment.NewLine
                            + "Export cancelled because validation failed.",
                        "Network Validation",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // Only valid records are sent to the INP writer.
                NetworkInpExportResult exportResult =
                    NetworkInpWriter.Write(
                        records,
                        new NetworkInpExportOptions
                        {
                            OutputDirectory = System.IO.Path.Combine(
                                Environment.GetFolderPath(
                                    Environment.SpecialFolder.DesktopDirectory),
                                "WaterCAD_Input"),

                            FileName = "WaterCAD_Network.inp",

                            // Darcy-Weisbach absolute roughness in mm.
                            DefaultPipeRoughnessMm = 0.1,

                            // Defaults are explicit and configurable.
                            DefaultMissingDemandToZero = true,
                            DefaultMissingMinorLossToZero = true,
                            DefaultMissingPipeStatusToOpen = true,
                            DefaultMissingPumpSpeedToOne = true,
                            DefaultMissingTankMinimumVolumeToZero = true
                        });

                var finalReport = new StringBuilder();

                finalReport.AppendLine(extractionReport.ToString());
                finalReport.AppendLine(layoutReport);
                finalReport.AppendLine(validationReport);
                finalReport.AppendLine();
                finalReport.AppendLine(exportResult.BuildReport());

                MessageBox.Show(
                    finalReport.ToString(),
                    "Network INP Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "The network operation could not be completed."
                    + Environment.NewLine
                    + Environment.NewLine
                    + ex.Message,
                    "Network INP Workflow",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static object[,] ToMatrix(object value)
        {
            if (value is object[,] source)
            {
                int rows = source.GetLength(0);
                int columns = source.GetLength(1);

                int rowStart = source.GetLowerBound(0);
                int columnStart = source.GetLowerBound(1);

                var result = new object[rows, columns];

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < columns; c++)
                    {
                        result[r, c] =
                            source[rowStart + r, columnStart + c];
                    }
                }

                return result;
            }

            // Excel may return a scalar rather than a two-dimensional
            // array if the user selects only one cell.
            return new[,] { { value } };
        }
    }
}
