
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

                // Ask the user to select the Excel network data,
                // including its headers.
                dynamic selectedRange = excel.InputBox(
                    "Select the Excel cells containing your network data, "
                    + "including the headers. You may select one rectangle "
                    + "or multiple areas.",
                    "Select Network Data",
                    Type: 8);

                if (selectedRange == null || selectedRange is bool)
                    return;

                var records = new List<NetworkRecord>();

                dynamic areas = selectedRange.Areas;
                int areaCount = (int)areas.Count;

                for (int i = 1; i <= areaCount; i++)
                {
                    dynamic area = areas[i];

                    object[,] values = ToMatrix(area.Value2);

                    records.AddRange(
                        NetworkDataExtractor.ExtractRecords(values));
                }

                // Stop if no supported network records were found.
                if (records.Count == 0)
                {
                    MessageBox.Show(
                        "No network records were recognized.\n\n"
                        + "Check that your selection contains supported "
                        + "network data and the correct column headers.\n\n"
                        + "The INP file was not generated.",
                        "No Network Data",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // Validate before generating coordinates or exporting.
                string validationReport =
                    NetworkDataValidator.Validate(records);

                if (!validationReport.Contains(
                        "VALIDATION PASSED",
                        StringComparison.OrdinalIgnoreCase))
                {
                    string errorMessage =
                        BuildValidationErrorMessage(validationReport);

                    MessageBox.Show(
                        errorMessage,
                        "Network Validation Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // Only valid records proceed to coordinate generation.
                NetworkAutoLayout.GenerateCoordinates(records);

                // Preserve the existing automatic export behavior.
                // The same file is overwritten on each successful export.
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

                            // Existing export defaults.
                            DefaultMissingDemandToZero = true,
                            DefaultMissingMinorLossToZero = true,
                            DefaultMissingPipeStatusToOpen = true,
                            DefaultMissingPumpSpeedToOne = true,
                            DefaultMissingTankMinimumVolumeToZero = true
                        });

                // Intentionally no success popup.
                // The writer creates/replaces the INP file automatically.
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

        private static string BuildValidationErrorMessage(
            string validationReport)
        {
            var message = new StringBuilder();

            message.AppendLine(
                "The network contains errors that must be fixed "
                + "before it can be exported.");

            message.AppendLine();
            message.AppendLine("Validation errors:");

            string[] lines = validationReport.Split(
                new[] { "\r\n", "\n" },
                StringSplitOptions.None);

            int errorCount = 0;

            foreach (string line in lines)
            {
                string trimmedLine = line.Trim();

                if (trimmedLine.StartsWith(
                        "- ",
                        StringComparison.Ordinal))
                {
                    message.AppendLine(trimmedLine);
                    errorCount++;
                }
            }

            if (errorCount == 0)
            {
                // Fallback in case the validator's report format changes.
                message.AppendLine();
                message.AppendLine(
                    "Validation failed, but no individual error "
                    + "details could be extracted.");

                message.AppendLine();
                message.AppendLine(
                    "Please check the selected network data.");
            }

            message.AppendLine();
            message.AppendLine(
                $"Total errors: {errorCount}");

            message.AppendLine();
            message.AppendLine(
                "Correct these errors in Excel and run the tool again.");

            message.AppendLine(
                "The INP file was not generated.");

            return message.ToString();
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

            // Excel may return a scalar for a single selected cell.
            return new[,] { { value } };
        }
    }
}
