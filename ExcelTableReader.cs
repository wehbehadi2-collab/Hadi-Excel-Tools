
using System;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace Hadi.Excel
{
    public static class ExcelTableReader
    {
        public static void ReadAndPreview()
        {
            try
            {
                TableData table = ReadSelectedTable();

                if (table == null)
                    return;

                // Step 1: Validate the selected pipe table.
                List<string> errors =
                    PipeTableValidator.Validate(table);

                if (errors.Count > 0)
                {
                    ShowErrors(
                        errors,
                        "Pipe Table Validation",
                        "The selected table contains");

                    ShowPreview(table);
                    return;
                }

                // Step 2: Map the validated table into pipe records.
                List<MappedPipe> pipes;

                List<string> mappingErrors =
                    PipeTableMapper.Map(table, out pipes);

                if (mappingErrors.Count > 0)
                {
                    ShowErrors(
                        mappingErrors,
                        "Pipe Column Mapping",
                        "The table could not be mapped");

                    ShowPreview(table);
                    return;
                }

                // Step 3: Confirm validation and mapping succeeded.
                MessageBox.Show(
                    "Pipe table validation successful.\n\n"
                    + "Data rows checked: " + table.RowCount
                    + "\nValidation errors: 0"
                    + "\nMapped pipes: " + pipes.Count,
                    "Pipe Table Validator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // Keep the original Excel table preview.
                ShowPreview(table);

                // Step 4: Show the new mapped-pipe preview.
                ShowMappedPipePreview(pipes);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Could not read, validate, or map "
                    + "the selected Excel range."
                    + "\n\n" + ex.Message,
                    "Read Excel Table",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // Reads the selected Excel range and returns reusable data.
        // Returns null if the user cancels the selection.
        public static void ReadJunctionsAndPreview()
        {
            try
            {
                TableData table = ReadSelectedTable();

                if (table == null)
                    return;

                List<MappedJunction> junctions;
                List<string> errors;

                bool success = JunctionTableMapper.Map(
                    table,
                    out junctions,
                    out errors);

                if (!success)
                {
                    ShowErrors(
                        errors,
                        "Junction Table Mapping",
                        "The selected table could not be mapped");

                    ShowPreview(table);
                    return;
                }
                string outputPath = JunctionCSVExporter.Export(junctions);

                MessageBox.Show(
                      "Junction mapping and CSV export successful.\n\n"
                      + "Mapped junctions: " + junctions.Count
                      + "\nCoordinates generated automatically: "
                      + junctions.FindAll(
                       j => j.CoordinatesGenerated).Count
                      + "\n\nCSV saved to:\n" + outputPath,
                      "Junction Table Mapper",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                ShowPreview(table);
                ShowMappedJunctionPreview(junctions);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Could not read or map the junction table."
                    + "\n\n" + ex.Message,
                    "Junction Table Mapper",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public static TableData ReadSelectedTable()
        {
            dynamic excel = ExcelDnaUtil.Application;

            dynamic selectedRange = excel.InputBox(
                "Select the table cells, including the header row.",
                "Read Excel Table",
                Type: 8);

            if (selectedRange == null || selectedRange is bool)
                return null;

            int areaCount =
                Convert.ToInt32(selectedRange.Areas.Count);

            if (areaCount != 1)
            {
                throw new InvalidOperationException(
                    "Please select one continuous rectangular range.");
            }

            int rowCount =
                Convert.ToInt32(selectedRange.Rows.Count);

            int columnCount =
                Convert.ToInt32(selectedRange.Columns.Count);

            if (rowCount < 1 || columnCount < 1)
            {
                throw new InvalidOperationException(
                    "The selected range is empty.");
            }

            string sheetName =
                Convert.ToString(selectedRange.Worksheet.Name);

            string address =
                Convert.ToString(selectedRange.Address);

            object rawValue = selectedRange.Value2;

            object[,] values = ToMatrix(
                rawValue, rowCount, columnCount);

            // The first selected row contains the column headers.
            List<string> headers = new List<string>();

            for (int c = 0; c < columnCount; c++)
            {
                string header = FormatValue(values[0, c]);

                if (header == "(blank)")
                    header = "Column " + (c + 1);

                headers.Add(header);
            }

            // The remaining rows contain the data.
            List<List<object>> rows =
                new List<List<object>>();

            for (int r = 1; r < rowCount; r++)
            {
                List<object> dataRow = new List<object>();

                for (int c = 0; c < columnCount; c++)
                {
                    dataRow.Add(values[r, c]);
                }

                rows.Add(dataRow);
            }

            return new TableData(
                sheetName,
                address,
                headers,
                rows);
        }

        private static void ShowErrors(
            List<string> errors,
            string title,
            string introduction)
        {
            StringBuilder report = new StringBuilder();

            report.AppendLine(
                introduction + " "
                + errors.Count + " error(s):");

            report.AppendLine();

            int maxErrorsToShow = Math.Min(errors.Count, 25);

            for (int i = 0; i < maxErrorsToShow; i++)
            {
                report.AppendLine(
                    (i + 1) + ". " + errors[i]);
            }

            if (errors.Count > maxErrorsToShow)
            {
                report.AppendLine();
                report.AppendLine(
                    "... and "
                    + (errors.Count - maxErrorsToShow)
                    + " more errors.");
            }

            MessageBox.Show(
                report.ToString(),
                title,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private static void ShowPreview(TableData table)
        {
            StringBuilder report = new StringBuilder();

            report.AppendLine("Excel table read successfully.");
            report.AppendLine();
            report.AppendLine(
                "Worksheet: " + table.WorksheetName);
            report.AppendLine(
                "Range: " + table.RangeAddress);
            report.AppendLine(
                "Data rows: " + table.RowCount);
            report.AppendLine(
                "Columns: " + table.ColumnCount);
            report.AppendLine();
            report.AppendLine("HEADERS");

            int columnsToShow =
                Math.Min(table.ColumnCount, 8);

            for (int c = 0; c < columnsToShow; c++)
            {
                report.AppendLine(
                    (c + 1) + ". " + table.Headers[c]);
            }

            if (table.ColumnCount > columnsToShow)
            {
                report.AppendLine(
                    "... and "
                    + (table.ColumnCount - columnsToShow)
                    + " more columns.");
            }

            report.AppendLine();
            report.AppendLine("SAMPLE DATA");

            int rowsToShow = Math.Min(table.RowCount, 6);

            if (rowsToShow == 0)
            {
                report.AppendLine(
                    "No data rows below the header.");
            }

            for (int r = 0; r < rowsToShow; r++)
            {
                report.AppendLine();
                report.AppendLine("Row " + (r + 1) + ":");

                for (int c = 0; c < columnsToShow; c++)
                {
                    report.AppendLine(
                        "  " + table.Headers[c] + ": "
                        + FormatValue(table.Rows[r][c]));
                }
            }

            if (table.RowCount > rowsToShow)
            {
                report.AppendLine();
                report.AppendLine(
                    "... and "
                    + (table.RowCount - rowsToShow)
                    + " more data rows.");
            }

            MessageBox.Show(
                report.ToString(),
                "Excel Table Preview",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static void ShowMappedPipePreview(
            List<MappedPipe> pipes)
        {
            StringBuilder report = new StringBuilder();

            report.AppendLine(
                "Pipe column mapping successful.");
            report.AppendLine();
            report.AppendLine(
                "Mapped pipe records: " + pipes.Count);
            report.AppendLine();
            report.AppendLine(
                "Each record now has consistent fields:");
            report.AppendLine(
                "Pipe ID, Start Node, End Node, Length, Diameter");

            int pipesToShow = Math.Min(pipes.Count, 10);

            for (int i = 0; i < pipesToShow; i++)
            {
                MappedPipe pipe = pipes[i];

                report.AppendLine();
                report.AppendLine(
                    "Pipe " + (i + 1) + ": " + pipe.PipeId);
                report.AppendLine(
                    "  Start Node: " + pipe.StartNode);
                report.AppendLine(
                    "  End Node: " + pipe.EndNode);
                report.AppendLine(
                    "  Length: " + pipe.Length.ToString(
                        "G", CultureInfo.CurrentCulture));
                report.AppendLine(
                    "  Diameter: " + pipe.Diameter.ToString(
                        "G", CultureInfo.CurrentCulture));
            }

            if (pipes.Count > pipesToShow)
            {
                report.AppendLine();
                report.AppendLine(
                    "... and " + (pipes.Count - pipesToShow)
                    + " more mapped pipes.");
            }

            MessageBox.Show(
                report.ToString(),
                "Mapped Pipe Preview",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }


        private static void ShowMappedJunctionPreview(
            List<MappedJunction> junctions)
        {
            StringBuilder report = new StringBuilder();

            report.AppendLine("Junction mapping successful.");
            report.AppendLine();
            report.AppendLine(
                "Mapped junctions: " + junctions.Count);
            report.AppendLine();

            int count = Math.Min(junctions.Count, 10);

            for (int i = 0; i < count; i++)
            {
                MappedJunction junction = junctions[i];

                report.AppendLine(
                    "Junction " + (i + 1) + ": "
                    + junction.Label);

                report.AppendLine(
                    "  Elevation: "
                    + junction.Elevation.ToString(
                        "G", CultureInfo.CurrentCulture));

                report.AppendLine(
                    "  X: "
                    + junction.X.ToString(
                        "G", CultureInfo.CurrentCulture));

                report.AppendLine(
                    "  Y: "
                    + junction.Y.ToString(
                        "G", CultureInfo.CurrentCulture));

                report.AppendLine(
                    "  Coordinates generated: "
                    + junction.CoordinatesGenerated);

                report.AppendLine();
            }

            if (junctions.Count > count)
            {
                report.AppendLine(
                    "... and " + (junctions.Count - count)
                    + " more junctions.");
            }

            MessageBox.Show(
                report.ToString(),
                "Mapped Junction Preview",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static object[,] ToMatrix(
            object value, int rows, int columns)
        {
            object[,] result = new object[rows, columns];

            if (rows == 1 && columns == 1)
            {
                result[0, 0] = value;
                return result;
            }

            Array source = value as Array;

            if (source == null || source.Rank != 2)
            {
                throw new InvalidOperationException(
                    "Excel did not return a two-dimensional table.");
            }

            int rowStart = source.GetLowerBound(0);
            int columnStart = source.GetLowerBound(1);

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++)
                {
                    result[r, c] = source.GetValue(
                        rowStart + r,
                        columnStart + c);
                }
            }

            return result;
        }

        private static string FormatValue(object value)
        {
            if (value == null)
                return "(blank)";

            string text = Convert.ToString(
                value, CultureInfo.CurrentCulture);

            if (string.IsNullOrWhiteSpace(text))
                return "(blank)";

            if (text.Length > 60)
                text = text.Substring(0, 57) + "...";

            return text;
        }
    }
}
