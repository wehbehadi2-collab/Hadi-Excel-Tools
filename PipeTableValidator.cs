
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Hadi.Excel
{
    public static class PipeTableValidator
    {
        public static List<string> Validate(TableData table)
        {
            List<string> errors = new List<string>();

            if (table == null)
            {
                errors.Add("No table data was provided.");
                return errors;
            }

            // Find the required columns.
            int pipeIdCol = table.FindColumn("Pipe ID");
            int startNodeCol = table.FindColumn("Start Node");
            int endNodeCol = table.FindColumn("End Node");
            int lengthCol = table.FindColumn("Length");
            int diameterCol = table.FindColumn("Diameter");

            CheckColumn(pipeIdCol, "Pipe ID", errors);
            CheckColumn(startNodeCol, "Start Node", errors);
            CheckColumn(endNodeCol, "End Node", errors);
            CheckColumn(lengthCol, "Length", errors);
            CheckColumn(diameterCol, "Diameter", errors);

            // Do not inspect rows if required columns are missing.
            if (errors.Count > 0)
                return errors;

            HashSet<string> pipeIds =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int r = 0; r < table.RowCount; r++)
            {
                // Excel row number: selected range starts with header.
                int excelRow = r + 2;

                string pipeId = GetText(table, r, pipeIdCol);
                string startNode = GetText(table, r, startNodeCol);
                string endNode = GetText(table, r, endNodeCol);

                // Required text values.
                if (string.IsNullOrWhiteSpace(pipeId))
                {
                    errors.Add(
                        "Row " + excelRow + ": Pipe ID is missing.");
                }
                else if (!pipeIds.Add(pipeId))
                {
                    errors.Add(
                        "Row " + excelRow
                        + ": Duplicate Pipe ID '" + pipeId + "'.");
                }

                if (string.IsNullOrWhiteSpace(startNode))
                {
                    errors.Add(
                        "Row " + excelRow + ": Start Node is missing.");
                }

                if (string.IsNullOrWhiteSpace(endNode))
                {
                    errors.Add(
                        "Row " + excelRow + ": End Node is missing.");
                }

                if (!string.IsNullOrWhiteSpace(startNode)
                    && !string.IsNullOrWhiteSpace(endNode)
                    && string.Equals(
                        startNode.Trim(),
                        endNode.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add(
                        "Row " + excelRow
                        + ": Start Node and End Node are identical ('"
                        + startNode + "').");
                }

                // Length must be a positive number.
                CheckPositiveNumber(
                    table.GetValue(r, lengthCol),
                    "Length",
                    excelRow,
                    errors);

                // Diameter must be a positive number.
                CheckPositiveNumber(
                    table.GetValue(r, diameterCol),
                    "Diameter",
                    excelRow,
                    errors);
            }

            return errors;
        }

        private static void CheckColumn(
            int columnIndex,
            string columnName,
            List<string> errors)
        {
            if (columnIndex < 0)
            {
                errors.Add(
                    "Required column not found: '" + columnName + "'.");
            }
        }

        private static string GetText(
            TableData table, int row, int column)
        {
            object value = table.GetValue(row, column);

            if (value == null)
                return "";

            return Convert.ToString(
                value, CultureInfo.CurrentCulture).Trim();
        }

        private static void CheckPositiveNumber(
            object value,
            string columnName,
            int excelRow,
            List<string> errors)
        {
            double number;

            if (!TryGetNumber(value, out number))
            {
                errors.Add(
                    "Row " + excelRow + ": " + columnName
                    + " must be a valid number.");
                return;
            }

            if (double.IsNaN(number) || double.IsInfinity(number)
                || number <= 0)
            {
                errors.Add(
                    "Row " + excelRow + ": " + columnName
                    + " must be greater than zero.");
            }
        }

        private static bool TryGetNumber(
            object value, out double number)
        {
            number = 0;

            if (value == null)
                return false;

            // Excel numeric cells commonly arrive as double values.
            if (value is double || value is float
                || value is decimal || value is int
                || value is long || value is short)
            {
                try
                {
                    number = Convert.ToDouble(
                        value, CultureInfo.InvariantCulture);

                    return !double.IsNaN(number)
                        && !double.IsInfinity(number);
                }
                catch
                {
                    return false;
                }
            }

            string text = Convert.ToString(
                value, CultureInfo.CurrentCulture);

            if (string.IsNullOrWhiteSpace(text))
                return false;

            return double.TryParse(
                text,
                NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.CurrentCulture,
                out number)
                && !double.IsNaN(number)
                && !double.IsInfinity(number);
        }
    }
}
