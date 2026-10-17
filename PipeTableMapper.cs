
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Hadi.Excel
{
    public static class PipeTableMapper
    {
        public static List<string> Map(
            TableData table,
            out List<MappedPipe> pipes)
        {
            pipes = new List<MappedPipe>();
            List<string> errors = new List<string>();

            if (table == null)
            {
                errors.Add("No table data was provided.");
                return errors;
            }

            int pipeIdCol = FindUniqueColumn(
                table, "Pipe ID", errors);

            int startNodeCol = FindUniqueColumn(
                table, "Start Node", errors);

            int endNodeCol = FindUniqueColumn(
                table, "End Node", errors);

            int lengthCol = FindUniqueColumn(
                table, "Length", errors);

            int diameterCol = FindUniqueColumn(
                table, "Diameter", errors);

            if (errors.Count > 0)
                return errors;

            for (int r = 0; r < table.RowCount; r++)
            {
                string pipeId = GetText(
                    table.GetValue(r, pipeIdCol));

                string startNode = GetText(
                    table.GetValue(r, startNodeCol));

                string endNode = GetText(
                    table.GetValue(r, endNodeCol));

                double length;
                double diameter;

                if (!TryGetNumber(
                    table.GetValue(r, lengthCol), out length)
                    || length <= 0)
                {
                    errors.Add(
                        "Data row " + (r + 1)
                        + ": Length must be a positive number.");
                    continue;
                }

                if (!TryGetNumber(
                    table.GetValue(r, diameterCol), out diameter)
                    || diameter <= 0)
                {
                    errors.Add(
                        "Data row " + (r + 1)
                        + ": Diameter must be a positive number.");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(pipeId)
                    || string.IsNullOrWhiteSpace(startNode)
                    || string.IsNullOrWhiteSpace(endNode))
                {
                    errors.Add(
                        "Data row " + (r + 1)
                        + ": Pipe ID, Start Node, and End Node "
                        + "must not be empty.");
                    continue;
                }

                pipes.Add(new MappedPipe(
                    pipeId,
                    startNode,
                    endNode,
                    length,
                    diameter));
            }

            if (errors.Count > 0)
                pipes.Clear();

            return errors;
        }

        private static int FindUniqueColumn(
            TableData table,
            string requiredHeader,
            List<string> errors)
        {
            int foundIndex = -1;
            int matchCount = 0;

            for (int c = 0; c < table.Headers.Count; c++)
            {
                string header = table.Headers[c] ?? "";

                if (string.Equals(
                    header.Trim(),
                    requiredHeader,
                    StringComparison.OrdinalIgnoreCase))
                {
                    foundIndex = c;
                    matchCount++;
                }
            }

            if (matchCount == 0)
            {
                errors.Add(
                    "Required column not found: '"
                    + requiredHeader + "'.");
                return -1;
            }

            if (matchCount > 1)
            {
                errors.Add(
                    "More than one column matches '"
                    + requiredHeader + "'. Keep only one.");
                return -1;
            }

            return foundIndex;
        }

        private static string GetText(object value)
        {
            if (value == null)
                return "";

            return Convert.ToString(
                value, CultureInfo.CurrentCulture).Trim();
        }

        private static bool TryGetNumber(
            object value,
            out double number)
        {
            number = 0;

            if (value == null)
                return false;

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
