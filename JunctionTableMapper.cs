
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Hadi.Excel
{
    public static class JunctionTableMapper
    {
        public static bool Map(
            TableData table,
            out List<MappedJunction> junctions,
            out List<string> errors)
        {
            junctions = new List<MappedJunction>();
            errors = new List<string>();

            if (table == null)
            {
                errors.Add("No junction table was supplied.");
                return false;
            }

            int labelColumn = table.FindColumn("Label");
            int elevationColumn = table.FindColumn("Elevation");
            int xColumn = table.FindColumn("X");
            int yColumn = table.FindColumn("Y");

            if (labelColumn < 0)
                errors.Add("Missing required column: Label.");

            if (elevationColumn < 0)
                errors.Add("Missing required column: Elevation.");

            bool hasX = xColumn >= 0;
            bool hasY = yColumn >= 0;

            if (hasX != hasY)
            {
                errors.Add(
                    "X and Y must either both exist or both be absent.");
            }

            if (errors.Count > 0)
                return false;

            bool generateCoordinates = !hasX;

            var usedLabels = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];

                string label = Convert.ToString(
                    row[labelColumn])?.Trim() ?? "";

                if (string.IsNullOrWhiteSpace(label))
                {
                    errors.Add(
                        $"Row {i + 1}: Junction Label is empty.");
                    continue;
                }

                if (!usedLabels.Add(label))
                {
                    errors.Add(
                        $"Row {i + 1}: Duplicate junction Label '{label}'.");
                    continue;
                }

                if (!TryGetNumber(
                    row[elevationColumn], out double elevation))
                {
                    errors.Add(
                        $"Junction '{label}': Elevation is not a valid number.");
                    continue;
                }

                double x;
                double y;

                if (generateCoordinates)
                {
                    // Grid layout: 5 junctions per row.
                    x = (i % 5) * 100.0;
                    y = (i / 5) * 100.0;
                }
                else
                {
                    if (!TryGetNumber(row[xColumn], out x) ||
                        !TryGetNumber(row[yColumn], out y))
                    {
                        errors.Add(
                            $"Junction '{label}': X and Y must be valid numbers.");
                        continue;
                    }
                }

                junctions.Add(new MappedJunction(
                    label,
                    elevation,
                    x,
                    y,
                    generateCoordinates));
            }

            if (errors.Count > 0)
            {
                junctions.Clear();
                return false;
            }

            if (junctions.Count == 0)
            {
                errors.Add("The table contains no junction records.");
                return false;
            }

            return true;
        }

        private static bool TryGetNumber(
            object value,
            out double result)
        {
            string text = Convert.ToString(
                value, CultureInfo.InvariantCulture)?.Trim() ?? "";

            return double.TryParse(
                text,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out result)
                && !double.IsNaN(result)
                && !double.IsInfinity(result);
        }
    }
}
