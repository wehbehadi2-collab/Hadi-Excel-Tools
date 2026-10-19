
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Hadi.Excel
{
    public static class JunctionCSVExporter
    {
        private static readonly string OutputFolder =
            @"C:\Users\Hadi Wehbi\Desktop\WaterCAD_Input";

        private static readonly string OutputFile =
            Path.Combine(OutputFolder, "Junctions.csv");

        public static string Export(
            List<MappedJunction> junctions)
        {
            if (junctions == null || junctions.Count == 0)
            {
                throw new ArgumentException(
                    "There are no junctions to export.",
                    nameof(junctions));
            }

            Directory.CreateDirectory(OutputFolder);

            var csv = new StringBuilder();

            csv.AppendLine("Label,Elevation,X,Y");

            foreach (MappedJunction junction in junctions)
            {
                csv.Append(EscapeCsv(junction.Label));
                csv.Append(',');
                csv.Append(junction.Elevation.ToString(
                    "G", CultureInfo.InvariantCulture));
                csv.Append(',');
                csv.Append(junction.X.ToString(
                    "G", CultureInfo.InvariantCulture));
                csv.Append(',');
                csv.AppendLine(junction.Y.ToString(
                    "G", CultureInfo.InvariantCulture));
            }

            File.WriteAllText(
                OutputFile,
                csv.ToString(),
                new UTF8Encoding(true));

            return OutputFile;
        }

        private static string EscapeCsv(string value)
        {
            if (value == null)
                return "";

            if (value.Contains(",") ||
                value.Contains("\"") ||
                value.Contains("\r") ||
                value.Contains("\n"))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }

            return value;
        }
    }
}
