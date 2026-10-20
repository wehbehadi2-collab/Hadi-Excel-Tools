
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hadi.Excel
{
    public static class NetworkHeaderDetector
    {
        private static readonly Dictionary<string, string> Aliases =
            new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase)
            {
                ["node"] = "Node ID",
                ["node id"] = "Node ID",
                ["junction"] = "Node ID",
                ["junction id"] = "Node ID",
                ["junction label"] = "Node ID",
                ["node label"] = "Node ID",
                ["label"] = "Node ID",

                ["elevation"] = "Elevation",
                ["elevation (m)"] = "Elevation",
                ["elev"] = "Elevation",
                ["ground elevation"] = "Elevation",

                ["pipe"] = "Pipe ID",
                ["pipe id"] = "Pipe ID",
                ["pipe label"] = "Pipe ID",

                ["start node"] = "Start Node",
                ["start node id"] = "Start Node",
                ["from node"] = "Start Node",
                ["node1"] = "Start Node",
                ["node 1"] = "Start Node",

                ["end node"] = "End Node",
                ["end node id"] = "End Node",
                ["stop node"] = "End Node",
                ["to node"] = "End Node",
                ["node2"] = "End Node",
                ["node 2"] = "End Node",

                ["length"] = "Length",
                ["length (m)"] = "Length",
                ["pipe length"] = "Length",

                ["diameter"] = "Diameter",
                ["diameter (mm)"] = "Diameter",
                ["pipe diameter"] = "Diameter",

                ["roughness"] = "Roughness",
                ["roughness (mm)"] = "Roughness",
                ["absolute roughness"] = "Roughness",
                ["absolute roughness (mm)"] = "Roughness",

                ["demand"] = "Demand",
                ["demand (l/s)"] = "Demand",
                ["base demand"] = "Demand",

                ["reservoir"] = "Reservoir ID",
                ["reservoir id"] = "Reservoir ID",
                ["reservoir label"] = "Reservoir ID",

                ["head"] = "Head",
                ["hydraulic head"] = "Head",
                ["total head"] = "Head",
                ["head (m)"] = "Head",

                ["tank"] = "Tank ID",
                ["tank id"] = "Tank ID",

                ["pump"] = "Pump ID",
                ["pump id"] = "Pump ID",

                ["valve"] = "Valve ID",
                ["valve id"] = "Valve ID",

                ["x"] = "X",
                ["x coordinate"] = "X",
                ["x-coordinate"] = "X",

                ["y"] = "Y",
                ["y coordinate"] = "Y",
                ["y-coordinate"] = "Y"
            };

        public static string Analyze(object[,] values)
        {
            if (values == null)
                throw new ArgumentNullException(nameof(values));

            int rowCount = values.GetLength(0);
            int columnCount = values.GetLength(1);

            List<string> findings = new List<string>();

            // Look for headers running horizontally across rows.
            for (int r = 0; r < rowCount; r++)
            {
                List<string> fields = new List<string>();
                List<string> originals = new List<string>();

                for (int c = 0; c < columnCount; c++)
                {
                    string original =
                        Convert.ToString(values[r, c])?.Trim() ?? "";

                    if (TryMapHeader(original, out string field))
                    {
                        fields.Add(field);
                        originals.Add(
                            original + " -> " + field);
                    }
                }

                if (IsCandidate(fields))
                {
                    findings.Add(
                        "Horizontal headers at row " + (r + 1)
                        + ": " + string.Join(", ", originals)
                        + "\nDetected type: "
                        + Classify(fields));
                }
            }

            // Look for headers running vertically down columns.
            for (int c = 0; c < columnCount; c++)
            {
                List<string> fields = new List<string>();
                List<string> originals = new List<string>();

                for (int r = 0; r < rowCount; r++)
                {
                    string original =
                        Convert.ToString(values[r, c])?.Trim() ?? "";

                    if (TryMapHeader(original, out string field))
                    {
                        fields.Add(field);
                        originals.Add(
                            original + " -> " + field);
                    }
                }

                if (IsCandidate(fields))
                {
                    findings.Add(
                        "Vertical headers at column " + (c + 1)
                        + ": " + string.Join(", ", originals)
                        + "\nDetected type: "
                        + Classify(fields));
                }
            }

            if (findings.Count == 0)
            {
                return
                    "No recognizable header groups were found."
                    + "\n\nRecognized names include Node, Elevation, "
                    + "Pipe, Start Node, End Node, Length, Diameter, "
                    + "Roughness, Reservoir, Head, Demand, X and Y."
                    + "\n\nCheck the header wording or extend the "
                    + "alias list in NetworkHeaderDetector.cs.";
            }

            StringBuilder report = new StringBuilder();

            report.AppendLine("Header detection results");
            report.AppendLine();

            for (int i = 0; i < findings.Count; i++)
            {
                report.AppendLine(
                    (i + 1) + ". " + findings[i]);

                report.AppendLine();
            }

            report.AppendLine(
                "Note: this test identifies header locations and "
                + "likely element types. It does not yet extract "
                + "records or generate the INP file.");

            return report.ToString();
        }

        private static bool TryMapHeader(
            string header,
            out string canonicalField)
        {
            string normalized = Normalize(header);

            return Aliases.TryGetValue(
                normalized,
                out canonicalField);
        }

        private static string Normalize(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            return string.Join(
                " ",
                text.Trim()
                    .ToLowerInvariant()
                    .Split(
                        new[] { ' ', '\t', '\r', '\n' },
                        StringSplitOptions.RemoveEmptyEntries));
        }

        private static bool IsCandidate(List<string> fields)
        {
            if (fields.Count < 2)
                return false;

            bool hasIdentifier = fields.Any(field =>
                field == "Node ID"
                || field == "Pipe ID"
                || field == "Reservoir ID"
                || field == "Tank ID"
                || field == "Pump ID"
                || field == "Valve ID"
                || field == "Start Node"
                || field == "End Node");

            return hasIdentifier || fields.Count >= 3;
        }

        private static string Classify(List<string> fields)
        {
            bool pipe =
                fields.Contains("Pipe ID")
                || (fields.Contains("Start Node")
                    && fields.Contains("End Node"));

            bool junction =
                fields.Contains("Node ID")
                && fields.Contains("Elevation");

            bool reservoir =
                fields.Contains("Reservoir ID")
                && fields.Contains("Head");

            bool tank = fields.Contains("Tank ID");
            bool pump = fields.Contains("Pump ID");
            bool valve = fields.Contains("Valve ID");

            List<string> types = new List<string>();

            if (junction)
                types.Add("Junction");

            if (pipe)
                types.Add("Pipe");

            if (reservoir)
                types.Add("Reservoir");

            if (tank)
                types.Add("Tank");

            if (pump)
                types.Add("Pump");

            if (valve)
                types.Add("Valve");

            if (types.Count == 0)
                return "Recognized fields; element type uncertain";

            if (types.Count > 1)
                return "Mixed: " + string.Join(" + ", types);

            return types[0];
        }
    }
}
