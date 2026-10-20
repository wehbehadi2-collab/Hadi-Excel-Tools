using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Hadi.Excel
{
    public static class NetworkDataExtractor
    {
        private sealed class HeaderLayout
        {
            public bool Horizontal { get; set; }
            public int HeaderIndex { get; set; }

            public Dictionary<int, string> Fields { get; set; } =
                new Dictionary<int, string>();

            public int Score =>
                Fields.Values.Distinct(
                    StringComparer.OrdinalIgnoreCase).Count();
        }

        private static readonly string[] ElementTypes =
        {
            "Junction", "Reservoir", "Pipe", "Tank", "Pump", "Valve"
        };

        public static string Extract(object[,] values)
        {
            List<NetworkRecord> records = ExtractRecords(values);

            var report = new StringBuilder();
            report.AppendLine("Network data extraction preview");
            report.AppendLine();

            if (records.Count == 0)
            {
                report.AppendLine("No network element records were recognized.");
                report.AppendLine("Check the headers and element data.");
                return report.ToString();
            }

            foreach (NetworkRecord record in records)
            {
                report.AppendLine(
                    record.Type + ": " + DescribeRecord(record));
            }

            report.AppendLine();
            report.AppendLine("Extraction totals:");

            foreach (string type in ElementTypes)
            {
                int count = records.Count(r => r.Type == type);

                if (count > 0)
                    report.AppendLine($"{type}: {count}");
            }

            return report.ToString();
        }

        public static List<NetworkRecord> ExtractRecords(object[,] values)
        {
            if (values == null)
                throw new ArgumentNullException(nameof(values));

            var records = new List<NetworkRecord>();

            int rows = values.GetLength(0);
            int columns = values.GetLength(1);

            if (rows == 0 || columns == 0)
                return records;

            HeaderLayout layout = FindBestLayout(values);

            if (layout == null || layout.Score < 2)
                return records;

            var availableFields = new HashSet<string>(
                layout.Fields.Values,
                StringComparer.OrdinalIgnoreCase);

            if (layout.Horizontal)
            {
                for (int r = layout.HeaderIndex + 1; r < rows; r++)
                {
                    var fields = new Dictionary<string, string>(
                        StringComparer.OrdinalIgnoreCase);

                    foreach (var header in layout.Fields)
                    {
                        string value = CleanValue(values[r, header.Key]);

                        if (!string.IsNullOrWhiteSpace(value))
                            fields[header.Value] = value;
                    }

                    AddRecords(fields, "Excel row " + (r + 1),
                        availableFields, records);
                }
            }
            else
            {
                for (int c = layout.HeaderIndex + 1; c < columns; c++)
                {
                    var fields = new Dictionary<string, string>(
                        StringComparer.OrdinalIgnoreCase);

                    foreach (var header in layout.Fields)
                    {
                        string value = CleanValue(values[header.Key, c]);

                        if (!string.IsNullOrWhiteSpace(value))
                            fields[header.Value] = value;
                    }

                    AddRecords(fields, "Excel column " + (c + 1),
                        availableFields, records);
                }
            }

            return records;
        }

        private static HeaderLayout FindBestLayout(object[,] values)
        {
            int rows = values.GetLength(0);
            int columns = values.GetLength(1);

            HeaderLayout bestHorizontal = null;
            HeaderLayout bestVertical = null;

            for (int r = 0; r < rows; r++)
            {
                var found = new Dictionary<int, string>();

                for (int c = 0; c < columns; c++)
                {
                    if (NetworkParameterCatalog.TryGetCanonicalName(
                            CleanValue(values[r, c]),
                            out string canonical))
                    {
                        found[c] = canonical;
                    }
                }

                var candidate = new HeaderLayout
                {
                    Horizontal = true,
                    HeaderIndex = r,
                    Fields = found
                };

                if (candidate.Score > (bestHorizontal?.Score ?? 0))
                    bestHorizontal = candidate;
            }

            for (int c = 0; c < columns; c++)
            {
                var found = new Dictionary<int, string>();

                for (int r = 0; r < rows; r++)
                {
                    if (NetworkParameterCatalog.TryGetCanonicalName(
                            CleanValue(values[r, c]),
                            out string canonical))
                    {
                        found[r] = canonical;
                    }
                }

                var candidate = new HeaderLayout
                {
                    Horizontal = false,
                    HeaderIndex = c,
                    Fields = found
                };

                if (candidate.Score > (bestVertical?.Score ?? 0))
                    bestVertical = candidate;
            }

            int horizontalScore = bestHorizontal?.Score ?? 0;
            int verticalScore = bestVertical?.Score ?? 0;

            return horizontalScore >= verticalScore
                ? bestHorizontal
                : bestVertical;
        }

        private static void AddRecords(
            Dictionary<string, string> fields,
            string source,
            HashSet<string> availableFields,
            List<NetworkRecord> records)
        {
            if (fields.Count == 0)
                return;

            // Junction-specific evidence. Elevation alone is not enough
            // because tanks also have an elevation.
            if (Has(fields, "Node ID")
                || (availableFields.Contains("Node ID")
                    && HasAny(fields, "Demand", "Emitter Coefficient")))
            {
                AddRecord("Junction", fields, source,
                    availableFields, records);
            }

            if (Has(fields, "Reservoir ID")
                || (availableFields.Contains("Reservoir ID")
                    && Has(fields, "Head")))
            {
                AddRecord("Reservoir", fields, source,
                    availableFields, records);
            }

            bool pumpEvidence = Has(fields, "Pump ID")
                || HasAny(fields, "Pump Curve", "Power", "Speed",
                    "Speed Pattern");

            bool valveEvidence = Has(fields, "Valve ID")
                || HasAny(fields, "Valve Type", "Setting");

            // Pipe-specific evidence excludes Diameter and Minor Loss
            // alone because those fields can also belong to tanks/valves.
            bool pipeEvidence = Has(fields, "Pipe ID")
                || HasAny(fields, "Length", "Roughness",
                    "Bulk Reaction", "Wall Reaction")
                || (HasAny(fields, "Start Node", "End Node")
                    && !pumpEvidence
                    && !valveEvidence);

            if (pipeEvidence)
            {
                AddRecord("Pipe", fields, source,
                    availableFields, records);
            }

            // A tank can be identified by its ID or tank-specific levels.
            if (Has(fields, "Tank ID")
                || HasAny(fields, "Initial Level", "Minimum Level",
                    "Maximum Level", "Minimum Volume", "Volume Curve"))
            {
                AddRecord("Tank", fields, source,
                    availableFields, records);
            }

            if (pumpEvidence)
            {
                AddRecord("Pump", fields, source,
                    availableFields, records);
            }

            if (valveEvidence)
            {
                AddRecord("Valve", fields, source,
                    availableFields, records);
            }
        }

        private static void AddRecord(
            string type,
            Dictionary<string, string> fields,
            string source,
            HashSet<string> availableFields,
            List<NetworkRecord> records)
        {
            var record = new NetworkRecord
            {
                Type = type,
                Source = source
            };

            foreach (string field in NetworkParameterCatalog.FieldsFor(type))
            {
                if (fields.TryGetValue(field, out string value))
                    record.Fields[field] = value;
            }

            foreach (string field in availableFields)
                record.AvailableFields.Add(field);

            records.Add(record);
        }

        private static bool Has(
            Dictionary<string, string> fields,
            string name)
        {
            return fields.TryGetValue(name, out string value)
                && !string.IsNullOrWhiteSpace(value);
        }

        private static bool HasAny(
            Dictionary<string, string> fields,
            params string[] names)
        {
            return names.Any(name => Has(fields, name));
        }

        private static string DescribeRecord(NetworkRecord record)
        {
            return record.Fields.Count == 0
                ? "(no values)"
                : string.Join(", ", record.Fields.Select(
                    pair => pair.Key + "=" + pair.Value));
        }

        private static string CleanValue(object value)
        {
            return Convert.ToString(
                value,
                CultureInfo.InvariantCulture)?.Trim() ?? "";
        }
    }
}