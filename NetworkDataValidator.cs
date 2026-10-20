using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Hadi.Excel
{
    public sealed class NetworkRecord
    {
        public string Type { get; set; } = "";

        public Dictionary<string, string> Fields { get; } =
            new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);

        public HashSet<string> AvailableFields { get; } =
            new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);

        public string Source { get; set; } = "";

        public string Get(string field)
        {
            return Fields.TryGetValue(field, out string value)
                ? value
                : "";
        }
    }

    public static class NetworkDataValidator
    {
        public static string Validate(List<NetworkRecord> records)
        {
            if (records == null)
                throw new ArgumentNullException(nameof(records));

            var errors = new List<string>();

            ValidateIdentifiers(records, errors);

            foreach (NetworkRecord record in records)
            {
                ValidateRequiredFields(record, errors);
                ValidateAvailableFields(record, errors);
                ValidateElementRules(record, errors);
            }

            ValidateConnections(records, errors);

            var report = new StringBuilder();

            report.AppendLine("Network validation report");
            report.AppendLine();

            report.AppendLine($"Records checked: {records.Count}");

            foreach (string type in new[]
            {
                "Junction", "Reservoir", "Pipe",
                "Tank", "Pump", "Valve"
            })
            {
                report.AppendLine(
                    $"{type}s: {records.Count(r => r.Type == type)}");
            }

            report.AppendLine();

            if (errors.Count == 0)
            {
                report.AppendLine("VALIDATION PASSED");
                report.AppendLine(
                    "No errors were detected in the supported fields.");
            }
            else
            {
                report.AppendLine(
                    $"VALIDATION FAILED — {errors.Count} error(s)");

                foreach (string error in errors)
                    report.AppendLine("- " + error);
            }

            report.AppendLine();
            report.AppendLine(
                "This validator checks recognized fields. "
                + "It does not yet generate an INP file.");

            return report.ToString();
        }

        private static void ValidateIdentifiers(
            List<NetworkRecord> records,
            List<string> errors)
        {
            var ids = new Dictionary<string, NetworkRecord>(
                StringComparer.OrdinalIgnoreCase);

            foreach (NetworkRecord record in records)
            {
                string id = GetId(record);

                if (string.IsNullOrWhiteSpace(id))
                    continue;

                // Node IDs must be unique across junctions, reservoirs,
                // and tanks. Link IDs must be unique across pipes,
                // pumps, and valves.
                string group = IsNode(record.Type) ? "NODE" : "LINK";
                string key = group + ":" + id;

                if (ids.TryGetValue(key, out NetworkRecord previous))
                {
                    errors.Add(
                        $"{record.Source}: duplicate {group.ToLowerInvariant()} "
                        + $"ID '{id}' (also used by {previous.Type} at "
                        + $"{previous.Source}).");
                }
                else
                {
                    ids[key] = record;
                }
            }
        }

        private static void ValidateRequiredFields(
            NetworkRecord record,
            List<string> errors)
        {
            foreach (string field in
                     NetworkParameterCatalog.RequiredFields(record.Type))
            {
                if (string.IsNullOrWhiteSpace(record.Get(field)))
                {
                    errors.Add(
                        $"{record.Source}: {record.Type} "
                        + $"'{GetId(record)}' requires {field}.");
                }
            }

            if (record.Type == "Pump"
                && string.IsNullOrWhiteSpace(record.Get("Pump Curve"))
                && string.IsNullOrWhiteSpace(record.Get("Power")))
            {
                errors.Add(
                    $"{record.Source}: pump '{GetId(record)}' "
                    + "requires either Pump Curve or Power.");
            }
        }

        private static void ValidateAvailableFields(
            NetworkRecord record,
            List<string> errors)
        {
            HashSet<string> required = new HashSet<string>(
                NetworkParameterCatalog.RequiredFields(record.Type),
                StringComparer.OrdinalIgnoreCase);

            foreach (NetworkParameterDefinition definition
                     in NetworkParameterCatalog.All)
            {
                if (!definition.AppliesTo(record.Type))
                    continue;

                if (!record.AvailableFields.Contains(definition.Name))
                    continue;

                string value = record.Get(definition.Name);

                // Required fields are already checked separately.
                if (required.Contains(definition.Name))
                    continue;

                // A pump can use either a curve or power. A blank Power
                // cell is acceptable if a pump curve is provided.
                if (record.Type == "Pump"
                    && definition.Name == "Power"
                    && string.IsNullOrWhiteSpace(value)
                    && !string.IsNullOrWhiteSpace(record.Get("Pump Curve")))
                {
                    continue;
                }

                // Some optional fields can be left blank. Others, such as
                // roughness or minor loss when their column exists, need
                // a valid value for each applicable record.
                if (string.IsNullOrWhiteSpace(value))
                {
                    if (definition.RequiredWhenColumnExists)
                    {
                        errors.Add(
                            $"{record.Source}: {record.Type} "
                            + $"'{GetId(record)}' has a blank "
                            + $"{definition.Name} value.");
                    }

                    continue;
                }

                ValidateFieldValue(record, definition, value, errors);
            }
        }

        private static void ValidateFieldValue(
            NetworkRecord record,
            NetworkParameterDefinition definition,
            string value,
            List<string> errors)
        {
            bool valid = true;

            switch (definition.Kind)
            {
                case NetworkFieldKind.Text:
                    // Text fields are checked for presence where required.
                    break;

                case NetworkFieldKind.Number:
                    valid = TryParseNumber(value, out _);
                    break;

                case NetworkFieldKind.PositiveNumber:
                    valid = TryParseNumber(value, out double positive)
                        && positive > 0;
                    break;

                case NetworkFieldKind.NonNegativeNumber:
                    valid = TryParseNumber(value, out double nonNegative)
                        && nonNegative >= 0;
                    break;

                case NetworkFieldKind.Status:
                    valid = IsValidStatus(record.Type, value);
                    break;

                case NetworkFieldKind.ValveType:
                    valid = IsValidValveType(value);
                    break;

                case NetworkFieldKind.ValveSetting:
                    if (record.Get("Valve Type").Equals(
                            "GPV",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        // A GPV setting identifies a headloss curve.
                        valid = !string.IsNullOrWhiteSpace(value);
                    }
                    else
                    {
                        valid = TryParseNumber(value, out double setting)
                            && setting >= 0;
                    }
                    break;
            }

            if (!valid)
            {
                errors.Add(
                    $"{record.Source}: {record.Type} '{GetId(record)}' "
                    + $"has an invalid {definition.Name} value '{value}'.");
            }
        }

        private static void ValidateElementRules(
            NetworkRecord record,
            List<string> errors)
        {
            if (record.Type == "Tank")
            {
                CheckLevelOrder(record, errors);
            }
        }

        private static void CheckLevelOrder(
            NetworkRecord record,
            List<string> errors)
        {
            bool hasInitial = TryParseNumber(
                record.Get("Initial Level"), out double initial);

            bool hasMinimum = TryParseNumber(
                record.Get("Minimum Level"), out double minimum);

            bool hasMaximum = TryParseNumber(
                record.Get("Maximum Level"), out double maximum);

            if (hasInitial && hasMinimum && minimum > initial)
            {
                errors.Add(
                    $"{record.Source}: tank '{GetId(record)}' has a "
                    + "Minimum Level greater than its Initial Level.");
            }

            if (hasInitial && hasMaximum && initial > maximum)
            {
                errors.Add(
                    $"{record.Source}: tank '{GetId(record)}' has an "
                    + "Initial Level greater than its Maximum Level.");
            }

            if (hasMinimum && hasMaximum && minimum >= maximum)
            {
                errors.Add(
                    $"{record.Source}: tank '{GetId(record)}' must have "
                    + "a Minimum Level lower than its Maximum Level.");
            }
        }

        private static void ValidateConnections(
            List<NetworkRecord> records,
            List<string> errors)
        {
            var nodeIds = new HashSet<string>(
                records
                    .Where(r => IsNode(r.Type))
                    .Select(GetId)
                    .Where(id => !string.IsNullOrWhiteSpace(id)),
                StringComparer.OrdinalIgnoreCase);

            foreach (NetworkRecord link in records.Where(
                         r => r.Type == "Pipe"
                           || r.Type == "Pump"
                           || r.Type == "Valve"))
            {
                string id = GetId(link);
                string start = link.Get("Start Node");
                string end = link.Get("End Node");

                CheckEndpoint(link, id, "Start Node",
                    start, nodeIds, errors);

                CheckEndpoint(link, id, "End Node",
                    end, nodeIds, errors);

                if (!string.IsNullOrWhiteSpace(start)
                    && !string.IsNullOrWhiteSpace(end)
                    && start.Equals(
                        end,
                        StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add(
                        $"{link.Source}: {link.Type.ToLowerInvariant()} "
                        + $"'{id}' starts and ends at the same node.");
                }
            }
        }

        private static void CheckEndpoint(
            NetworkRecord link,
            string id,
            string field,
            string node,
            HashSet<string> nodeIds,
            List<string> errors)
        {
            if (string.IsNullOrWhiteSpace(node))
                return; // Missing values are reported by required checks.

            if (!nodeIds.Contains(node))
            {
                errors.Add(
                    $"{link.Source}: {link.Type.ToLowerInvariant()} "
                    + $"'{id}' references unknown {field} '{node}'.");
            }
        }

        private static bool IsNode(string type)
        {
            return type == "Junction"
                || type == "Reservoir"
                || type == "Tank";
        }

        private static string GetId(NetworkRecord record)
        {
            switch (record.Type)
            {
                case "Junction":
                    return record.Get("Node ID");
                case "Reservoir":
                    return record.Get("Reservoir ID");
                case "Tank":
                    return record.Get("Tank ID");
                case "Pipe":
                    return record.Get("Pipe ID");
                case "Pump":
                    return record.Get("Pump ID");
                case "Valve":
                    return record.Get("Valve ID");
                default:
                    return "";
            }
        }

        private static bool IsValidStatus(
            string elementType,
            string value)
        {
            string status = value.Trim().ToUpperInvariant();

            switch (elementType)
            {
                case "Pipe":
                    return status == "OPEN"
                        || status == "CLOSED"
                        || status == "CV";

                case "Pump":
                    return status == "OPEN"
                        || status == "CLOSED";

                case "Valve":
                    return status == "OPEN"
                        || status == "CLOSED"
                        || status == "ACTIVE";

                default:
                    return false;
            }
        }

        private static bool IsValidValveType(string value)
        {
            string type = value.Trim().ToUpperInvariant();

            return type == "PRV"
                || type == "PSV"
                || type == "PBV"
                || type == "FCV"
                || type == "TCV"
                || type == "GPV";
        }

        private static bool TryParseNumber(
            string value,
            out double number)
        {
            number = 0;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            if (double.TryParse(
                    value,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out number))
            {
                return !double.IsNaN(number)
                    && !double.IsInfinity(number);
            }

            if (double.TryParse(
                    value,
                    NumberStyles.Float,
                    CultureInfo.CurrentCulture,
                    out number))
            {
                return !double.IsNaN(number)
                    && !double.IsInfinity(number);
            }

            return false;
        }
    }
}