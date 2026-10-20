
using System;
using System.Collections.Generic;
using System.Linq;

namespace Hadi.Excel
{
    public enum NetworkFieldKind
    {
        Text,
        Number,
        PositiveNumber,
        NonNegativeNumber,
        Status,
        ValveType,
        ValveSetting
    }

    public sealed class NetworkParameterDefinition
    {
        public string Name { get; }
        public string[] Aliases { get; }
        public string[] ElementTypes { get; }
        public NetworkFieldKind Kind { get; }

        // If this header is present, the field must contain a value,
        // unless the validator implements a specific exception.
        public bool RequiredWhenColumnExists { get; }

        public NetworkParameterDefinition(
            string name,
            NetworkFieldKind kind,
            string[] elementTypes,
            bool requiredWhenColumnExists,
            params string[] aliases)
        {
            Name = name;
            Kind = kind;
            ElementTypes = elementTypes;
            RequiredWhenColumnExists = requiredWhenColumnExists;
            Aliases = aliases;
        }

        public bool AppliesTo(string elementType)
        {
            return ElementTypes.Contains(
                elementType,
                StringComparer.OrdinalIgnoreCase);
        }
    }

    public static class NetworkParameterCatalog
    {
        private static readonly NetworkParameterDefinition[] Definitions =
        {
            // -------------------- JUNCTIONS --------------------

            new NetworkParameterDefinition(
                "Node ID", NetworkFieldKind.Text,
                new[] { "Junction" }, false,
                "node", "node id", "junction", "junction id",
                "junction label", "node label", "label"),

            new NetworkParameterDefinition(
                "Elevation", NetworkFieldKind.Number,
                new[] { "Junction", "Tank" }, true,
                "elevation", "elevation (m)", "elev",
                "ground elevation", "bottom elevation"),

            new NetworkParameterDefinition(
                "Demand", NetworkFieldKind.Number,
                new[] { "Junction" }, false,
                "demand", "demand (l/s)", "base demand"),

            new NetworkParameterDefinition(
                "Demand Pattern", NetworkFieldKind.Text,
                new[] { "Junction" }, false,
                "demand pattern", "demand pattern id"),

            new NetworkParameterDefinition(
                "Emitter Coefficient", NetworkFieldKind.NonNegativeNumber,
                new[] { "Junction" }, true,
                "emitter", "emitter coefficient"),

            new NetworkParameterDefinition(
                "Initial Quality", NetworkFieldKind.Number,
                new[] { "Junction", "Reservoir", "Tank" }, false,
                "initial quality", "initial water quality"),

            // -------------------- RESERVOIRS --------------------

            new NetworkParameterDefinition(
                "Reservoir ID", NetworkFieldKind.Text,
                new[] { "Reservoir" }, false,
                "reservoir", "reservoir id", "reservoir label"),

            new NetworkParameterDefinition(
                "Head", NetworkFieldKind.Number,
                new[] { "Reservoir" }, true,
                "head", "hydraulic head", "total head", "head (m)"),

            new NetworkParameterDefinition(
                "Head Pattern", NetworkFieldKind.Text,
                new[] { "Reservoir" }, false,
                "head pattern", "head pattern id"),

            // -------------------- TANKS --------------------

            new NetworkParameterDefinition(
                "Tank ID", NetworkFieldKind.Text,
                new[] { "Tank" }, false,
                "tank", "tank id", "tank label"),

            new NetworkParameterDefinition(
                "Initial Level", NetworkFieldKind.NonNegativeNumber,
                new[] { "Tank" }, true,
                "initial level", "initial water level",
                "initial level (m)"),

            new NetworkParameterDefinition(
                "Minimum Level", NetworkFieldKind.NonNegativeNumber,
                new[] { "Tank" }, true,
                "minimum level", "min level", "minimum water level"),

            new NetworkParameterDefinition(
                "Maximum Level", NetworkFieldKind.NonNegativeNumber,
                new[] { "Tank" }, true,
                "maximum level", "max level", "maximum water level"),

            // Tank diameter is in metres, unlike pipe diameter in mm.
            new NetworkParameterDefinition(
                "Tank Diameter", NetworkFieldKind.PositiveNumber,
                new[] { "Tank" }, true,
                "tank diameter", "tank diameter (m)"),

            new NetworkParameterDefinition(
                "Minimum Volume", NetworkFieldKind.NonNegativeNumber,
                new[] { "Tank" }, false,
                "minimum volume", "min volume"),

            new NetworkParameterDefinition(
                "Volume Curve", NetworkFieldKind.Text,
                new[] { "Tank" }, false,
                "volume curve", "volume curve id"),

            // -------------------- PIPES --------------------

            new NetworkParameterDefinition(
                "Pipe ID", NetworkFieldKind.Text,
                new[] { "Pipe" }, false,
                "pipe", "pipe id", "pipe label"),

            new NetworkParameterDefinition(
                "Start Node", NetworkFieldKind.Text,
                new[] { "Pipe", "Pump", "Valve" }, false,
                "start node", "start node id", "from node",
                "node1", "node 1"),

            new NetworkParameterDefinition(
                "End Node", NetworkFieldKind.Text,
                new[] { "Pipe", "Pump", "Valve" }, false,
                "end node", "end node id", "stop node", "to node",
                "node2", "node 2"),

            new NetworkParameterDefinition(
                "Length", NetworkFieldKind.PositiveNumber,
                new[] { "Pipe" }, true,
                "length", "length (m)", "pipe length"),

            // Pipe and valve diameter is in millimetres in our workflow.
            new NetworkParameterDefinition(
                "Diameter", NetworkFieldKind.PositiveNumber,
                new[] { "Pipe", "Valve" }, true,
                "diameter", "diameter (mm)", "pipe diameter"),

            new NetworkParameterDefinition(
                "Roughness", NetworkFieldKind.PositiveNumber,
                new[] { "Pipe" }, true,
                "roughness", "roughness (mm)",
                "absolute roughness", "absolute roughness (mm)"),

            new NetworkParameterDefinition(
                "Minor Loss", NetworkFieldKind.NonNegativeNumber,
                new[] { "Pipe", "Valve" }, true,
                "minor loss", "minor loss coefficient",
                "minor loss coefficient (k)", "loss coefficient"),

            new NetworkParameterDefinition(
                "Status", NetworkFieldKind.Status,
                new[] { "Pipe", "Pump", "Valve" }, false,
                "status", "initial status"),

            new NetworkParameterDefinition(
                "Material", NetworkFieldKind.Text,
                new[] { "Pipe" }, false,
                "material", "pipe material"),

            new NetworkParameterDefinition(
                "Bulk Reaction", NetworkFieldKind.Number,
                new[] { "Pipe" }, false,
                "bulk reaction", "bulk reaction coefficient"),

            new NetworkParameterDefinition(
                "Wall Reaction", NetworkFieldKind.Number,
                new[] { "Pipe" }, false,
                "wall reaction", "wall reaction coefficient"),

            // -------------------- PUMPS --------------------

            new NetworkParameterDefinition(
                "Pump ID", NetworkFieldKind.Text,
                new[] { "Pump" }, false,
                "pump", "pump id", "pump label"),

            new NetworkParameterDefinition(
                "Pump Curve", NetworkFieldKind.Text,
                new[] { "Pump" }, false,
                "pump curve", "pump curve id", "curve id"),

            new NetworkParameterDefinition(
                "Power", NetworkFieldKind.PositiveNumber,
                new[] { "Pump" }, true,
                "power", "pump power", "power (kw)", "power (hp)"),

            new NetworkParameterDefinition(
                "Speed", NetworkFieldKind.PositiveNumber,
                new[] { "Pump" }, false,
                "speed", "pump speed", "relative speed"),

            new NetworkParameterDefinition(
                "Speed Pattern", NetworkFieldKind.Text,
                new[] { "Pump" }, false,
                "speed pattern", "speed pattern id"),

            // -------------------- VALVES --------------------

            new NetworkParameterDefinition(
                "Valve ID", NetworkFieldKind.Text,
                new[] { "Valve" }, false,
                "valve", "valve id", "valve label"),

            new NetworkParameterDefinition(
                "Valve Type", NetworkFieldKind.ValveType,
                new[] { "Valve" }, true,
                "valve type", "type of valve", "valve kind"),

            new NetworkParameterDefinition(
                "Setting", NetworkFieldKind.ValveSetting,
                new[] { "Valve" }, true,
                "setting", "valve setting", "pressure setting",
                "flow setting", "valve control setting"),

            // -------------------- MAP COORDINATES --------------------

            new NetworkParameterDefinition(
                "X", NetworkFieldKind.Number,
                new[] { "Junction", "Reservoir", "Tank" }, false,
                "x", "x coordinate", "x-coordinate"),

            new NetworkParameterDefinition(
                "Y", NetworkFieldKind.Number,
                new[] { "Junction", "Reservoir", "Tank" }, false,
                "y", "y coordinate", "y-coordinate")
        };

        private static readonly Dictionary<string, string> AliasMap =
            BuildAliasMap();

        private static Dictionary<string, string> BuildAliasMap()
        {
            var map = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);

            foreach (NetworkParameterDefinition definition in Definitions)
            {
                AddAlias(map, definition.Name, definition.Name);

                foreach (string alias in definition.Aliases)
                    AddAlias(map, alias, definition.Name);
            }

            return map;
        }

        private static void AddAlias(
            Dictionary<string, string> map,
            string alias,
            string canonicalName)
        {
            string key = Normalize(alias);

            if (map.TryGetValue(key, out string existing))
            {
                if (!existing.Equals(
                        canonicalName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    // An ambiguous alias is not allowed to silently
                    // replace an existing mapping.
                    return;
                }
            }

            map[key] = canonicalName;
        }

        public static IReadOnlyList<NetworkParameterDefinition> All =>
            Definitions;

        public static string Normalize(string text)
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

        public static bool TryGetCanonicalName(
            string header,
            out string canonicalName)
        {
            return AliasMap.TryGetValue(
                Normalize(header),
                out canonicalName);
        }

        public static NetworkParameterDefinition Get(string name)
        {
            return Definitions.FirstOrDefault(
                d => d.Name.Equals(
                    name,
                    StringComparison.OrdinalIgnoreCase));
        }

        public static bool AppliesTo(string field, string elementType)
        {
            NetworkParameterDefinition definition = Get(field);

            return definition != null
                && definition.AppliesTo(elementType);
        }

        public static IEnumerable<string> FieldsFor(string elementType)
        {
            return Definitions
                .Where(d => d.AppliesTo(elementType))
                .Select(d => d.Name);
        }

        public static string[] RequiredFields(string elementType)
        {
            switch (elementType)
            {
                case "Junction":
                    return new[] { "Node ID", "Elevation" };

                case "Reservoir":
                    return new[] { "Reservoir ID", "Head" };

                case "Pipe":
                    return new[]
                    {
                        "Pipe ID", "Start Node", "End Node",
                        "Length", "Diameter"
                    };

                case "Tank":
                    return new[]
                    {
                        "Tank ID", "Elevation", "Initial Level",
                        "Minimum Level", "Maximum Level", "Tank Diameter"
                    };

                case "Pump":
                    // The validator separately enforces Pump Curve OR Power.
                    return new[]
                    {
                        "Pump ID", "Start Node", "End Node"
                    };

                case "Valve":
                    return new[]
                    {
                        "Valve ID", "Start Node", "End Node",
                        "Diameter", "Valve Type", "Setting"
                    };

                default:
                    return Array.Empty<string>();
            }
        }
    }
}
