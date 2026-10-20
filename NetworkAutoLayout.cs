
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Hadi.Excel
{
    public static class NetworkAutoLayout
    {
        // These are drawing distances, not real pipe lengths.
        private const double PreferredSpacing = 100.0;
        private const int Iterations = 250;
        private const double CoordinateTolerance = 0.000000001;

        private sealed class LayoutNode
        {
            public string Id { get; set; } = "";

            public List<NetworkRecord> Records { get; } = new();

            public HashSet<string> Neighbors { get; } =
                new(StringComparer.OrdinalIgnoreCase);

            public double X { get; set; }
            public double Y { get; set; }

            public bool HasX { get; set; }
            public bool HasY { get; set; }

            public bool FixedX { get; set; }
            public bool FixedY { get; set; }

            public int ComponentNumber { get; set; }

            public string CoordinateSource
            {
                get
                {
                    if (FixedX && FixedY)
                        return "Preserved X and Y";

                    if (FixedX)
                        return "Preserved X; generated Y";

                    if (FixedY)
                        return "Generated X; preserved Y";

                    return "Generated X and Y";
                }
            }
        }

        public static string GenerateCoordinates(
            IReadOnlyList<NetworkRecord> records)
        {
            if (records == null)
                throw new ArgumentNullException(nameof(records));

            var nodes = new Dictionary<string, LayoutNode>(
                StringComparer.OrdinalIgnoreCase);

            // 1. Collect nodes and validate their existing coordinates.
            // Conflicts are detected before any record is modified.
            foreach (NetworkRecord record in records)
            {
                if (record == null)
                    continue;

                string idField = GetNodeIdField(record.Type);

                if (idField == null)
                    continue;

                string id = record.Get(idField)?.Trim();

                if (string.IsNullOrWhiteSpace(id))
                    continue;

                if (!nodes.TryGetValue(id, out LayoutNode node))
                {
                    node = new LayoutNode { Id = id };
                    nodes.Add(id, node);
                }

                node.Records.Add(record);

                string xValue = record.Get("X");
                string yValue = record.Get("Y");

                if (TryReadNumber(xValue, out double x))
                {
                    if (node.HasX &&
                        !CoordinatesMatch(node.X, x))
                    {
                        throw new InvalidOperationException(
                            BuildCoordinateConflictMessage(
                                node.Id,
                                "X",
                                node.X,
                                x,
                                xValue));
                    }

                    if (!node.HasX)
                    {
                        node.X = x;
                        node.HasX = true;
                        node.FixedX = true;
                    }
                }

                if (TryReadNumber(yValue, out double y))
                {
                    if (node.HasY &&
                        !CoordinatesMatch(node.Y, y))
                    {
                        throw new InvalidOperationException(
                            BuildCoordinateConflictMessage(
                                node.Id,
                                "Y",
                                node.Y,
                                y,
                                yValue));
                    }

                    if (!node.HasY)
                    {
                        node.Y = y;
                        node.HasY = true;
                        node.FixedY = true;
                    }
                }
            }

            if (nodes.Count == 0)
            {
                return
                    "Automatic layout: no junctions, reservoirs, " +
                    "or tanks were found.";
            }

            int originalCompleteCoordinateCount =
                nodes.Values.Count(n => n.HasX && n.HasY);

            int originalPartialCoordinateCount =
                nodes.Values.Count(n => n.HasX != n.HasY);

            int originalMissingCoordinateCount =
                nodes.Values.Count(n => !n.HasX && !n.HasY);

            // 2. Build connectivity from pipes, pumps, and valves.
            foreach (NetworkRecord record in records)
            {
                if (record == null || !IsLinkType(record.Type))
                    continue;

                string start = record.Get("Start Node")?.Trim();
                string end = record.Get("End Node")?.Trim();

                if (string.IsNullOrWhiteSpace(start) ||
                    string.IsNullOrWhiteSpace(end))
                {
                    continue;
                }

                if (start.Equals(
                    end,
                    StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Unknown endpoints are handled by normal validation.
                if (!nodes.ContainsKey(start) ||
                    !nodes.ContainsKey(end))
                {
                    continue;
                }

                nodes[start].Neighbors.Add(end);
                nodes[end].Neighbors.Add(start);
            }

            // 3. Find connected components.
            List<List<LayoutNode>> components =
                FindComponents(nodes);

            int componentNumber = 0;

            foreach (List<LayoutNode> component in components)
            {
                componentNumber++;

                foreach (LayoutNode node in component)
                    node.ComponentNumber = componentNumber;

                InitializePositions(component, componentNumber);
                OptimizeComponent(component);
            }

            // 4. Prepare the complete report before writing coordinates.
            // This also makes the coordinate source explicit for every node.
            var report = new StringBuilder();

            int generatedCoordinateNodeCount =
                nodes.Values.Count(n => !n.HasX || !n.HasY);

            int isolatedNodeCount =
                nodes.Values.Count(n => n.Neighbors.Count == 0);

            report.AppendLine("Automatic layout completed.");
            report.AppendLine(
                $"Unique nodes: {nodes.Count}");
            report.AppendLine(
                $"Connected components: {components.Count}");
            report.AppendLine(
                $"Nodes with existing complete coordinates: " +
                $"{originalCompleteCoordinateCount}");
            report.AppendLine(
                $"Nodes with partially existing coordinates: " +
                $"{originalPartialCoordinateCount}");
            report.AppendLine(
                $"Nodes with no existing coordinates: " +
                $"{originalMissingCoordinateCount}");
            report.AppendLine(
                $"Nodes requiring at least one generated coordinate: " +
                $"{generatedCoordinateNodeCount}");
            report.AppendLine(
                $"Isolated nodes (no recognized links): " +
                $"{isolatedNodeCount}");
            report.AppendLine();
            report.AppendLine("CONNECTED COMPONENTS");

            for (int i = 0; i < components.Count; i++)
            {
                List<LayoutNode> component = components[i];

                string nodeIds = string.Join(
                    ", ",
                    component
                        .OrderBy(
                            n => n.Id,
                            StringComparer.OrdinalIgnoreCase)
                        .Select(n => n.Id));

                report.AppendLine(
                    $"Component {i + 1}: {component.Count} node(s)");

                report.AppendLine($"  Nodes: {nodeIds}");
            }

            report.AppendLine();
            report.AppendLine("FINAL NODE COORDINATES");
            report.AppendLine(
                "Coordinates are layout positions, not pipe lengths.");

            foreach (LayoutNode node in nodes.Values
                .OrderBy(
                    n => n.Id,
                    StringComparer.OrdinalIgnoreCase))
            {
                report.AppendLine(
                    $"{node.Id}: X={FormatNumber(node.X)}, " +
                    $"Y={FormatNumber(node.Y)} | " +
                    $"Component {node.ComponentNumber} | " +
                    node.CoordinateSource);
            }

            report.AppendLine();
            report.AppendLine(
                "Existing valid coordinate axes were preserved.");

            // 5. Write identical coordinates to every record for each node.
            foreach (LayoutNode node in nodes.Values)
            {
                string xText = FormatNumber(node.X);
                string yText = FormatNumber(node.Y);

                foreach (NetworkRecord nodeRecord in node.Records)
                {
                    nodeRecord.Fields["X"] = xText;
                    nodeRecord.Fields["Y"] = yText;

                    nodeRecord.AvailableFields.Add("X");
                    nodeRecord.AvailableFields.Add("Y");
                }
            }

            return report.ToString();
        }

        private static string GetNodeIdField(string elementType)
        {
            switch (elementType)
            {
                case "Junction":
                    return "Node ID";

                case "Reservoir":
                    return "Reservoir ID";

                case "Tank":
                    return "Tank ID";

                default:
                    return null;
            }
        }

        private static bool IsLinkType(string elementType)
        {
            return elementType == "Pipe"
                || elementType == "Pump"
                || elementType == "Valve";
        }

        private static List<List<LayoutNode>> FindComponents(
            Dictionary<string, LayoutNode> nodes)
        {
            var components = new List<List<LayoutNode>>();

            var visited = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);

            foreach (LayoutNode start in nodes.Values
                .OrderBy(
                    n => n.Id,
                    StringComparer.OrdinalIgnoreCase))
            {
                if (!visited.Add(start.Id))
                    continue;

                var component = new List<LayoutNode>();
                var queue = new Queue<LayoutNode>();

                queue.Enqueue(start);

                while (queue.Count > 0)
                {
                    LayoutNode current = queue.Dequeue();
                    component.Add(current);

                    foreach (string neighborId in current.Neighbors)
                    {
                        if (!visited.Add(neighborId))
                            continue;

                        queue.Enqueue(nodes[neighborId]);
                    }
                }

                components.Add(component);
            }

            return components;
        }

        private static void InitializePositions(
            List<LayoutNode> component,
            int componentNumber)
        {
            int count = component.Count;

            bool hasAnyX = component.Any(n => n.FixedX);
            bool hasAnyY = component.Any(n => n.FixedY);

            double averageX = hasAnyX
                ? component
                    .Where(n => n.FixedX)
                    .Average(n => n.X)
                : (componentNumber - 1) * 500.0;

            double averageY = hasAnyY
                ? component
                    .Where(n => n.FixedY)
                    .Average(n => n.Y)
                : 0.0;

            double radius = Math.Max(
                PreferredSpacing,
                count * PreferredSpacing / (2.0 * Math.PI));

            // Stable ordering makes initialization repeatable.
            List<LayoutNode> ordered = component
                .OrderBy(
                    n => n.Id,
                    StringComparer.OrdinalIgnoreCase)
                .ToList();

            for (int i = 0; i < ordered.Count; i++)
            {
                LayoutNode node = ordered[i];

                double angle = 2.0 * Math.PI * i / count;

                if (!node.FixedX)
                {
                    node.X = averageX +
                        radius * Math.Cos(angle);
                }

                if (!node.FixedY)
                {
                    node.Y = averageY +
                        radius * Math.Sin(angle);
                }
            }
        }

        private static void OptimizeComponent(
            List<LayoutNode> component)
        {
            if (component.Count <= 1)
                return;

            // Direct node lookup avoids repeatedly searching the component
            // for each connected neighbor during every iteration.
            var nodesById = component.ToDictionary(
                n => n.Id,
                StringComparer.OrdinalIgnoreCase);

            double temperature = PreferredSpacing;
            double cooling = temperature / Iterations;

            for (int iteration = 0;
                 iteration < Iterations;
                 iteration++)
            {
                var forceX = component.ToDictionary(
                    n => n,
                    n => 0.0);

                var forceY = component.ToDictionary(
                    n => n,
                    n => 0.0);

                // Repulsion: spread nodes apart.
                for (int i = 0; i < component.Count; i++)
                {
                    for (int j = i + 1;
                         j < component.Count;
                         j++)
                    {
                        LayoutNode a = component[i];
                        LayoutNode b = component[j];

                        double dx = a.X - b.X;
                        double dy = a.Y - b.Y;

                        double distance =
                            Math.Sqrt(dx * dx + dy * dy);

                        if (distance < 0.01)
                        {
                            dx = 0.01;
                            dy = 0.01;

                            distance = Math.Sqrt(
                                dx * dx + dy * dy);
                        }

                        double force =
                            PreferredSpacing *
                            PreferredSpacing / distance;

                        double fx = force * dx / distance;
                        double fy = force * dy / distance;

                        forceX[a] += fx;
                        forceY[a] += fy;

                        forceX[b] -= fx;
                        forceY[b] -= fy;
                    }
                }

                // Attraction: keep connected nodes relatively close.
                // Each connection is processed only once.
                foreach (LayoutNode a in component)
                {
                    foreach (string neighborId in a.Neighbors)
                    {
                        if (!nodesById.TryGetValue(
                            neighborId,
                            out LayoutNode b))
                        {
                            continue;
                        }

                        if (string.Compare(
                            a.Id,
                            b.Id,
                            StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            continue;
                        }

                        double dx = b.X - a.X;
                        double dy = b.Y - a.Y;

                        double distance =
                            Math.Sqrt(dx * dx + dy * dy);

                        if (distance < 0.01)
                            continue;

                        double force =
                            distance * distance /
                            PreferredSpacing;

                        double fx = force * dx / distance;
                        double fy = force * dy / distance;

                        forceX[a] += fx;
                        forceY[a] += fy;

                        forceX[b] -= fx;
                        forceY[b] -= fy;
                    }
                }

                // Move only coordinate axes that were not supplied.
                foreach (LayoutNode node in component)
                {
                    double dx = forceX[node];
                    double dy = forceY[node];

                    double magnitude =
                        Math.Sqrt(dx * dx + dy * dy);

                    if (magnitude < 0.0001)
                        continue;

                    double scale =
                        Math.Min(magnitude, temperature) /
                        magnitude;

                    if (!node.FixedX)
                        node.X += dx * scale;

                    if (!node.FixedY)
                        node.Y += dy * scale;
                }

                temperature = Math.Max(
                    1.0,
                    temperature - cooling);
            }
        }

        private static bool TryReadNumber(
            string value,
            out double result)
        {
            result = 0;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            bool parsed =
                double.TryParse(
                    value,
                    NumberStyles.Float |
                    NumberStyles.AllowThousands,
                    CultureInfo.InvariantCulture,
                    out result)
                ||
                double.TryParse(
                    value,
                    NumberStyles.Float |
                    NumberStyles.AllowThousands,
                    CultureInfo.CurrentCulture,
                    out result);

            return parsed && double.IsFinite(result);
        }

        private static bool CoordinatesMatch(
            double first,
            double second)
        {
            return Math.Abs(first - second) <=
                CoordinateTolerance;
        }

        private static string BuildCoordinateConflictMessage(
            string nodeId,
            string axis,
            double existingValue,
            double conflictingValue,
            string conflictingText)
        {
            return
                "Automatic layout cancelled because conflicting " +
                "coordinates were found." +
                Environment.NewLine +
                $"Node ID: {nodeId}" +
                Environment.NewLine +
                $"Coordinate axis: {axis}" +
                Environment.NewLine +
                $"First valid value: {FormatNumber(existingValue)}" +
                Environment.NewLine +
                $"Conflicting value: {conflictingText} " +
                $"({FormatNumber(conflictingValue)})" +
                Environment.NewLine +
                "Check all Excel records that refer to this node. " +
                "No coordinates were written by the layout operation.";
        }

        private static string FormatNumber(double value)
        {
            return value.ToString(
                "0.###",
                CultureInfo.InvariantCulture);
        }
    }
}
