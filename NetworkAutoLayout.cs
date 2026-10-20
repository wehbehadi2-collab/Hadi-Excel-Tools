
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Hadi.Excel
{
    public static class NetworkAutoLayout
    {
        // These are drawing distances, not real pipe lengths.
        private const double PreferredSpacing = 100.0;
        private const int Iterations = 250;

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
        }

        public static string GenerateCoordinates(
            IReadOnlyList<NetworkRecord> records)
        {
            if (records == null)
                throw new ArgumentNullException(nameof(records));

            var nodes = new Dictionary<string, LayoutNode>(
                StringComparer.OrdinalIgnoreCase);

            int originalCoordinateCount = 0;

            // 1. Collect each unique junction, reservoir, and tank.
            foreach (NetworkRecord record in records)
            {
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

                if (TryReadNumber(record.Get("X"), out double x))
                {
                    if (!node.HasX)
                    {
                        node.X = x;
                        node.HasX = true;
                        node.FixedX = true;
                    }
                }

                if (TryReadNumber(record.Get("Y"), out double y))
                {
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
                return "Automatic layout: no junctions, reservoirs, "
                     + "or tanks were found.";
            }

            originalCoordinateCount = nodes.Values.Count(
                n => n.HasX && n.HasY);

            // 2. Build connectivity from pipes, pumps, and valves.
            foreach (NetworkRecord record in records)
            {
                if (!IsLinkType(record.Type))
                    continue;

                string start = record.Get("Start Node")?.Trim();
                string end = record.Get("End Node")?.Trim();

                if (string.IsNullOrWhiteSpace(start)
                    || string.IsNullOrWhiteSpace(end))
                    continue;

                if (start.Equals(end, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Ignore unknown endpoints here. The normal validator
                // will report them as errors.
                if (!nodes.ContainsKey(start) || !nodes.ContainsKey(end))
                    continue;

                nodes[start].Neighbors.Add(end);
                nodes[end].Neighbors.Add(start);
            }

            // 3. Find connected components so separate networks can
            // receive separate layouts.
            var components = FindComponents(nodes);

            int generatedNodeCount = 0;
            int componentNumber = 0;

            foreach (List<LayoutNode> component in components)
            {
                componentNumber++;

                // Fully positioned nodes stay fixed. If only X or Y
                // was provided, preserve that axis and generate the other.
                InitializePositions(component, componentNumber);

                OptimizeComponent(component);

                foreach (LayoutNode node in component)
                {
                    bool originallyComplete = node.HasX && node.HasY;

                    if (!node.HasX || !node.HasY)
                        generatedNodeCount++;

                    node.HasX = true;
                    node.HasY = true;

                    string xText = FormatNumber(node.X);
                    string yText = FormatNumber(node.Y);

                    // Write the same coordinates to every record carrying
                    // this node ID, so connected links share one position.
                    foreach (NetworkRecord nodeRecord in node.Records)
                    {
                        nodeRecord.Fields["X"] = xText;
                        nodeRecord.Fields["Y"] = yText;

                        nodeRecord.AvailableFields.Add("X");
                        nodeRecord.AvailableFields.Add("Y");
                    }
                }
            }

            return
                "Automatic layout completed."
                + Environment.NewLine
                + $"Unique nodes: {nodes.Count}"
                + Environment.NewLine
                + $"Connected components: {components.Count}"
                + Environment.NewLine
                + $"Nodes with existing complete coordinates: "
                + $"{originalCoordinateCount}"
                + Environment.NewLine
                + $"Nodes requiring generated coordinates: "
                + $"{generatedNodeCount}"
                + Environment.NewLine
                + "Existing coordinate values were preserved.";
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

            foreach (LayoutNode start in nodes.Values.OrderBy(n => n.Id))
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
                ? component.Where(n => n.FixedX).Average(n => n.X)
                : (componentNumber - 1) * 500.0;

            double averageY = hasAnyY
                ? component.Where(n => n.FixedY).Average(n => n.Y)
                : 0.0;

            double radius = Math.Max(
                PreferredSpacing,
                count * PreferredSpacing / (2.0 * Math.PI));

            // Sort by ID so the same network produces a repeatable layout.
            List<LayoutNode> ordered = component
                .OrderBy(n => n.Id, StringComparer.OrdinalIgnoreCase)
                .ToList();

            for (int i = 0; i < ordered.Count; i++)
            {
                LayoutNode node = ordered[i];

                double angle = 2.0 * Math.PI * i / count;

                if (!node.FixedX)
                    node.X = averageX + radius * Math.Cos(angle);

                if (!node.FixedY)
                    node.Y = averageY + radius * Math.Sin(angle);
            }
        }

        private static void OptimizeComponent(List<LayoutNode> component)
        {
            if (component.Count <= 1)
                return;

            double temperature = PreferredSpacing;
            double cooling = temperature / Iterations;

            for (int iteration = 0; iteration < Iterations; iteration++)
            {
                var forceX = component.ToDictionary(n => n, n => 0.0);
                var forceY = component.ToDictionary(n => n, n => 0.0);

                // Repulsion: spread nodes apart to reduce overlap.
                for (int i = 0; i < component.Count; i++)
                {
                    for (int j = i + 1; j < component.Count; j++)
                    {
                        LayoutNode a = component[i];
                        LayoutNode b = component[j];

                        double dx = a.X - b.X;
                        double dy = a.Y - b.Y;

                        double distance = Math.Sqrt(dx * dx + dy * dy);

                        if (distance < 0.01)
                        {
                            dx = 0.01;
                            dy = 0.01;
                            distance = Math.Sqrt(dx * dx + dy * dy);
                        }

                        double force =
                            PreferredSpacing * PreferredSpacing / distance;

                        double fx = force * dx / distance;
                        double fy = force * dy / distance;

                        forceX[a] += fx;
                        forceY[a] += fy;
                        forceX[b] -= fx;
                        forceY[b] -= fy;
                    }
                }

                // Attraction: connected nodes stay near each other.
                foreach (LayoutNode a in component)
                {
                    foreach (string neighborId in a.Neighbors)
                    {
                        LayoutNode b = component.FirstOrDefault(
                            n => n.Id.Equals(
                                neighborId,
                                StringComparison.OrdinalIgnoreCase));

                        if (b == null
                            || string.Compare(
                                a.Id, b.Id,
                                StringComparison.OrdinalIgnoreCase) >= 0)
                            continue;

                        double dx = b.X - a.X;
                        double dy = b.Y - a.Y;

                        double distance = Math.Sqrt(dx * dx + dy * dy);

                        if (distance < 0.01)
                            continue;

                        double force =
                            distance * distance / PreferredSpacing;

                        double fx = force * dx / distance;
                        double fy = force * dy / distance;

                        forceX[a] += fx;
                        forceY[a] += fy;
                        forceX[b] -= fx;
                        forceY[b] -= fy;
                    }
                }

                // Move unfixed coordinates gradually.
                foreach (LayoutNode node in component)
                {
                    double dx = forceX[node];
                    double dy = forceY[node];

                    double magnitude = Math.Sqrt(dx * dx + dy * dy);

                    if (magnitude < 0.0001)
                        continue;

                    double scale = Math.Min(magnitude, temperature)
                                 / magnitude;

                    if (!node.FixedX)
                        node.X += dx * scale;

                    if (!node.FixedY)
                        node.Y += dy * scale;
                }

                temperature = Math.Max(1.0, temperature - cooling);
            }
        }

        private static bool TryReadNumber(string value, out double result)
        {
            result = 0;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            return double.TryParse(
                       value,
                       NumberStyles.Float | NumberStyles.AllowThousands,
                       CultureInfo.InvariantCulture,
                       out result)
                || double.TryParse(
                       value,
                       NumberStyles.Float | NumberStyles.AllowThousands,
                       CultureInfo.CurrentCulture,
                       out result);
        }

        private static string FormatNumber(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
