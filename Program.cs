
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace WaterCadTest48
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.Title = "WaterCAD Network Generator - NEW VERSION";

            Console.WriteLine("========================================");
            Console.WriteLine(" WATER CAD NETWORK GENERATOR - NEW V1 ");
            Console.WriteLine("========================================");
            Console.WriteLine();

            string desktop = Environment.GetFolderPath(
                Environment.SpecialFolder.DesktopDirectory);

            string inputFolder = Path.Combine(desktop, "WaterCAD_Input");
            string outputPath = Path.Combine(
                desktop, "WaterCAD_TestNetwork.inp");

            Console.WriteLine("RUNNING EXECUTABLE:");
            Console.WriteLine(Assembly.GetExecutingAssembly().Location);
            Console.WriteLine();

            Console.WriteLine("INPUT FOLDER:");
            Console.WriteLine(inputFolder);
            Console.WriteLine();

            Console.WriteLine("OUTPUT FILE:");
            Console.WriteLine(outputPath);
            Console.WriteLine();

            try
            {
                WaterNetwork network = NetworkCsvReader.Load(inputFolder);

                Console.WriteLine("CSV FILES READ SUCCESSFULLY.");
                Console.WriteLine();

                Console.WriteLine("PIPES READ FROM CSV:");
                foreach (Pipe pipe in network.Pipes)
                {
                    Console.WriteLine(
                        pipe.Id +
                        " | Length=" + Format(pipe.Length) +
                        " | Diameter=" + Format(pipe.Diameter) +
                        " | Roughness=" + Format(pipe.Roughness) +
                        " | MinorLoss=" + Format(pipe.MinorLoss));
                }

                Console.WriteLine();
                Console.WriteLine("VALIDATING NETWORK...");

                List<string> errors = NetworkValidator.Validate(network);

                if (errors.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("VALIDATION FAILED:");

                    foreach (string error in errors)
                    {
                        Console.WriteLine("- " + error);
                    }

                    Console.WriteLine();
                    Console.WriteLine("No INP file was generated.");
                    return;
                }

                Console.WriteLine("VALIDATION PASSED.");
                Console.WriteLine();

                EpanetFileWriter.Write(network, outputPath);

                Console.WriteLine("========================================");
                Console.WriteLine("INP FILE CREATED SUCCESSFULLY.");
                Console.WriteLine("========================================");
                Console.WriteLine(outputPath);
                Console.WriteLine();

                Console.WriteLine("PIPES WRITTEN TO THE INP FILE:");

                foreach (Pipe pipe in network.Pipes)
                {
                    Console.WriteLine(
                        pipe.Id + " | " +
                        pipe.StartNode + " -> " + pipe.EndNode +
                        " | Length=" + Format(pipe.Length) +
                        " | Diameter=" + Format(pipe.Diameter));
                }

                Console.WriteLine();
                Console.WriteLine("Open the INP file in Notepad to verify it.");
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("ERROR:");
                Console.WriteLine(ex.Message);
                Console.WriteLine();
                Console.WriteLine("Details:");
                Console.WriteLine(ex.ToString());
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static string Format(double value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }
    }

    public class WaterNetwork
    {
        public List<Junction> Junctions { get; private set; }
        public List<Reservoir> Reservoirs { get; private set; }
        public List<Pipe> Pipes { get; private set; }

        public WaterNetwork()
        {
            Junctions = new List<Junction>();
            Reservoirs = new List<Reservoir>();
            Pipes = new List<Pipe>();
        }
    }

    public class Junction
    {
        public string Id { get; set; }
        public double Elevation { get; set; }
        public double Demand { get; set; }
    }

    public class Reservoir
    {
        public string Id { get; set; }
        public double Head { get; set; }
    }

    public class Pipe
    {
        public string Id { get; set; }
        public string StartNode { get; set; }
        public string EndNode { get; set; }
        public double Length { get; set; }
        public double Diameter { get; set; }
        public double Roughness { get; set; }
        public double MinorLoss { get; set; }
    }

    public static class NetworkCsvReader
    {
        public static WaterNetwork Load(string folder)
        {
            if (!Directory.Exists(folder))
            {
                throw new DirectoryNotFoundException(
                    "Input folder not found: " + folder);
            }

            WaterNetwork network = new WaterNetwork();

            ReadJunctions(
                Path.Combine(folder, "Junctions.csv"), network);

            ReadReservoirs(
                Path.Combine(folder, "Reservoirs.csv"), network);

            ReadPipes(
                Path.Combine(folder, "Pipes.csv"), network);

            return network;
        }

        private static void ReadJunctions(
            string path, WaterNetwork network)
        {
            List<string[]> rows = ReadRows(path, 3);

            foreach (string[] f in rows)
            {
                network.Junctions.Add(new Junction
                {
                    Id = f[0].Trim(),
                    Elevation = ParseNumber(f[1], path, "Elevation"),
                    Demand = ParseNumber(f[2], path, "Demand")
                });
            }
        }

        private static void ReadReservoirs(
            string path, WaterNetwork network)
        {
            List<string[]> rows = ReadRows(path, 2);

            foreach (string[] f in rows)
            {
                network.Reservoirs.Add(new Reservoir
                {
                    Id = f[0].Trim(),
                    Head = ParseNumber(f[1], path, "Head")
                });
            }
        }

        private static void ReadPipes(
            string path, WaterNetwork network)
        {
            List<string[]> rows = ReadRows(path, 7);

            foreach (string[] f in rows)
            {
                network.Pipes.Add(new Pipe
                {
                    Id = f[0].Trim(),
                    StartNode = f[1].Trim(),
                    EndNode = f[2].Trim(),

                    // CSV columns:
                    // 0=Id, 1=StartNode, 2=EndNode,
                    // 3=Length, 4=Diameter, 5=Roughness,
                    // 6=MinorLoss

                    Length = ParseNumber(f[3], path, "Length"),
                    Diameter = ParseNumber(f[4], path, "Diameter"),
                    Roughness = ParseNumber(f[5], path, "Roughness"),
                    MinorLoss = ParseNumber(f[6], path, "MinorLoss")
                });
            }
        }

        private static List<string[]> ReadRows(
            string path, int expectedColumns)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException(
                    "CSV file not found: " + path, path);
            }

            string[] lines = File.ReadAllLines(path);

            if (lines.Length == 0)
            {
                throw new InvalidDataException(
                    "CSV file is empty: " + path);
            }

            List<string[]> rows = new List<string[]>();

            // The first line is the header.
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i]))
                    continue;

                string[] fields = lines[i].Split(',');

                if (fields.Length != expectedColumns)
                {
                    throw new InvalidDataException(
                        Path.GetFileName(path) +
                        ", line " + (i + 1) +
                        ": expected " + expectedColumns +
                        " columns but found " + fields.Length + ".");
                }

                rows.Add(fields);
            }

            return rows;
        }

        private static double ParseNumber(
            string value, string path, string column)
        {
            double result;

            if (!double.TryParse(
                value.Trim(),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out result))
            {
                throw new InvalidDataException(
                    "Invalid number in " + Path.GetFileName(path) +
                    ", column " + column +
                    ": '" + value + "'.");
            }

            if (double.IsNaN(result) || double.IsInfinity(result))
            {
                throw new InvalidDataException(
                    "Non-finite number in " + Path.GetFileName(path) +
                    ", column " + column + ".");
            }

            return result;
        }
    }

    public static class NetworkValidator
    {
        public static List<string> Validate(WaterNetwork network)
        {
            List<string> errors = new List<string>();

            Dictionary<string, string> nodes =
                new Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase);

            HashSet<string> pipeIds =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);

            if (network.Junctions.Count == 0)
                errors.Add("No junctions were loaded.");

            if (network.Reservoirs.Count == 0)
                errors.Add("No reservoirs were loaded.");

            if (network.Pipes.Count == 0)
                errors.Add("No pipes were loaded.");

            foreach (Junction j in network.Junctions)
            {
                RegisterNode(j.Id, "Junction", nodes, errors);

                if (!IsFinite(j.Elevation))
                    errors.Add("Invalid elevation at junction " + j.Id);

                if (!IsFinite(j.Demand) || j.Demand < 0)
                    errors.Add("Invalid demand at junction " + j.Id);
            }

            foreach (Reservoir r in network.Reservoirs)
            {
                RegisterNode(r.Id, "Reservoir", nodes, errors);

                if (!IsFinite(r.Head))
                    errors.Add("Invalid head at reservoir " + r.Id);
            }

            foreach (Pipe p in network.Pipes)
            {
                if (string.IsNullOrWhiteSpace(p.Id))
                {
                    errors.Add("A pipe has an empty ID.");
                }
                else
                {
                    if (!pipeIds.Add(p.Id))
                        errors.Add("Duplicate pipe ID: " + p.Id);

                    if (nodes.ContainsKey(p.Id))
                        errors.Add("Pipe ID conflicts with a node ID: " + p.Id);
                }

                if (string.IsNullOrWhiteSpace(p.StartNode) ||
                    !nodes.ContainsKey(p.StartNode))
                {
                    errors.Add(
                        "Pipe " + p.Id +
                        " has a missing start node: " + p.StartNode);
                }

                if (string.IsNullOrWhiteSpace(p.EndNode) ||
                    !nodes.ContainsKey(p.EndNode))
                {
                    errors.Add(
                        "Pipe " + p.Id +
                        " has a missing end node: " + p.EndNode);
                }

                if (!string.IsNullOrWhiteSpace(p.StartNode) &&
                    string.Equals(
                        p.StartNode, p.EndNode,
                        StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add(
                        "Pipe " + p.Id +
                        " has identical start and end nodes.");
                }

                if (!IsFinite(p.Length) || p.Length <= 0)
                    errors.Add("Invalid length for pipe " + p.Id);

                if (!IsFinite(p.Diameter) || p.Diameter <= 0)
                    errors.Add("Invalid diameter for pipe " + p.Id);

                if (!IsFinite(p.Roughness) || p.Roughness <= 0)
                    errors.Add("Invalid roughness for pipe " + p.Id);

                if (!IsFinite(p.MinorLoss) || p.MinorLoss < 0)
                    errors.Add("Invalid minor loss for pipe " + p.Id);
            }

            return errors;
        }

        private static void RegisterNode(
            string id,
            string type,
            Dictionary<string, string> nodes,
            List<string> errors)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                errors.Add("A " + type + " has an empty ID.");
                return;
            }

            if (nodes.ContainsKey(id))
            {
                errors.Add("Duplicate node ID: " + id);
            }
            else
            {
                nodes.Add(id, type);
            }
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) &&
                   !double.IsInfinity(value);
        }
    }

    public static class EpanetFileWriter
    {
        public static void Write(
            WaterNetwork network, string path)
        {
            using (StreamWriter writer = new StreamWriter(path, false))
            {
                writer.WriteLine("[TITLE]");
                writer.WriteLine("WaterCAD Network Generator - NEW V1");
                writer.WriteLine();

                writer.WriteLine("[JUNCTIONS]");
                writer.WriteLine(";ID\tElevation\tDemand\tPattern");

                foreach (Junction j in network.Junctions)
                {
                    writer.WriteLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0}\t{1}\t{2}",
                        j.Id, j.Elevation, j.Demand));
                }

                writer.WriteLine();
                writer.WriteLine("[RESERVOIRS]");
                writer.WriteLine(";ID\tHead\tPattern");

                foreach (Reservoir r in network.Reservoirs)
                {
                    writer.WriteLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0}\t{1}",
                        r.Id, r.Head));
                }

                writer.WriteLine();
                writer.WriteLine("[PIPES]");
                writer.WriteLine(
                    ";ID\tNode1\tNode2\tLength\tDiameter\tRoughness\tMinorLoss\tStatus");

                foreach (Pipe p in network.Pipes)
                {
                    // Values are written directly from the CSV data
                    // parsed into the Pipe object.
                    writer.WriteLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\tOPEN",
                        p.Id,
                        p.StartNode,
                        p.EndNode,
                        p.Length,
                        p.Diameter,
                        p.Roughness,
                        p.MinorLoss));
                }

                writer.WriteLine();
                writer.WriteLine("[OPTIONS]");
                writer.WriteLine("UNITS LPS");
                writer.WriteLine("HEADLOSS H-W");

                writer.WriteLine();
                writer.WriteLine("[COORDINATES]");
                writer.WriteLine(";Node\tX-Coord\tY-Coord");

                int x = 0;

                foreach (Reservoir r in network.Reservoirs)
                {
                    writer.WriteLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0}\t{1}\t0", r.Id, x));

                    x += 500;
                }

                x = 500;

                foreach (Junction j in network.Junctions)
                {
                    writer.WriteLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0}\t{1}\t0", j.Id, x));

                    x += 300;
                }

                writer.WriteLine();
                writer.WriteLine("[END]");
                WaterCADController.Run();
                AutoHotkeyTest.Run();
            }
        }
    }
}
