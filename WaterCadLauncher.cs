
using System;
using System.Diagnostics;
using System.IO;

namespace WaterCadTest48
{
    public static class WaterCadLauncher
    {
        public static void Launch()
        {
            string waterCadExe =
                @"C:\Program Files (x86)\Bentley\WaterCAD\x64\WaterCAD.exe";

            string projectPath =
                Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.Desktop),
                    "WaterCAD_TestNetwork.wtg");

            if (!File.Exists(waterCadExe))
            {
                Console.WriteLine("WaterCAD.exe was not found.");
                return;
            }

            if (!File.Exists(projectPath))
            {
                Console.WriteLine("The WaterCAD project was not found:");
                Console.WriteLine(projectPath);
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = waterCadExe,
                Arguments = "\"" + projectPath + "\"",
                UseShellExecute = true
            };

            Process.Start(startInfo);

            Console.WriteLine("Requested WaterCAD to open:");
            Console.WriteLine(projectPath);
        }
    }
}
