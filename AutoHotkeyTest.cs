
using System;
using System.Diagnostics;
using System.IO;

namespace WaterCadTest48
{
    internal class AutoHotkeyTest
    {
        public static void Run()
        {
            string autoHotkeyPath =
                @"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe";

            string scriptPath =
                @"C:\Users\Hadi Wehbi\Desktop\Coding\01_Project\WaterCadTest48\WaterCadTest48\TestWaterCADExport.ahk";

            if (!File.Exists(autoHotkeyPath))
            {
                Console.WriteLine("AutoHotkey executable not found.");
                return;
            }

            if (!File.Exists(scriptPath))
            {
                Console.WriteLine("AutoHotkey script not found:");
                Console.WriteLine(scriptPath);
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = autoHotkeyPath,
                Arguments = "\"" + scriptPath + "\"",
                UseShellExecute = true
            };

            Process process = Process.Start(startInfo);

            if (process == null)
            {
                Console.WriteLine("Could not start AutoHotkey.");
                return;
            }

            Console.WriteLine("AutoHotkey export script started.");
            process.WaitForExit();

            Console.WriteLine(
                "AutoHotkey script exited with code: "
                + process.ExitCode);
            process.Dispose();
        }
    }
}
