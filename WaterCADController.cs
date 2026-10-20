
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace WaterCadTest48
{
    internal class WaterCADController
    {
        private const string WaterCADPath =
            @"C:\Program Files (x86)\Bentley\WaterCAD\x64\WaterCAD.exe";

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(
            EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(
            IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(
            IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(
            IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(
            IntPtr hWnd, StringBuilder text, int maxCount);

        private delegate bool EnumWindowsProc(
            IntPtr hWnd, IntPtr lParam);

        private class WindowInfo
        {
            public IntPtr Handle;
            public string Title = "";
        }

        public static void Run()
        {
            Console.WriteLine();
            Console.WriteLine("=== WaterCAD Controller ===");

            Process waterCAD = FindWaterCAD();

            if (waterCAD == null)
            {
                Console.WriteLine("Starting WaterCAD...");

                if (!File.Exists(WaterCADPath))
                {
                    Console.WriteLine(
                        "WaterCAD.exe not found: " + WaterCADPath);
                    return;
                }

                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = WaterCADPath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return;
                }

                for (int i = 0; i < 60; i++)
                {
                    Thread.Sleep(500);
                    waterCAD = FindWaterCAD();

                    if (waterCAD != null)
                        break;
                }
            }
            else
            {
                Console.WriteLine("Existing WaterCAD process found.");
            }

            if (waterCAD == null)
            {
                Console.WriteLine("Could not find WaterCAD process.");
                return;
            }

            WindowInfo window = FindWaterCADWindow(waterCAD.Id);

            for (int i = 0; window == null && i < 20; i++)
            {
                Thread.Sleep(500);
                window = FindWaterCADWindow(waterCAD.Id);
            }

            if (window == null)
            {
                Console.WriteLine(
                    "No titled WaterCAD window was found.");
                return;
            }

            Console.WriteLine("Window title: " + window.Title);
            Console.WriteLine(
                "Was minimized: " + IsIconic(window.Handle));

            // Explicitly request restoration.
            ShowWindowAsync(window.Handle, SW_RESTORE);
            ShowWindow(window.Handle, SW_RESTORE);
            Thread.Sleep(1000);

            // Also request that the window move to the foreground.
            BringWindowToTop(window.Handle);
            bool activated = SetForegroundWindow(window.Handle);

            Thread.Sleep(500);

            Console.WriteLine(
                "Minimized after restore: " + IsIconic(window.Handle));

            Console.WriteLine(
                activated
                    ? "Windows accepted the activation request."
                    : "Windows did not confirm foreground activation.");

            Console.WriteLine("Controller finished.");
        }

        private static Process FindWaterCAD()
        {
            Process[] processes =
                Process.GetProcessesByName("WaterCAD");

            foreach (Process process in processes)
            {
                try
                {
                    if (!process.HasExited)
                        return process;
                }
                catch
                {
                    // Process may have exited.
                }
            }

            return null;
        }

        private static WindowInfo FindWaterCADWindow(int processId)
        {
            WindowInfo best = null;

            EnumWindows((hWnd, lParam) =>
            {
                GetWindowThreadProcessId(
                    hWnd, out uint windowProcessId);

                if (windowProcessId != (uint)processId)
                    return true;

                StringBuilder title = new StringBuilder(512);
                GetWindowText(hWnd, title, title.Capacity);

                string windowTitle = title.ToString().Trim();

                if (IsWindowVisible(hWnd) &&
                    windowTitle.Length > 0)
                {
                    // Prefer a titled main window.
                    best = new WindowInfo
                    {
                        Handle = hWnd,
                        Title = windowTitle
                    };

                    // Stop at the first titled window for this process.
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return best;
        }
    }
}
