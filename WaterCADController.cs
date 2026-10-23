
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace Hadi.Excel
{
    /// <summary>
    /// Starts WaterCAD if necessary and attempts to restore
    /// and activate its main window.
    /// </summary>
    internal static class WaterCADController
    {
        private const string WaterCADPath =
            @"C:\Program Files (x86)\Bentley\WaterCAD\x64\WaterCAD.exe";

        private const string ProcessName = "WaterCAD";

        private const int SW_RESTORE = 9;
        private const int GW_OWNER = 4;

        private const int StartupTimeoutMilliseconds = 30000;
        private const int WindowTimeoutMilliseconds = 15000;
        private const int PollIntervalMilliseconds = 250;

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(
            EnumWindowsProc callback,
            IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(
            IntPtr hWnd,
            out uint processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(
            IntPtr hWnd,
            int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(
            IntPtr hWnd,
            int uCmd);

        [DllImport(
            "user32.dll",
            CharSet = CharSet.Unicode,
            SetLastError = true)]
        private static extern int GetWindowText(
            IntPtr hWnd,
            StringBuilder text,
            int maxCount);

        private delegate bool EnumWindowsProc(
            IntPtr hWnd,
            IntPtr lParam);

        /// <summary>
        /// Ensures WaterCAD is running, then attempts to activate it.
        /// This method does not import an INP file.
        /// </summary>
        public static WaterCADActivationResult EnsureRunningAndActivate()
        {
            Process process = null;
            bool startedNewProcess = false;

            try
            {
                // Prefer an existing WaterCAD instance.
                process = FindWaterCADProcess();

                if (process == null)
                {
                    if (!File.Exists(WaterCADPath))
                    {
                        return WaterCADActivationResult.Failure(
                            "WaterCAD.exe was not found at:\n" +
                            WaterCADPath);
                    }

                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = WaterCADPath,
                            UseShellExecute = true
                        });

                        startedNewProcess = true;
                    }
                    catch (Exception ex)
                    {
                        return WaterCADActivationResult.Failure(
                            "WaterCAD could not be started.\n\n" +
                            ex.Message);
                    }

                    // Wait for the application process to appear.
                    process = WaitForWaterCADProcess(
                        StartupTimeoutMilliseconds);

                    if (process == null)
                    {
                        return WaterCADActivationResult.Failure(
                            "WaterCAD was started, but its process " +
                            "could not be detected within 30 seconds.");
                    }
                }

                int processId = process.Id;

                // Wait for the main application window.
                IntPtr windowHandle = IntPtr.Zero;
                string windowTitle = string.Empty;

                Stopwatch stopwatch = Stopwatch.StartNew();

                while (stopwatch.ElapsedMilliseconds <
                       WindowTimeoutMilliseconds)
                {
                    // Prefer a WaterCAD window belonging to the
                    // selected process.
                    windowHandle = FindMainWindow(
                        processId,
                        out windowTitle);

                    if (windowHandle != IntPtr.Zero)
                        break;

                    Thread.Sleep(PollIntervalMilliseconds);
                }

                if (windowHandle == IntPtr.Zero)
                {
                    return WaterCADActivationResult.Failure(
                        "WaterCAD is running, but its main window " +
                        "could not be found.\n\n" +
                        "The application may still be loading.");
                }

                bool wasMinimized = IsIconic(windowHandle);

                // Restore the window if minimized. Calling this
                // for a non-minimized window is also harmless.
                ShowWindowAsync(windowHandle, SW_RESTORE);

                // Allow Windows time to process the restore request.
                Thread.Sleep(300);

                // Ask Windows to bring WaterCAD to the foreground.
                bool activationAccepted =
                    SetForegroundWindow(windowHandle);

                // Check which window actually has foreground focus.
                Thread.Sleep(300);

                bool isForeground =
                    GetForegroundWindow() == windowHandle;

                string message;

                if (isForeground)
                {
                    message = startedNewProcess
                        ? "WaterCAD was started and activated successfully."
                        : "The existing WaterCAD window was activated successfully.";
                }
                else if (wasMinimized || IsIconic(windowHandle))
                {
                    message =
                        "WaterCAD was found, but Windows did not " +
                        "confirm that its window was restored.";
                }
                else
                {
                    message =
                        "WaterCAD is running and its window was found, " +
                        "but it did not become the foreground window. " +
                        "Windows may be restricting foreground activation.";
                }

                return new WaterCADActivationResult
                {
                    Success = true,
                    StartedNewProcess = startedNewProcess,
                    ActivationAccepted = activationAccepted,
                    IsForeground = isForeground,
                    ProcessId = processId,
                    WindowTitle = windowTitle,
                    Message = message
                };
            }
            catch (Exception ex)
            {
                return WaterCADActivationResult.Failure(
                    "An unexpected error occurred while activating WaterCAD.\n\n" +
                    ex.Message);
            }
            finally
            {
                process?.Dispose();
            }
        }


        private static Process FindWaterCADProcess()
        {
            Process[] processes =
                Process.GetProcessesByName(ProcessName);

            Process selected = null;

            try
            {
                IntPtr foreground = GetForegroundWindow();

                GetWindowThreadProcessId(
                    foreground,
                    out uint foregroundProcessId);

                // Prefer the WaterCAD process currently in the foreground.
                foreach (Process candidate in processes)
                {
                    try
                    {
                        if (!candidate.HasExited &&
                            candidate.Id == (int)foregroundProcessId)
                        {
                            selected = candidate;
                            break;
                        }
                    }
                    catch
                    {
                        // The process may have exited while being inspected.
                    }
                }

                // Otherwise prefer a process with a main window.
                if (selected == null)
                {
                    foreach (Process candidate in processes)
                    {
                        try
                        {
                            if (!candidate.HasExited &&
                                candidate.MainWindowHandle != IntPtr.Zero)
                            {
                                selected = candidate;
                                break;
                            }
                        }
                        catch
                        {
                            // Continue checking.
                        }
                    }
                }

                // Last resort: use any process that is still running.
                if (selected == null)
                {
                    foreach (Process candidate in processes)
                    {
                        try
                        {
                            if (!candidate.HasExited)
                            {
                                selected = candidate;
                                break;
                            }
                        }
                        catch
                        {
                            // Continue checking.
                        }
                    }
                }

                return selected;
            }
            finally
            {
                // Dispose every process object except the one returned
                // to the caller.
                foreach (Process candidate in processes)
                {
                    if (!ReferenceEquals(candidate, selected))
                        candidate.Dispose();
                }
            }
        }


        private static Process WaitForWaterCADProcess(int timeoutMilliseconds)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            while (stopwatch.ElapsedMilliseconds < timeoutMilliseconds)
            {
                Process process = FindWaterCADProcess();

                if (process != null)
                    return process;

                Thread.Sleep(PollIntervalMilliseconds);
            }

            return null;
        }

        private static IntPtr FindMainWindow(
            int processId,
            out string windowTitle)
        {
            IntPtr bestHandle = IntPtr.Zero;
            string bestTitle = string.Empty;

            EnumWindows((hWnd, lParam) =>
            {
                GetWindowThreadProcessId(
                    hWnd,
                    out uint ownerProcessId);

                if (ownerProcessId != (uint)processId)
                    return true;

                if (!IsWindowVisible(hWnd))
                    return true;

                // Avoid selecting owned dialogs when a top-level
                // application window is available.
                if (GetWindow(hWnd, GW_OWNER) != IntPtr.Zero)
                    return true;

                StringBuilder title = new StringBuilder(512);

                GetWindowText(
                    hWnd,
                    title,
                    title.Capacity);

                string currentTitle = title.ToString().Trim();

                if (currentTitle.Length == 0)
                    return true;

                bestHandle = hWnd;
                bestTitle = currentTitle;

                // Stop once a suitable top-level window is found.
                return false;
            }, IntPtr.Zero);

            windowTitle = bestTitle;
            return bestHandle;
        }
    }

    /// <summary>
    /// Describes the outcome of a WaterCAD activation attempt.
    /// </summary>
    internal sealed class WaterCADActivationResult
    {
        public bool Success { get; set; }

        public bool StartedNewProcess { get; set; }

        public bool ActivationAccepted { get; set; }

        public bool IsForeground { get; set; }

        public int ProcessId { get; set; }

        public string WindowTitle { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public static WaterCADActivationResult Failure(string message)
        {
            return new WaterCADActivationResult
            {
                Success = false,
                Message = message
            };
        }
    }
}
