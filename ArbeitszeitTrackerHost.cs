using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        string baseDir = GetBaseDir(args);
        string logPath = Path.Combine(baseDir, "ArbeitszeitTrackerHost.log");

        try
        {
            string scriptPath = Path.Combine(baseDir, "Arbeitszeit.ps1");

            if (!File.Exists(scriptPath))
            {
                WriteLog(logPath, "Tracker-Skript nicht gefunden: " + scriptPath);
                return 2;
            }

            Environment.CurrentDirectory = baseDir;

            InitialSessionState sessionState = InitialSessionState.CreateDefault();
            sessionState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.RemoteSigned;

            using (Runspace runspace = RunspaceFactory.CreateRunspace(sessionState))
            {
                runspace.ApartmentState = ApartmentState.STA;
                runspace.Open();

                using (PowerShell powerShell = PowerShell.Create())
                {
                    powerShell.Runspace = runspace;
                    powerShell.AddCommand(scriptPath).AddParameter("BaseDir", baseDir);
                    powerShell.Invoke();

                    if (powerShell.HadErrors)
                    {
                        foreach (ErrorRecord error in powerShell.Streams.Error)
                        {
                            WriteLog(logPath, error.ToString());
                        }

                        return 1;
                    }
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            WriteLog(logPath, ex.ToString());
            return 1;
        }
    }

    private static string GetBaseDir(string[] args)
    {
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (string.Equals(args[i], "-BaseDir", StringComparison.OrdinalIgnoreCase))
            {
                return Path.GetFullPath(args[i + 1]);
            }
        }

        return AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
    }

    private static void WriteLog(string path, string message)
    {
        try
        {
            File.AppendAllText(
                path,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + message + Environment.NewLine);
        }
        catch
        {
        }
    }
}
