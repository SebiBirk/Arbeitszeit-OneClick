using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            string baseDir = GetBaseDir(args);
            string logPath = Path.Combine(baseDir, "ArbeitszeitAnzeige.log");
            string scriptPath = Path.Combine(baseDir, "ArbeitszeitAnzeige.ps1");

            if (!File.Exists(scriptPath))
            {
                WriteLog(logPath, "Anzeige-Skript nicht gefunden: " + scriptPath);
                MessageBox.Show(
                    "Anzeige-Skript nicht gefunden:\r\n" + scriptPath,
                    "Arbeitszeit",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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

                        MessageBox.Show(
                            "Die Anzeige konnte nicht vollständig gestartet werden.",
                            "Arbeitszeit",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return 1;
                    }
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            string fallbackDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
            WriteLog(Path.Combine(fallbackDir, "ArbeitszeitAnzeige.log"), ex.ToString());

            MessageBox.Show(
                ex.Message,
                "Arbeitszeit",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
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
