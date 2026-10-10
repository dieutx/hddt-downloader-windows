using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("HDDT Downloader")]
[assembly: AssemblyProduct("HDDT Downloader")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

internal static class Launcher
{
    [STAThread]
    private static int Main()
    {
        try
        {
            string root = AppDomain.CurrentDomain.BaseDirectory;
            string script = Path.Combine(root, "Start-HddtGui.ps1");
            if (!File.Exists(script))
                throw new FileNotFoundException("Hay giai nen toan bo goi app truoc khi chay.", script);
            string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
                @"WindowsPowerShell\v1.0\powershell.exe");
            ProcessStartInfo info = new ProcessStartInfo(powershell);
            info.Arguments = "-NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File \"" + script + "\"";
            info.WorkingDirectory = root;
            info.UseShellExecute = false;
            info.CreateNoWindow = true;
            info.RedirectStandardError = true;
            using (Process process = Process.Start(info))
            {
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();
                if (process.ExitCode != 0)
                    MessageBox.Show(String.IsNullOrWhiteSpace(error) ? "Khong khoi dong duoc giao dien WPF." : error,
                        "HDDT Downloader", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return process.ExitCode;
            }
        }
        catch (Exception error)
        {
            MessageBox.Show(error.Message, "HDDT Downloader", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }
}
