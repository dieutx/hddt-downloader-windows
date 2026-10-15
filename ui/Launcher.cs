using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

[assembly: AssemblyTitle("HDDT Downloader")]
[assembly: AssemblyProduct("HDDT Downloader")]

internal static class Launcher
{
    private static string runtimeDirectory;

    private static string Quote(string value)
    {
        // Cac doi so la duong dan file/thu muc, khong chap nhan dau ngoac kep.
        if (value.IndexOf('"') >= 0) throw new ArgumentException("Duong dan khong hop le.");
        int trailing = value.Length - value.TrimEnd('\\').Length;
        return "\"" + value + new String('\\', trailing) + "\"";
    }

    private static string ExtractPayload(Stream payload)
    {
        string root = Path.Combine(Path.GetTempPath(), "HDDT-Downloader", "run-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        runtimeDirectory = root;
        string prefix = Path.GetFullPath(root) + Path.DirectorySeparatorChar;
        using (ZipArchive archive = new ZipArchive(payload, ZipArchiveMode.Read, false))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                string destination = Path.GetFullPath(Path.Combine(root, entry.FullName.Replace('/', Path.DirectorySeparatorChar)));
                if (!destination.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Duong dan trong goi chuong trinh khong hop le.");
                if (String.IsNullOrEmpty(entry.Name)) { Directory.CreateDirectory(destination); continue; }
                Directory.CreateDirectory(Path.GetDirectoryName(destination));
                using (Stream source = entry.Open())
                using (FileStream target = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    source.CopyTo(target);
            }
        }
        return root;
    }

    private static void CleanRuntime()
    {
        if (runtimeDirectory == null) return;
        // Chi xoa thu muc ngau nhien do chinh lan chay nay tao trong TEMP.
        string expectedParent = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "HDDT-Downloader"));
        if (!String.Equals(Path.GetDirectoryName(runtimeDirectory), expectedParent, StringComparison.OrdinalIgnoreCase)) return;
        try { Directory.Delete(runtimeDirectory, true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        bool smokeTest = args.Length > 0 && args[0] == "--smoke-test";
        string testRoot = null;
        string report = null;
        try
        {
            if (args.Length > 0 && (!smokeTest || args.Length < 2 || args.Length > 3))
                throw new ArgumentException("Tham so launcher khong hop le.");
            if (smokeTest)
            {
                testRoot = Path.GetFullPath(args[1]);
                Directory.CreateDirectory(testRoot);
                report = Path.Combine(testRoot, "smoke-report.json");
            }
            string root = AppDomain.CurrentDomain.BaseDirectory;
            bool embedded = false;
            Stream payload = Assembly.GetExecutingAssembly().GetManifestResourceStream("Hddt.Payload.zip");
            if (payload != null)
            {
                embedded = true;
                using (payload) { root = ExtractPayload(payload); }
            }
            string script = Path.Combine(root, "Start-HddtGui.ps1");
            if (!File.Exists(script))
                throw new FileNotFoundException("Khong tim thay thanh phan chuong trinh. Hay tai lai ban EXE day du.", script);
            string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
                Path.Combine("WindowsPowerShell", Path.Combine("v1.0", "powershell.exe")));
            ProcessStartInfo info = new ProcessStartInfo(powershell);
            info.Arguments = "-NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File " + Quote(script);
            if (embedded)
            {
                string data = smokeTest ? Path.Combine(testRoot, "data") : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HDDT-Downloader");
                string output = smokeTest ? Path.Combine(testRoot, "output") : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "HDDT-Downloader");
                if (!Path.IsPathRooted(data) || !Path.IsPathRooted(output)) throw new IOException("Khong xac dinh duoc thu muc du lieu nguoi dung.");
                info.Arguments += " -DataDirectory " + Quote(data) + " -DefaultOutputDirectory " + Quote(output);
            }
            if (smokeTest)
            {
                info.Arguments += " -SmokeTest -SmokeReportPath " + Quote(report);
                if (args.Length == 3) info.Arguments += " -SmokeXmlDirectory " + Quote(Path.GetFullPath(args[2]));
            }
            info.WorkingDirectory = root;
            info.UseShellExecute = false;
            info.CreateNoWindow = true;
            info.RedirectStandardError = true;
            using (Process process = Process.Start(info))
            {
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    if (smokeTest) File.WriteAllText(Path.Combine(testRoot, "launcher-error.txt"), error, new UTF8Encoding(false));
                    else MessageBox.Show(String.IsNullOrWhiteSpace(error) ? "Khong khoi dong duoc giao dien WPF." : error,
                        "HDDT Downloader", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return process.ExitCode;
            }
        }
        catch (Exception error)
        {
            if (smokeTest && testRoot != null) File.WriteAllText(Path.Combine(testRoot, "launcher-error.txt"), error.ToString(), new UTF8Encoding(false));
            else MessageBox.Show(error.Message, "HDDT Downloader", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
        finally { CleanRuntime(); }
    }
}
