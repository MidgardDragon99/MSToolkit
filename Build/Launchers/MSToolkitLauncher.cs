// MSToolkit launcher - compiled to MSToolkit.EXE by Build\Build-All.ps1 with the
// csc.exe that ships with Windows (.NET Framework 4.x, C# 5 only).
//
// It embeds nothing. It starts Launch-MSToolkit.bat from the folder this EXE is
// in (the Program Files install folder), exactly as Launcher.ps1 does. That .bat
// asks for the admin username in its console, so the console must stay visible:
// UseShellExecute = false and no CreateNoWindow. No username or domain is passed
// or added here.

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("MSToolkit")]
[assembly: AssemblyProduct("MSToolkit")]
[assembly: AssemblyDescription("Starts MSToolkit from its install folder")]
// AssemblyVersion, AssemblyFileVersion and AssemblyInformationalVersion come
// from the VERSION file in the repo root: Build-All.ps1 generates them at
// compile time. The installer reads the product version back from this EXE
// and writes it to Apps & Features as DisplayVersion.

namespace MSToolkitLauncher
{
    internal static class Program
    {
        private const string Caption = "MSToolkit";
        private const string LaunchFile = "Launch-MSToolkit.bat";

        [STAThread]
        private static int Main()
        {
            string installDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
            string launchBat = Path.Combine(installDir, LaunchFile);

            if (!File.Exists(launchBat))
            {
                ShowError("MSToolkit is not installed correctly: " + launchBat +
                    " was not found.\r\n\r\nRun MSToolkit-Install.exe again.");
                return 1;
            }

            string cmdExe = Path.Combine(GetSystemFolder(), "cmd.exe");

            // /c ""path"" - cmd strips the outer pair and keeps the inner quotes,
            // so a path with spaces or parentheses (Program Files (x86)) survives.
            ProcessStartInfo startInfo = new ProcessStartInfo(cmdExe, "/c \"\"" + launchBat + "\"\"");
            startInfo.UseShellExecute = false;
            startInfo.WorkingDirectory = installDir;

            try
            {
                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                ShowError("MSToolkit could not start " + launchBat + ".\r\n\r\n" + ex.Message);
                return 1;
            }

            return 0;
        }

        // The 64-bit System32. A 32-bit process on 64-bit Windows sees SysWOW64 as
        // System32, so it has to ask for Sysnative instead.
        private static string GetSystemFolder()
        {
            if (Environment.Is64BitOperatingSystem && !Environment.Is64BitProcess)
            {
                string sysnative = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Sysnative");
                if (Directory.Exists(sysnative))
                {
                    return sysnative;
                }
            }

            return Environment.SystemDirectory;
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
