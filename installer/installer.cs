using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace AutoCADBalletInstaller
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Determine if running as uninstaller based on executable name
            string exeName = Path.GetFileName(Assembly.GetExecutingAssembly().Location).ToLower();
            bool isUninstaller = exeName.Contains("uninstall") || (args.Length > 0 && args[0] == "/uninstall");

            if (isUninstaller)
            {
                new BalletUninstaller().Run();
            }
            else
            {
                new BalletInstaller().Run();
            }
        }
    }

    public class BalletInstaller
    {
        private const string PRODUCT_NAME = "AutoCAD Ballet";
        private const string UNINSTALL_KEY = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AutoCADBallet";

        // Simple acad.lsp loader that just adds trusted path and loads startup-loader.lsp
        private const string ACAD_LSP_LOADER = @";; autocad-ballet
;; add to trusted paths and load startup loader
(setq ballet-installer-path (strcat (getenv ""APPDATA"") ""\\autocad-ballet\\installer""))
(setq tp (getvar ""TRUSTEDPATHS""))
;; Check for exact path by adding semicolons to both strings
(if (not (vl-string-search (strcat "";"" ballet-installer-path "";"") 
                           (strcat "";"" tp "";"")))
  (setvar ""TRUSTEDPATHS"" (strcat tp "";"" ballet-installer-path)))
(load (strcat ballet-installer-path ""\\startup-loader.lsp""))
";

        private readonly List<AutoCADInstallation> installations = new List<AutoCADInstallation>();

        public void Run()
        {
            try
            {
                if (IsInstalled())
                {
                    if (MessageBox.Show("AutoCAD Ballet is already installed. Reinstall?",
                        "AutoCAD Ballet", MessageBoxButtons.YesNo) != DialogResult.Yes)
                    {
                        return;
                    }
                }

                DetectAutoCADInstallations();

                if (installations.Count == 0)
                {
                    MessageBox.Show("No AutoCAD installations found.",
                        "AutoCAD Ballet", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string targetDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "autocad-ballet");
                Directory.CreateDirectory(targetDir);

                // Create temp folder for runtime dll copies
                string tempDir = Path.Combine(targetDir, "commands", "bin", "temp");
                Directory.CreateDirectory(tempDir);

                // Extract only DLLs for installed AutoCAD versions
                ExtractVersionSpecificDlls(targetDir);

                // Extract startup-loader.lsp to installer folder
                ExtractStartupLoader(targetDir);

                // Extract all LISP files (including startup-loader.lsp)
                ExtractLispFiles(targetDir);

                CopyInstallerForUninstall(targetDir);

                var modifiedInstallations = new List<AutoCADInstallation>();
                foreach (var installation in installations)
                {
                    bool modified = ModifyAcadLsp(installation);
                    if (modified)
                    {
                        modifiedInstallations.Add(installation);
                    }
                }

                RegisterUninstaller(targetDir);

                ShowSuccessDialog(modifiedInstallations);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Installation error: {ex.Message}",
                    "AutoCAD Ballet Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowSuccessDialog(List<AutoCADInstallation> modifiedInstallations)
        {
            var form = new Form
            {
                Text = "AutoCAD Ballet - Installation Complete",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.White
            };

            Image logo = null;

            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var logoResourceName = assembly.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith(".png"));

                if (logoResourceName != null)
                {
                    using (var stream = assembly.GetManifestResourceStream(logoResourceName))
                    {
                        if (stream != null)
                        {
                            logo = Image.FromStream(stream);
                        }
                    }
                }
            }
            catch { }

            int currentY = 30;

            if (logo != null)
            {
                var pictureBox = new PictureBox
                {
                    Image = logo,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Location = new Point(175, currentY),
                    Size = new Size(200, 150)
                };
                form.Controls.Add(pictureBox);
                currentY += 160;
            }
            else
            {
                var titleLabel = new Label
                {
                    Text = "AutoCAD Ballet",
                    Font = new Font("Segoe UI", 24, FontStyle.Bold),
                    ForeColor = Color.FromArgb(50, 50, 50),
                    Location = new Point(0, currentY),
                    Size = new Size(550, 40),
                    TextAlign = ContentAlignment.MiddleCenter
                };
                form.Controls.Add(titleLabel);
                currentY += 60;
            }

            string messageText = "AutoCAD Ballet has been successfully added to:\n\n";

            if (modifiedInstallations.Count > 0)
            {
                foreach (var installation in modifiedInstallations)
                {
                    messageText += $"• {installation.Name} ({installation.Year})\n";
                }
            }
            else
            {
                messageText += "The plugin files have been installed but no AutoCAD installations were modified.\n";
                messageText += "AutoCAD LT versions prior to 2024 are not supported for automatic loading.";
            }

            using (var g = form.CreateGraphics())
            {
                var textSize = g.MeasureString(messageText, new Font("Segoe UI", 10), 450);
                var textHeight = (int)Math.Ceiling(textSize.Height);

                var messageLabel = new Label
                {
                    Text = messageText,
                    Font = new Font("Segoe UI", 10),
                    ForeColor = Color.FromArgb(60, 60, 60),
                    Location = new Point(50, currentY),
                    Size = new Size(450, textHeight + 10),
                    TextAlign = ContentAlignment.TopCenter
                };
                form.Controls.Add(messageLabel);

                currentY += textHeight + 30;
            }

            var btnOK = new Button
            {
                Text = "OK",
                Font = new Font("Segoe UI", 10),
                Size = new Size(100, 30),
                Location = new Point(230, currentY),
                FlatStyle = FlatStyle.System,
                DialogResult = DialogResult.OK
            };
            form.Controls.Add(btnOK);
            form.AcceptButton = btnOK;

            currentY += 50;
            form.ClientSize = new Size(550, currentY);

            form.ShowDialog();
        }

        private bool IsInstalled()
        {
            try
            {
                return Registry.CurrentUser.OpenSubKey(UNINSTALL_KEY) != null;
            }
            catch
            {
                return false;
            }
        }

        private void DetectAutoCADInstallations()
        {
            string[] regPaths = {
                @"SOFTWARE\Autodesk\AutoCAD"
            };

            foreach (string regPath in regPaths)
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(regPath))
                    {
                        if (key == null) continue;

                        foreach (string versionKey in key.GetSubKeyNames())
                        {
                            using (var vKey = key.OpenSubKey(versionKey))
                            {
                                if (vKey == null) continue;

                                foreach (string releaseKey in vKey.GetSubKeyNames())
                                {
                                    using (var rKey = vKey.OpenSubKey(releaseKey))
                                    {
                                        if (rKey == null) continue;

                                        var roamableRoot = rKey.GetValue("RoamableRootFolder")?.ToString();

                                        if (string.IsNullOrEmpty(roamableRoot)) continue;

                                        var productName = rKey.GetValue("ProductName")?.ToString();

                                        if (string.IsNullOrEmpty(productName))
                                        {
                                            productName = GetProductNameFromHKLM(versionKey, releaseKey) ?? "AutoCAD";
                                        }

                                        string supportPath = Path.Combine(roamableRoot, "enu", "Support");
                                        if (!Directory.Exists(supportPath))
                                        {
                                            supportPath = Path.Combine(roamableRoot, "Support");
                                        }

                                        if (Directory.Exists(supportPath))
                                        {
                                            var installation = new AutoCADInstallation
                                            {
                                                Name = productName,
                                                Version = ParseVersion(versionKey),
                                                VersionKey = versionKey,
                                                ReleaseKey = releaseKey,
                                                IsLT = productName.Contains("LT"),
                                                SupportPath = supportPath
                                            };

                                            // Determine year from version
                                            installation.Year = GetYearFromVersion(versionKey);

                                            installations.Add(installation);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch { }
            }
        }

        private string GetYearFromVersion(string versionKey)
        {
            // Map version keys to years
            if (versionKey.StartsWith("R21")) return "2017";
            if (versionKey.StartsWith("R22")) return "2018";
            if (versionKey.StartsWith("R23.0")) return "2019";
            if (versionKey.StartsWith("R23.1")) return "2020";
            if (versionKey.StartsWith("R24.0")) return "2021";
            if (versionKey.StartsWith("R24.1")) return "2022";
            if (versionKey.StartsWith("R24.2")) return "2023";
            if (versionKey.StartsWith("R24.3")) return "2024";
            if (versionKey.StartsWith("R25.0")) return "2025";
            if (versionKey.StartsWith("R25.1")) return "2026";

            // Fallback
            int version = ParseVersion(versionKey);
            if (version >= 2017 && version <= 2026)
                return version.ToString();

            return "unknown";
        }

        private void ExtractVersionSpecificDlls(string targetDir)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();

            // Get unique years from detected installations
            var installedYears = installations
                .Select(i => i.Year)
                .Where(y => y != "unknown")
                .Distinct()
                .ToList();

            foreach (string year in installedYears)
            {
                // Create year-specific directory
                string yearDir = Path.Combine(targetDir, "commands", "bin", year);
                Directory.CreateDirectory(yearDir);

                // Look for all DLL resources for this year (autocad-ballet.dll and Mono.Cecil*.dll)
                var yearDllResources = resourceNames.Where(r =>
                    r.Contains($"_{year}.") && r.EndsWith(".dll")).ToList();

                foreach (var dllResource in yearDllResources)
                {
                    // Extract filename from resource name
                    // Resource names are like: installer._2026.autocad-ballet.dll or installer._2026.Mono.Cecil.dll
                    string fileName = dllResource.Substring(dllResource.LastIndexOf('.', dllResource.Length - 5) + 1);

                    // Handle pattern like: installer._2026.Mono.Cecil.dll
                    // We want to extract just "Mono.Cecil.dll" or "autocad-ballet.dll"
                    int yearDotIndex = dllResource.IndexOf($"_{year}.");
                    if (yearDotIndex >= 0)
                    {
                        fileName = dllResource.Substring(yearDotIndex + year.Length + 2); // +2 for "_{year}."
                    }

                    string dllPath = Path.Combine(yearDir, fileName);
                    ExtractResource(dllResource, dllPath);
                }
            }
        }

        private void ExtractStartupLoader(string targetDir)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();

            string installerDir = Path.Combine(targetDir, "installer");
            Directory.CreateDirectory(installerDir);

            // Look for startup-loader.lsp in embedded resources
            var startupLoaderResource = resourceNames.FirstOrDefault(r => r.EndsWith("startup-loader.lsp"));

            if (startupLoaderResource != null)
            {
                string startupLoaderPath = Path.Combine(installerDir, "startup-loader.lsp");
                ExtractResource(startupLoaderResource, startupLoaderPath);
            }
            else
            {
                throw new Exception("startup-loader.lsp not found in embedded resources");
            }
        }

        private void ExtractLispFiles(string targetDir)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();

            string lispDir = Path.Combine(targetDir, "commands");
            Directory.CreateDirectory(lispDir);

            foreach (string resourceName in resourceNames)
            {
                if (resourceName.EndsWith(".lsp"))
                {
                    // Skip startup-loader.lsp as it's handled separately
                    if (resourceName.EndsWith("startup-loader.lsp"))
                        continue;

                    string fileName = resourceName;
                    int lastDot = resourceName.LastIndexOf('.');
                    if (lastDot > 0)
                    {
                        int secondLastDot = resourceName.LastIndexOf('.', lastDot - 1);
                        if (secondLastDot >= 0)
                        {
                            fileName = resourceName.Substring(secondLastDot + 1);
                        }
                    }
                    ExtractResource(resourceName, Path.Combine(lispDir, fileName));
                }
            }
        }

        private string GetProductNameFromHKLM(string versionKey, string releaseKey)
        {
            string[] regPaths = {
                @"SOFTWARE\Autodesk\AutoCAD",
                @"SOFTWARE\Wow6432Node\Autodesk\AutoCAD"
            };

            foreach (string regPath in regPaths)
            {
                try
                {
                    using (var key = Registry.LocalMachine.OpenSubKey($@"{regPath}\{versionKey}\{releaseKey}"))
                    {
                        if (key != null)
                        {
                            var productName = key.GetValue("ProductName")?.ToString();
                            if (!string.IsNullOrEmpty(productName))
                                return productName;
                        }
                    }
                }
                catch { }
            }

            return null;
        }

        private int ParseVersion(string versionKey)
        {
            if (versionKey.StartsWith("R") && versionKey.Length > 1)
            {
                string numPart = versionKey.Substring(1).Split('.')[0];
                if (int.TryParse(numPart, out int version))
                {
                    return 2000 + version;
                }
            }
            return 0;
        }

        private void CopyInstallerForUninstall(string targetDir)
        {
            try
            {
                string source = Assembly.GetExecutingAssembly().Location;
                string dest = Path.Combine(targetDir, "uninstaller.exe");
                File.Copy(source, dest, true);
            }
            catch { }
        }

        private void RegisterUninstaller(string targetDir)
        {
            try
            {
                string uninstallerPath = Path.Combine(targetDir, "uninstaller.exe");

                using (var key = Registry.CurrentUser.CreateSubKey(UNINSTALL_KEY))
                {
                    key.SetValue("DisplayName", PRODUCT_NAME);
                    key.SetValue("DisplayVersion", "2.0.0");
                    key.SetValue("Publisher", "AutoCAD Ballet");
                    key.SetValue("InstallLocation", targetDir);
                    key.SetValue("UninstallString", $"\"{uninstallerPath}\"");
                    key.SetValue("QuietUninstallString", $"\"{uninstallerPath}\" /quiet");
                    key.SetValue("NoModify", 1);
                    key.SetValue("NoRepair", 1);
                    key.SetValue("EstimatedSize", 1000);
                    key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
                    key.SetValue("DisplayIcon", uninstallerPath);
                }
            }
            catch { }
        }

        private void ExtractResource(string resourceName, string targetPath)
        {
            try
            {
                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        using (FileStream fs = new FileStream(targetPath, FileMode.Create))
                            stream.CopyTo(fs);
                    }
                }
            }
            catch { }
        }

        private bool ModifyAcadLsp(AutoCADInstallation installation)
        {
            if (installation.IsLT && installation.Version < 2024)
            {
                return false;
            }

            string acadLspPath = Path.Combine(installation.SupportPath, "acad.lsp");

            try
            {
                string backupPath = acadLspPath + ".ballet-backup";
                if (!File.Exists(backupPath) && File.Exists(acadLspPath))
                    File.Copy(acadLspPath, backupPath);

                if (File.Exists(acadLspPath))
                {
                    string content = File.ReadAllText(acadLspPath);
                    if (!content.Contains(";; autocad-ballet"))
                    {
                        File.AppendAllText(acadLspPath, Environment.NewLine + ACAD_LSP_LOADER);
                        return true;
                    }
                }
                else
                {
                    File.WriteAllText(acadLspPath, ACAD_LSP_LOADER);
                    return true;
                }
            }
            catch { }

            return false;
        }

        private class AutoCADInstallation
        {
            public string Name { get; set; }
            public int Version { get; set; }
            public string Year { get; set; }
            public string VersionKey { get; set; }
            public string ReleaseKey { get; set; }
            public bool IsLT { get; set; }
            public string SupportPath { get; set; }
        }
    }

    public class BalletUninstaller
    {
        public void Run()
        {
            try
            {
                string currentPath = Assembly.GetExecutingAssembly().Location;
                string appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "autocad-ballet");

                if (currentPath.StartsWith(appDataPath, StringComparison.OrdinalIgnoreCase))
                {
                    string tempDir = Path.Combine(Path.GetTempPath(), "autocad-ballet-uninstall");
                    Directory.CreateDirectory(tempDir);

                    string tempExe = Path.Combine(tempDir, "uninstaller.exe");
                    File.Copy(currentPath, tempExe, true);

                    string batchFile = Path.Combine(tempDir, "cleanup.bat");
                    string batchContent = $@"@echo off
timeout /t 2 /nobreak >nul
rd /s /q ""{appDataPath}""
del ""%~f0""";
                    File.WriteAllText(batchFile, batchContent);

                    var startInfo = new ProcessStartInfo
                    {
                        FileName = tempExe,
                        Arguments = $"/actualuninstall \"{batchFile}\"",
                        UseShellExecute = false
                    };
                    Process.Start(startInfo);

                    Application.Exit();
                    return;
                }

                string batchCleanup = null;
                string[] args = Environment.GetCommandLineArgs();
                if (args.Length > 2 && args[1] == "/actualuninstall")
                {
                    batchCleanup = args[2];
                }

                if (MessageBox.Show("Uninstall AutoCAD Ballet?", "AutoCAD Ballet",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    return;

                RemoveFromAcadLsp();

                string targetDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "autocad-ballet");
                if (Directory.Exists(targetDir))
                {
                    try
                    {
                        foreach (string file in Directory.GetFiles(targetDir, "*", SearchOption.AllDirectories))
                        {
                            try
                            {
                                File.SetAttributes(file, FileAttributes.Normal);
                                File.Delete(file);
                            }
                            catch { }
                        }
                        foreach (string dir in Directory.GetDirectories(targetDir, "*", SearchOption.AllDirectories).OrderByDescending(d => d.Length))
                        {
                            try { Directory.Delete(dir, true); } catch { }
                        }

                        try
                        {
                            Directory.Delete(targetDir, true);
                        }
                        catch
                        {
                            if (!string.IsNullOrEmpty(batchCleanup) && File.Exists(batchCleanup))
                            {
                                Process.Start(new ProcessStartInfo
                                {
                                    FileName = batchCleanup,
                                    WindowStyle = ProcessWindowStyle.Hidden,
                                    CreateNoWindow = true
                                });
                            }
                        }
                    }
                    catch { }
                }

                RemoveFromRegistry();

                MessageBox.Show("AutoCAD Ballet uninstalled successfully.", "Uninstaller",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Uninstallation failed: {ex.Message}", "Uninstaller",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RemoveFromAcadLsp()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string autodeskPath = Path.Combine(appData, "Autodesk");

            if (!Directory.Exists(autodeskPath)) return;

            foreach (string acadLspPath in Directory.GetFiles(autodeskPath, "acad.lsp", SearchOption.AllDirectories))
            {
                try
                {
                    string content = File.ReadAllText(acadLspPath);
                    if (content.Contains(";; autocad-ballet"))
                    {
                        var lines = content.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').ToList();

                        int startIdx = lines.FindIndex(l => l.Contains(";; autocad-ballet"));

                        if (startIdx >= 0)
                        {
                            // Find the end of the ballet loader (now much shorter)
                            int endIdx = -1;
                            for (int i = startIdx; i < lines.Count; i++)
                            {
                                if (lines[i].Contains("(load (strcat ballet-installer-path"))
                                {
                                    endIdx = i;
                                    break;
                                }
                            }

                            if (endIdx >= 0)
                            {
                                lines.RemoveRange(startIdx, endIdx - startIdx + 1);

                                if (startIdx == 0)
                                {
                                    while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[0]))
                                        lines.RemoveAt(0);
                                }

                                File.WriteAllText(acadLspPath, string.Join(Environment.NewLine, lines).Trim());
                            }
                        }

                        string backupPath = acadLspPath + ".ballet-backup";
                        if (File.Exists(backupPath))
                            File.Delete(backupPath);
                    }
                }
                catch { }
            }
        }

        private void RemoveFromRegistry()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AutoCADBallet", false);
            }
            catch { }
        }
    }
}
