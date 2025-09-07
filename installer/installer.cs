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

        private const string LOADER_CODE = @";; autocad-ballet
;; please don't change these lines unless you know what you are doing
;; load dotnet commands
(setq temp-folder (strcat (getenv ""TEMP"") ""/autocad-ballet/""))
(setq source-dll (strcat (getenv ""APPDATA"") ""/autocad-ballet/autocad-ballet.dll""))
(setq temp-dll (strcat temp-folder ""/autocad-ballet.dll""))
;; create temp folder if doesn't exist
(vl-mkdir temp-folder)
;; copy dll to temp preserving date modified
(setq fso (vlax-create-object ""Scripting.FileSystemObject""))
(vl-catch-all-apply 'vlax-invoke-method
  (list fso 'CopyFile source-dll temp-dll :vlax-true))  ; overwrite
(vlax-release-object fso)
;; load dll assembly
(command ""netload"" temp-dll)
;; load lisp commands
(setq lisp-folder (strcat (getenv ""APPDATA"") ""/autocad-ballet/autocad-ballet/""))
(foreach file (vl-directory-files lisp-folder ""*.lsp"")
  (load (strcat lisp-folder file)))
";

        private const string LISP_ONLY_LOADER = @";; autocad-ballet
;; please don't change these lines unless you know what you are doing
;; load lisp commands (AutoCAD LT)
(setq lisp-folder (strcat (getenv ""APPDATA"") ""/autocad-ballet/autocad-ballet/""))
(foreach file (vl-directory-files lisp-folder ""*.lsp"")
  (load (strcat lisp-folder file)))
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

                ExtractEmbeddedFiles(targetDir);
                CopyInstallerForUninstall(targetDir);

                var modifiedInstallations = new List<AutoCADInstallation>();
                foreach (var installation in installations)
                {
                    bool modified = ModifyAcadLsp(installation);
                    if (modified)
                    {
                        modifiedInstallations.Add(installation);
                    }
                    AddToTrustedPaths(installation);
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
                // Load the embedded PNG resource
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
                // Logo image
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
                // Fallback text if logo not found
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

            // Build the message text
            string messageText = "autocad-ballet has been successfully installed\n";

            if (modifiedInstallations.Count > 0)
            {
                foreach (var installation in modifiedInstallations)
                {
                    messageText += $"• {installation.Name}\n";
                }
            }
            else
            {
                messageText += "The plugin files have been installed but no AutoCAD installations were modified.\n";
                messageText += "AutoCAD LT versions prior to 2024 are not supported for automatic loading.";
            }

            // Calculate text height
            using (var g = form.CreateGraphics())
            {
                var textSize = g.MeasureString(messageText, new Font("Segoe UI", 10), 450);
                var textHeight = (int)Math.Ceiling(textSize.Height);

                // Message
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

            // OK button
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

            // Set form client size based on actual content
            currentY += 50; // Add padding below button
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
                                            installations.Add(new AutoCADInstallation
                                            {
                                                Name = productName,
                                                Version = ParseVersion(versionKey),
                                                VersionKey = versionKey,
                                                ReleaseKey = releaseKey,
                                                IsLT = productName.Contains("LT"),
                                                SupportPath = supportPath
                                            });
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

        private void AddToTrustedPaths(AutoCADInstallation installation)
        {
            try
            {
                // Get the profiles path
                string profilesPath = $@"SOFTWARE\Autodesk\AutoCAD\{installation.VersionKey}\{installation.ReleaseKey}\Profiles";

                using (var profilesKey = Registry.CurrentUser.OpenSubKey(profilesPath))
                {
                    if (profilesKey == null) return;

                    foreach (string profileName in profilesKey.GetSubKeyNames())
                    {
                        string variablesPath = $@"{profilesPath}\{profileName}\Variables";

                        using (var variablesKey = Registry.CurrentUser.OpenSubKey(variablesPath, true))
                        {
                            if (variablesKey == null) continue;

                            // Get existing trusted paths
                            string trustedPaths = variablesKey.GetValue("TRUSTEDPATHS")?.ToString() ?? "";

                            // Use the environment variable format
                            string pathToAdd = @"%appdata%\autocad-ballet\autocad-ballet";

                            // Check if our path is already present (check both formats)
                            string expandedPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "autocad-ballet", "autocad-ballet");

                            if (!trustedPaths.Contains(pathToAdd) && !trustedPaths.Contains(expandedPath))
                            {
                                if (!string.IsNullOrEmpty(trustedPaths) && !trustedPaths.EndsWith(";"))
                                    trustedPaths += ";";

                                trustedPaths += pathToAdd + ";";

                                // Set as REG_EXPAND_SZ to preserve environment variables
                                variablesKey.SetValue("TRUSTEDPATHS", trustedPaths, RegistryValueKind.ExpandString);
                            }
                        }
                    }
                }
            }
            catch { }
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
                // For Add/Remove Programs icon, we'll just use the executable path
                string uninstallerPath = Path.Combine(targetDir, "uninstaller.exe");

                // Register in HKEY_CURRENT_USER for current user only
                using (var key = Registry.CurrentUser.CreateSubKey(UNINSTALL_KEY))
                {
                    key.SetValue("DisplayName", PRODUCT_NAME);
                    key.SetValue("DisplayVersion", "1.0.0");
                    key.SetValue("Publisher", "AutoCAD Ballet");
                    key.SetValue("InstallLocation", targetDir);
                    key.SetValue("UninstallString", $"\"{uninstallerPath}\"");
                    key.SetValue("QuietUninstallString", $"\"{uninstallerPath}\" /quiet");
                    key.SetValue("NoModify", 1);
                    key.SetValue("NoRepair", 1);
                    key.SetValue("EstimatedSize", 100);
                    key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
                    // Use the uninstaller exe as the icon source
                    key.SetValue("DisplayIcon", uninstallerPath);
                }
            }
            catch { }
        }

        private void ExtractEmbeddedFiles(string targetDir)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();

            foreach (string resourceName in resourceNames)
            {
                if (resourceName.EndsWith("autocad-ballet.dll"))
                {
                    ExtractResource(resourceName, Path.Combine(targetDir, "autocad-ballet.dll"));
                }
            }

            string lispDir = Path.Combine(targetDir, "autocad-ballet");
            Directory.CreateDirectory(lispDir);

            foreach (string resourceName in resourceNames)
            {
                if (resourceName.EndsWith(".lsp"))
                {
                    // Extract just the filename without any prefixes
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
            string loaderCode = installation.IsLT ? LISP_ONLY_LOADER : LOADER_CODE;

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
                        File.AppendAllText(acadLspPath, Environment.NewLine + loaderCode);
                        return true;
                    }
                }
                else
                {
                    File.WriteAllText(acadLspPath, loaderCode);
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
                // Check if we're running from the installation directory
                string currentPath = Assembly.GetExecutingAssembly().Location;
                string appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "autocad-ballet");

                if (currentPath.StartsWith(appDataPath, StringComparison.OrdinalIgnoreCase))
                {
                    // We're running from the installation directory, need to copy to temp and relaunch
                    string tempDir = Path.Combine(Path.GetTempPath(), "autocad-ballet-uninstall");
                    Directory.CreateDirectory(tempDir);

                    string tempExe = Path.Combine(tempDir, "uninstaller.exe");
                    File.Copy(currentPath, tempExe, true);

                    // Create a batch file to delete the directory after a delay
                    string batchFile = Path.Combine(tempDir, "cleanup.bat");
                    string batchContent = $@"@echo off
timeout /t 2 /nobreak >nul
rd /s /q ""{appDataPath}""
del ""%~f0""";
                    File.WriteAllText(batchFile, batchContent);

                    // Start the uninstaller from temp location
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = tempExe,
                        Arguments = $"/actualuninstall \"{batchFile}\"",
                        UseShellExecute = false
                    };
                    Process.Start(startInfo);

                    // Exit this instance
                    Application.Exit();
                    return;
                }

                // We're running from temp or with /actualuninstall flag, proceed with uninstall
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
                RemoveFromTrustedPaths();

                // Remove the installation directory completely
                string targetDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "autocad-ballet");
                if (Directory.Exists(targetDir))
                {
                    try
                    {
                        // Try to delete everything we can
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

                        // Try to delete the main directory
                        try
                        {
                            Directory.Delete(targetDir, true);
                        }
                        catch
                        {
                            // If we can't delete it, run the batch file to clean up
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
                        // Split into lines
                        var lines = content.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').ToList();

                        // Find the start of our block
                        int startIdx = lines.FindIndex(l => l.Contains(";; autocad-ballet"));

                        if (startIdx >= 0)
                        {
                            // Find the end - look for the last line of our code
                            int endIdx = -1;
                            for (int i = startIdx; i < lines.Count; i++)
                            {
                                if (lines[i].Contains("(load (strcat lisp-folder file)))"))
                                {
                                    endIdx = i;
                                    break;
                                }
                            }

                            if (endIdx >= 0)
                            {
                                // Remove the entire block including the last line
                                lines.RemoveRange(startIdx, endIdx - startIdx + 1);

                                // Clean up any blank lines at the start if our block was at the beginning
                                if (startIdx == 0)
                                {
                                    while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[0]))
                                        lines.RemoveAt(0);
                                }

                                File.WriteAllText(acadLspPath, string.Join(Environment.NewLine, lines).Trim());
                            }
                        }

                        // Remove backup file
                        string backupPath = acadLspPath + ".ballet-backup";
                        if (File.Exists(backupPath))
                            File.Delete(backupPath);
                    }
                }
                catch { }
            }
        }

        private void RemoveFromTrustedPaths()
        {
            try
            {
                string pathToRemove = @"%appdata%\autocad-ballet\autocad-ballet";
                string expandedPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "autocad-ballet", "autocad-ballet");

                using (var acadKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Autodesk\AutoCAD"))
                {
                    if (acadKey == null) return;

                    foreach (string versionKey in acadKey.GetSubKeyNames())
                    {
                        using (var vKey = acadKey.OpenSubKey(versionKey))
                        {
                            if (vKey == null) continue;

                            foreach (string releaseKey in vKey.GetSubKeyNames())
                            {
                                string profilesPath = $@"SOFTWARE\Autodesk\AutoCAD\{versionKey}\{releaseKey}\Profiles";

                                using (var profilesKey = Registry.CurrentUser.OpenSubKey(profilesPath))
                                {
                                    if (profilesKey == null) continue;

                                    foreach (string profileName in profilesKey.GetSubKeyNames())
                                    {
                                        string variablesPath = $@"{profilesPath}\{profileName}\Variables";

                                        using (var variablesKey = Registry.CurrentUser.OpenSubKey(variablesPath, true))
                                        {
                                            if (variablesKey == null) continue;

                                            string trustedPaths = variablesKey.GetValue("TRUSTEDPATHS")?.ToString() ?? "";

                                            // Remove both possible formats
                                            if (trustedPaths.Contains(pathToRemove) || trustedPaths.Contains(expandedPath))
                                            {
                                                trustedPaths = trustedPaths.Replace(pathToRemove + ";", "");
                                                trustedPaths = trustedPaths.Replace(pathToRemove, "");
                                                trustedPaths = trustedPaths.Replace(expandedPath + ";", "");
                                                trustedPaths = trustedPaths.Replace(expandedPath, "");
                                                variablesKey.SetValue("TRUSTEDPATHS", trustedPaths, RegistryValueKind.ExpandString);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { }
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
