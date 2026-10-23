// OfficeTabsLite-Setup - single-file, per-user installer for the OfficeTabsLite add-in.
// Embeds OfficeTabsLite.dll as a resource, extracts it to %LOCALAPPDATA%, and
// registers a managed COM in-proc server + Office add-in keys entirely under HKCU
// (no administrator rights, works on any laptop).
//
// Build:
//   csc /target:winexe /out:OfficeTabsLite-Setup.exe /win32icon:OfficeTabsLite.ico ^
//       /r:System.dll /r:System.Windows.Forms.dll ^
//       /resource:OfficeTabsLite.dll,OfficeTabsLite.dll ^
//       /resource:OfficeTabsLite-Center.exe,OfficeTabsLite-Center.exe Installer.cs
// Switches: /silent, /uninstall, /repair

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Office Tabs Lite Setup")]
[assembly: AssemblyProduct("Office Tabs Lite")]
[assembly: AssemblyVersion("2.7.0.0")]
[assembly: AssemblyFileVersion("2.7.0.0")]

namespace OfficeTabsLiteSetup
{
    internal static class Program
    {
        const string Version    = "2.7.0";
        const string CenterFile = "OfficeTabsLite-Center.exe";
        const string ShortcutName = "Office Tabs Lite Center.lnk";
        const string SetupFile  = "OfficeTabsLite-Setup.exe";
        const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeTabsLite";
        const string ProgId    = "OfficeTabsLite.Connect";
        const string Clsid      = "{B3D9F2A4-5C7E-4A1B-9E6D-2F8A1C0B3D44}";
        const string StripProgId = "OfficeTabsLite.TabStrip";
        const string StripClsid  = "{6E0C2D8A-3F41-4B7C-A9D2-5B1E7F3C8A61}";
        const string Friendly   = "Office Tabs Lite";
        const string Descr      = "Document tabs for Word, Excel and PowerPoint.";
        const string AsmName    = "OfficeTabsLite, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null";
        const string DllFile    = "OfficeTabsLite.dll";
        static readonly string[] Apps = { "Word", "Excel", "PowerPoint" };

        static string InstallDir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeTabsLite"); }
        }
        static string BinDir
        {
            get { return Path.Combine(InstallDir, "bin"); }
        }
        static string DllPath
        {
            get { return Path.Combine(BinDir, DllFile); }
        }

        [STAThread]
        static int Main(string[] args)
        {
            bool silent    = Array.Exists(args, a => a.Equals("/silent", StringComparison.OrdinalIgnoreCase));
            bool uninstall = Array.Exists(args, a => a.Equals("/uninstall", StringComparison.OrdinalIgnoreCase));
            bool repair    = Array.Exists(args, a => a.Equals("/repair", StringComparison.OrdinalIgnoreCase));

            try
            {
                if (repair)
                {
                    Install(true);
                    Info(silent, "Office Tabs Lite was repaired.\nRestart any open Office apps.");
                    return 0;
                }

                DialogResult existing = DialogResult.None;
                if (!uninstall && !silent && IsInstalled())
                {
                    existing = AskExisting();
                    if (existing == DialogResult.Cancel) return 0;
                }

                if (uninstall || existing == DialogResult.No)
                {
                    Uninstall();
                    Info(silent, "Office Tabs Lite was removed.\nRestart any open Office apps.");
                    return 0;
                }

                if (!silent && existing == DialogResult.None && !Confirm())
                    return 0;

                Install(false);
                Info(silent, "Office Tabs Lite installed successfully.\n\n" +
                             "Open Word, Excel or PowerPoint: a tab for each open file appears\n" +
                             "under the ribbon. (If an app is already open, restart it.)\n\n" +
                             "Settings: \"Office Tabs Lite Center\" on the desktop and in the Start menu.");
                return 0;
            }
            catch (Exception ex)
            {
                if (silent) { Console.Error.WriteLine(ex); }
                else MessageBox.Show("Setup failed:\n\n" + ex.Message, "Office Tabs Lite",
                                     MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        // ---------- install / uninstall ----------
        // repair: also turns the add-in back on for every app (an update keeps the user's choice).
        static void Install(bool repair)
        {
            Directory.CreateDirectory(BinDir);
            try { ExtractResource(DllFile, DllPath); }
            catch (IOException)
            {
                throw new Exception("The add-in file is in use.\n" +
                                    "Close Word, Excel and PowerPoint, then run setup again.");
            }
            // The Center may be the program running this repair; its file is then in use but current.
            try { ExtractResource(CenterFile, Path.Combine(BinDir, CenterFile)); } catch (IOException) { }

            // 1) managed COM in-proc servers, per-user (HKCU\Software\Classes):
            //    the add-in itself and the tab strip control its task panes host.
            //    Written to both registry views so 64-bit and 32-bit Office find them.
            foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                using (var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                using (var classes = hkcu.CreateSubKey(@"Software\Classes"))
                {
                    RegisterClass(classes, Clsid, ProgId);
                    RegisterClass(classes, StripClsid, StripProgId);
                }

            // 2) Office add-in registration, per-user, for each app
            foreach (var app in Apps)
            {
                using (var k = Registry.CurrentUser.CreateSubKey(
                           @"Software\Microsoft\Office\" + app + @"\Addins\" + ProgId))
                {
                    k.SetValue("FriendlyName", Friendly);
                    k.SetValue("Description", Descr);
                    if (repair || k.GetValue("LoadBehavior") == null)
                        k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                }
            }

            // 3) keep a copy of setup next to the add-in and list it in Settings > Apps
            string installedSetup = Path.Combine(InstallDir, SetupFile);
            string self = Application.ExecutablePath;
            if (!string.Equals(Path.GetFullPath(self), Path.GetFullPath(installedSetup), StringComparison.OrdinalIgnoreCase))
                File.Copy(self, installedSetup, true);

            using (var u = Registry.CurrentUser.CreateSubKey(UninstallKey))
            {
                u.SetValue("DisplayName", Friendly);
                u.SetValue("DisplayVersion", Version);
                u.SetValue("Publisher", Friendly);
                u.SetValue("DisplayIcon", installedSetup + ",0");
                u.SetValue("InstallLocation", InstallDir);
                u.SetValue("UninstallString", "\"" + installedSetup + "\" /uninstall");
                u.SetValue("QuietUninstallString", "\"" + installedSetup + "\" /uninstall /silent");
                u.SetValue("NoModify", 1, RegistryValueKind.DWord);
                u.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                u.SetValue("EstimatedSize", 200, RegistryValueKind.DWord);
            }

            // 4) shortcuts to the Center settings app
            string center = Path.Combine(BinDir, CenterFile);
            foreach (string dir in ShortcutFolders())
            {
                try { CreateShortcut(Path.Combine(dir, ShortcutName), center); } catch { }
            }
        }

        static string[] ShortcutFolders()
        {
            return new[] { Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                           Environment.GetFolderPath(Environment.SpecialFolder.Programs) };
        }

        // Uses the shell's WScript.Shell object, late-bound so no extra references are needed.
        static void CreateShortcut(string lnkPath, string target)
        {
            Type shellType = Type.GetTypeFromProgID("WScript.Shell");
            object shell = Activator.CreateInstance(shellType);
            object lnk = shellType.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, shell, new object[] { lnkPath });
            Type t = lnk.GetType();
            t.InvokeMember("TargetPath", BindingFlags.SetProperty, null, lnk, new object[] { target });
            t.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, lnk, new object[] { Path.GetDirectoryName(target) });
            t.InvokeMember("IconLocation", BindingFlags.SetProperty, null, lnk, new object[] { target + ",0" });
            t.InvokeMember("Description", BindingFlags.SetProperty, null, lnk, new object[] { "Office Tabs Lite settings" });
            t.InvokeMember("Save", BindingFlags.InvokeMethod, null, lnk, null);
            Marshal.ReleaseComObject(lnk);
            Marshal.ReleaseComObject(shell);
        }

        static void Uninstall()
        {
            foreach (var app in Apps)
                TryDelete(Registry.CurrentUser, @"Software\Microsoft\Office\" + app + @"\Addins\" + ProgId);

            foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                using (var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                {
                    TryDelete(hkcu, @"Software\Classes\CLSID\" + Clsid);
                    TryDelete(hkcu, @"Software\Classes\" + ProgId);
                    TryDelete(hkcu, @"Software\Classes\CLSID\" + StripClsid);
                    TryDelete(hkcu, @"Software\Classes\" + StripProgId);
                }
            TryDelete(Registry.CurrentUser, UninstallKey);
            TryDelete(Registry.CurrentUser, @"Software\OfficeTabsLite");
            foreach (string dir in ShortcutFolders())
            {
                try { File.Delete(Path.Combine(dir, ShortcutName)); } catch { }
            }

            try { if (Directory.Exists(InstallDir)) Directory.Delete(InstallDir, true); } catch { }

            // When run from Settings > Apps, this exe lives in InstallDir and is locked while
            // running: remove the folder a moment after we exit.
            if (Directory.Exists(InstallDir))
            {
                try
                {
                    var psi = new ProcessStartInfo("cmd.exe",
                        "/c ping 127.0.0.1 -n 3 > nul & rmdir /s /q \"" + InstallDir + "\"");
                    psi.CreateNoWindow = true;
                    psi.UseShellExecute = false;
                    Process.Start(psi);
                }
                catch { }
            }
        }

        // ---------- helpers ----------
        // Writes what RegAsm /codebase writes, but under HKCU so no admin rights are needed.
        static void RegisterClass(RegistryKey classes, string clsidText, string progId)
        {
            string codeBase = new Uri(DllPath).AbsoluteUri;
            using (var clsid = classes.CreateSubKey("CLSID\\" + clsidText))
            {
                clsid.SetValue(null, progId);
                using (var ips = clsid.CreateSubKey("InprocServer32"))
                {
                    ips.SetValue(null, "mscoree.dll");
                    ips.SetValue("ThreadingModel", "Both");
                    ips.SetValue("Class", progId);
                    ips.SetValue("Assembly", AsmName);
                    ips.SetValue("RuntimeVersion", "v4.0.30319");
                    ips.SetValue("CodeBase", codeBase);
                    using (var ver = ips.CreateSubKey("1.0.0.0"))
                    {
                        ver.SetValue("Class", progId);
                        ver.SetValue("Assembly", AsmName);
                        ver.SetValue("RuntimeVersion", "v4.0.30319");
                        ver.SetValue("CodeBase", codeBase);
                    }
                }
                clsid.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}").Close();
                using (var pid = clsid.CreateSubKey("ProgId")) pid.SetValue(null, progId);
            }
            using (var pk = classes.CreateSubKey(progId))
            {
                pk.SetValue(null, progId);
                using (var pc = pk.CreateSubKey("CLSID")) pc.SetValue(null, clsidText);
            }
        }

        static void ExtractResource(string name, string dest)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var s = asm.GetManifestResourceStream(name))
            {
                if (s == null) throw new Exception("Embedded payload not found: " + name);
                using (var f = File.Create(dest)) s.CopyTo(f);
            }
        }

        static bool IsInstalled()
        {
            using (var k = Registry.CurrentUser.OpenSubKey(@"Software\Classes\" + ProgId))
                return k != null;
        }

        static void TryDelete(RegistryKey root, string path)
        {
            try { root.DeleteSubKeyTree(path, false); } catch { }
        }

        static bool Confirm()
        {
            return MessageBox.Show(
                "Install Office Tabs Lite?\n\n" +
                "Adds a tab for each open file under the ribbon of\n" +
                "Word, Excel and PowerPoint. Installs for the current user only\n" +
                "(no administrator rights needed).",
                "Office Tabs Lite - Setup",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK;
        }

        // Yes = update/reinstall, No = remove, Cancel = do nothing.
        static DialogResult AskExisting()
        {
            return MessageBox.Show(
                "Office Tabs Lite is already installed.\n\n" +
                "Yes - update / reinstall\nNo - remove it\nCancel - do nothing",
                "Office Tabs Lite - Setup",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        }

        static void Info(bool silent, string msg)
        {
            if (!silent)
                MessageBox.Show(msg, "Office Tabs Lite", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
