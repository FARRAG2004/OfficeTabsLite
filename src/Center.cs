// OfficeTabsLite Center - settings app for the Office Tabs Lite add-in.
// Turns the tabs on/off per Office app, picks the tab look (with a live preview using the
// add-in's own TabStrip control), and resets or repairs the add-in. Installed next to
// OfficeTabsLite.dll, whose TabsSettings class it shares.
//
// Build:
//   csc /target:winexe /out:OfficeTabsLite-Center.exe /win32icon:OfficeTabsLite.ico ^
//       /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll ^
//       /r:OfficeTabsLite.dll Center.cs

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OfficeTabsLite;

[assembly: AssemblyTitle("Office Tabs Lite Center")]
[assembly: AssemblyProduct("Office Tabs Lite")]
[assembly: AssemblyVersion("2.7.0.0")]
[assembly: AssemblyFileVersion("2.7.0.0")]

namespace OfficeTabsLiteCenter
{
    internal static class Program
    {
        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main()
        {
            try { SetProcessDPIAware(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new CenterForm());
        }
    }

    internal class CenterForm : Form
    {
        const string Version = "2.7";
        static readonly Color NavBack = Color.FromArgb(243, 243, 243);
        static readonly Color Accent = Color.FromArgb(43, 87, 154);
        static readonly Color Muted = Color.FromArgb(105, 105, 105);

        readonly Label[] _navItems;
        readonly Panel[] _pages;
        readonly Panel _navBar = new Panel();

        readonly CheckBox[] _enable = new CheckBox[3];
        readonly CheckBox[] _reset = new CheckBox[3];
        readonly RadioButton _rounded = new RadioButton(), _trapezoid = new RadioButton(), _flat = new RadioButton();
        readonly CheckBox _group = new CheckBox(), _icons = new CheckBox(), _doubleClick = new CheckBox(), _hideSingle = new CheckBox();
        readonly TabStrip _preview = new TabStrip();

        public CenterForm()
        {
            SuspendLayout();
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = "Office Tabs Lite Center";
            Font = new Font("Segoe UI", 9.75f);
            ClientSize = new Size(700, 470);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            // bottom bar: OK / Cancel / Apply
            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 58, BackColor = NavBack };
            bottom.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Color.FromArgb(215, 215, 215) });
            var version = new Label { Text = "Office Tabs Lite " + Version, ForeColor = Muted, AutoSize = true, Location = new Point(16, 21) };
            var ok = MakeButton("OK", 380, true);
            var cancel = MakeButton("Cancel", 485, false);
            var apply = MakeButton("Apply", 590, false);
            ok.Click += delegate { if (ApplyChanges()) Close(); };
            cancel.Click += delegate { Close(); };
            apply.Click += delegate { ApplyChanges(); };
            AcceptButton = ok;
            CancelButton = cancel;
            bottom.Controls.AddRange(new Control[] { version, ok, cancel, apply });

            // left navigation
            var nav = new Panel { Dock = DockStyle.Left, Width = 160, BackColor = NavBack };
            _navBar.BackColor = Accent;
            _navBar.Size = new Size(4, 30);
            nav.Controls.Add(_navBar);
            string[] names = { "General", "Appearance", "Advanced", "About" };
            _navItems = new Label[names.Length];
            for (int i = 0; i < names.Length; i++)
            {
                int index = i;
                var item = new Label
                {
                    Text = names[i], AutoSize = false, Size = new Size(150, 36), Location = new Point(10, 18 + i * 44),
                    Font = new Font("Segoe UI", 13f), TextAlign = ContentAlignment.MiddleLeft, Cursor = Cursors.Hand
                };
                item.Click += delegate { ShowPage(index); };
                _navItems[i] = item;
                nav.Controls.Add(item);
            }

            // pages
            var content = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            _pages = new[] { BuildGeneral(), BuildAppearance(), BuildAdvanced(), BuildAbout() };
            foreach (Panel p in _pages) content.Controls.Add(p);

            Controls.Add(content);
            Controls.Add(nav);
            Controls.Add(bottom);
            ResumeLayout(false);

            LoadValues();
            ShowPage(0);
        }

        // ---------- pages ----------
        Panel BuildGeneral()
        {
            Panel p = NewPage("General Options");
            AddSection(p, "Enable Office Tabs Lite for", 64);
            for (int i = 0; i < 3; i++)
            {
                _enable[i] = new CheckBox { Text = "Enable tabs for " + TabsSettings.Apps[i], AutoSize = true, Location = new Point(28, 100 + i * 32) };
                p.Controls.Add(_enable[i]);
            }
            AddNote(p, "Changes take effect the next time you open the app.", 205);
            return p;
        }

        Panel BuildAppearance()
        {
            Panel p = NewPage("Appearance");
            AddSection(p, "Tab style", 64);
            _rounded.Text = "Rounded"; _trapezoid.Text = "Trapezoid"; _flat.Text = "Flat";
            RadioButton[] radios = { _rounded, _trapezoid, _flat };
            for (int i = 0; i < radios.Length; i++)
            {
                radios[i].AutoSize = true;
                radios[i].Location = new Point(28 + i * 120, 98);
                radios[i].CheckedChanged += delegate { UpdatePreview(); };
                p.Controls.Add(radios[i]);
            }

            var frame = new Panel { Location = new Point(28, 132), Size = new Size(470, 36), BorderStyle = BorderStyle.FixedSingle };
            _preview.Dock = DockStyle.Fill;
            frame.Controls.Add(_preview);
            p.Controls.Add(frame);
            AddNote(p, "Preview - click a tab to see how the selected tab looks.", 172);

            AddSection(p, "Options", 205);
            _group.Text = "One taskbar button for all files (one-window look)";
            _icons.Text = "Show file icons on tabs";
            _doubleClick.Text = "Double-click empty space to open a new file";
            _hideSingle.Text = "Hide the tab bar when only one file is open";
            CheckBox[] checks = { _group, _icons, _doubleClick, _hideSingle };
            for (int i = 0; i < checks.Length; i++)
            {
                checks[i].AutoSize = true;
                checks[i].Location = new Point(28, 240 + i * 30);
                p.Controls.Add(checks[i]);
            }
            _icons.CheckedChanged += delegate { UpdatePreview(); };
            AddNote(p, "Open Office windows pick up these changes within a few seconds.", 365);
            return p;
        }

        Panel BuildAdvanced()
        {
            Panel p = NewPage("Advanced Options");
            AddSection(p, "Reset component configuration", 64);
            for (int i = 0; i < 3; i++)
            {
                _reset[i] = new CheckBox { Text = "Reset Tabs for " + TabsSettings.Apps[i] + " configuration", AutoSize = true, Location = new Point(28, 100 + i * 32) };
                p.Controls.Add(_reset[i]);
            }
            AddNote(p, "Turns the add-in back on for that app and clears the records that make\n" +
                       "Office disable it after a crash or a slow start. Applied with OK or Apply.", 200);

            AddSection(p, "Repair Office Tabs Lite", 255);
            var repair = new Button { Text = "Repair Office Tabs Lite...", Size = new Size(210, 32), Location = new Point(28, 290) };
            repair.Click += delegate { Repair(); };
            p.Controls.Add(repair);
            return p;
        }

        Panel BuildAbout()
        {
            Panel p = NewPage(null);
            var logo = new PictureBox { Location = new Point(28, 28), Size = new Size(64, 64), SizeMode = PictureBoxSizeMode.Zoom };
            try { logo.Image = new Icon(Icon, 64, 64).ToBitmap(); } catch { }
            p.Controls.Add(logo);
            p.Controls.Add(new Label { Text = "Office Tabs Lite", Font = new Font("Segoe UI", 16f, FontStyle.Bold), AutoSize = true, Location = new Point(108, 30) });
            p.Controls.Add(new Label { Text = "Version " + Version, ForeColor = Muted, AutoSize = true, Location = new Point(110, 66) });
            p.Controls.Add(new Label
            {
                Text = "Adds a tab for every open file to Word, Excel and PowerPoint, right under the\n" +
                       "ribbon, and keeps all files of an app under one taskbar button.\n\n" +
                       "Click a tab to switch files, x or middle-click to close one, + to open a new one.",
                AutoSize = true, Location = new Point(28, 120)
            });
            AddNote(p, "Original work. Not affiliated with Microsoft or ExtendOffice.", 215);
            return p;
        }

        // ---------- values ----------
        void LoadValues()
        {
            for (int i = 0; i < 3; i++)
            {
                bool? on = TabsSettings.IsEnabledFor(TabsSettings.Apps[i]);
                _enable[i].Checked = on == true;
                _enable[i].Enabled = on.HasValue;
                _reset[i].Enabled = on.HasValue;
                if (!on.HasValue) _enable[i].Text += "  (not installed)";
            }
            TabsSettings s = TabsSettings.Load();
            _rounded.Checked = s.Style == TabStyle.Rounded;
            _trapezoid.Checked = s.Style == TabStyle.Trapezoid;
            _flat.Checked = s.Style == TabStyle.Flat;
            _group.Checked = s.GroupTaskbar;
            _icons.Checked = s.ShowIcons;
            _doubleClick.Checked = s.NewOnDoubleClick;
            _hideSingle.Checked = s.HideSingleTab;
            UpdatePreview();
        }

        TabStyle SelectedStyle()
        {
            if (_trapezoid.Checked) return TabStyle.Trapezoid;
            if (_flat.Checked) return TabStyle.Flat;
            return TabStyle.Rounded;
        }

        void UpdatePreview()
        {
            _preview.ShowPreview(SelectedStyle().ToString(), _icons.Checked);
        }

        bool ApplyChanges()
        {
            try
            {
                var s = new TabsSettings();
                s.Style = SelectedStyle();
                s.GroupTaskbar = _group.Checked;
                s.ShowIcons = _icons.Checked;
                s.NewOnDoubleClick = _doubleClick.Checked;
                s.HideSingleTab = _hideSingle.Checked;
                s.Save();

                string resetDone = "";
                for (int i = 0; i < 3; i++)
                {
                    string app = TabsSettings.Apps[i];
                    if (_reset[i].Checked)
                    {
                        TabsSettings.ResetApp(app);
                        _reset[i].Checked = false;
                        _enable[i].Checked = true;
                        resetDone += (resetDone.Length > 0 ? ", " : "") + app;
                    }
                    if (_enable[i].Enabled) TabsSettings.SetEnabledFor(app, _enable[i].Checked);
                }
                if (resetDone.Length > 0)
                    MessageBox.Show(this, "Configuration reset for: " + resetDone + ".\nRestart that app to load the tabs again.",
                                    Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Could not save the settings:\n\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // Re-runs the installed setup in repair mode (re-extracts and re-registers the add-in).
        void Repair()
        {
            string setup = Path.Combine(Path.GetDirectoryName(Application.StartupPath), "OfficeTabsLite-Setup.exe");
            if (!File.Exists(setup))
            {
                MessageBox.Show(this, "The installer was not found next to the program:\n" + setup +
                                "\n\nRun OfficeTabsLite-Setup.exe again to repair.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                using (Process p = Process.Start(setup, "/repair")) p.WaitForExit();
                LoadValues();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Repair failed:\n\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ---------- helpers ----------
        void ShowPage(int index)
        {
            for (int i = 0; i < _pages.Length; i++)
            {
                _pages[i].Visible = i == index;
                _navItems[i].Font = new Font("Segoe UI", 13f, i == index ? FontStyle.Bold : FontStyle.Regular);
                _navItems[i].ForeColor = i == index ? Color.Black : Color.FromArgb(70, 70, 70);
            }
            _navBar.Location = new Point(0, _navItems[index].Top + 3);
            _pages[index].BringToFront();
        }

        Panel NewPage(string title)
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Visible = false };
            if (title != null)
            {
                var logo = new PictureBox { Location = new Point(24, 18), Size = new Size(28, 28), SizeMode = PictureBoxSizeMode.Zoom };
                try { logo.Image = new Icon(Icon, 32, 32).ToBitmap(); } catch { }
                p.Controls.Add(logo);
                p.Controls.Add(new Label { Text = title, Font = new Font("Segoe UI", 13f), AutoSize = true, Location = new Point(58, 19) });
            }
            return p;
        }

        static void AddSection(Panel page, string text, int y)
        {
            page.Controls.Add(new Label { Text = text, Font = new Font("Segoe UI", 10.5f, FontStyle.Bold), AutoSize = true, Location = new Point(22, y) });
            page.Controls.Add(new Panel { BackColor = Color.FromArgb(215, 215, 215), Location = new Point(22, y + 26), Size = new Size(500, 1) });
        }

        static void AddNote(Panel page, string text, int y)
        {
            page.Controls.Add(new Label { Text = text, ForeColor = Muted, AutoSize = true, Location = new Point(28, y) });
        }

        static Button MakeButton(string text, int x, bool primary)
        {
            var b = new Button { Text = text, Size = new Size(95, 32), Location = new Point(x, 13) };
            if (primary)
            {
                b.BackColor = Accent;
                b.ForeColor = Color.White;
                b.FlatStyle = FlatStyle.Flat;
                b.FlatAppearance.BorderColor = Accent;
            }
            return b;
        }
    }
}
