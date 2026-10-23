// OfficeTabsLite - original tabbed-document add-in for Microsoft Office.
// Shows a tab strip under the ribbon of every Word, Excel and PowerPoint window
// (one tab per open document), plus a "Tabs" ribbon tab. Uses only the documented
// Office object model: Custom Task Panes docked at the top, Application.Windows and
// Window.Activate. No window reparenting, no hooks. Settings are shared with the
// OfficeTabsLite Center app through HKCU\Software\OfficeTabsLite.
//
// Build (interop types are embedded with /link, so no PIA is needed at runtime):
//   csc /target:library /out:OfficeTabsLite.dll ^
//       /r:System.dll /r:System.Core.dll /r:Microsoft.CSharp.dll ^
//       /r:System.Drawing.dll /r:System.Windows.Forms.dll ^
//       /link:Extensibility.dll /link:office.dll OfficeTabsLite.cs
//
// SPDX: original work. Not derived from any third-party source code.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Win32;

[assembly: AssemblyTitle("OfficeTabsLite")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("2.7.0.0")]
[assembly: ComVisible(false)]

namespace OfficeTabsLite
{
    internal static class Native
    {
        public const uint GA_ROOT = 2;
        public const int SW_SHOWNORMAL = 1, SW_SHOWMINIMIZED = 2, SW_SHOWMAXIMIZED = 3, SW_SHOWMINNOACTIVE = 7;
        public const int WPF_RESTORETOMAXIMIZED = 2;

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int X, Y; }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            public int length, flags, showCmd;
            public POINT ptMinPosition, ptMaxPosition;
            public RECT rcNormalPosition;
        }

        [DllImport("user32.dll")] public static extern bool GetWindowPlacement(IntPtr hwnd, ref WINDOWPLACEMENT p);
        [DllImport("user32.dll")] public static extern bool SetWindowPlacement(IntPtr hwnd, ref WINDOWPLACEMENT p);
        [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hwnd);
        [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int cmd);
        [DllImport("user32.dll")] public static extern bool DestroyIcon(IntPtr hIcon);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder text, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder name, int max);
        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);

        // Turns off the window open/maximize/restore animation for one window.
        public const int DWMWA_TRANSITIONS_FORCEDISABLED = 3;
        [DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        // Icon of the program registered for a file type (e.g. the Word icon for ".docx").
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
        }
        public const uint SHGFI_ICON = 0x100, SHGFI_SMALLICON = 0x1, SHGFI_USEFILEATTRIBUTES = 0x10;
        public const uint FILE_ATTRIBUTE_NORMAL = 0x80;
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr SHGetFileInfo(string path, uint attributes, ref SHFILEINFO info, uint size, uint flags);

        // Shell taskbar list: adds/removes a window's taskbar button.
        [ComImport, Guid("56FDF342-FD6D-11d0-958A-006097C9A090")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface ITaskbarList
        {
            void HrInit();
            void AddTab(IntPtr hwnd);
            void DeleteTab(IntPtr hwnd);
            void ActivateTab(IntPtr hwnd);
            void SetActiveAlt(IntPtr hwnd);
        }

        [ComImport, Guid("56FDF344-FD6D-11d0-958A-006097C9A090")]
        public class TaskbarList { }
    }

    internal class TabItem
    {
        public string Key;      // window handle, identifies the window
        public string Caption;  // text shown on the tab
        public string Ext;      // file extension, picks the tab icon
        public string Path;     // full file path, null while the file has never been saved
        public bool Modified;   // has unsaved changes
        public string Color;    // tab color name (see TabStrip.TabColors), null for none

        public bool SameAs(TabItem o)
        {
            return o != null && o.Key == Key && o.Caption == Caption && o.Ext == Ext &&
                   o.Path == Path && o.Modified == Modified && o.Color == Color;
        }
    }

    // Tab colors chosen from the tab menu. Saved files keep their color across sessions
    // (HKCU\Software\OfficeTabsLite\TabColors, one value per file path).
    internal static class TabColorStore
    {
        private const string KeyPath = TabsSettings.KeyPath + @"\TabColors";
        private static Dictionary<string, string> _byPath;
        private static readonly Dictionary<string, string> ByWindow = new Dictionary<string, string>();

        public static string Get(string windowKey, string path)
        {
            string c;
            if (path != null && Saved().TryGetValue(path, out c)) return c;
            return ByWindow.TryGetValue(windowKey, out c) ? c : null;
        }

        public static void Set(string windowKey, string path, string color)
        {
            if (color == null) ByWindow.Remove(windowKey); else ByWindow[windowKey] = color;
            if (path == null) return;
            if (color == null) Saved().Remove(path); else Saved()[path] = color;
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(KeyPath))
                    if (color == null) k.DeleteValue(path, false); else k.SetValue(path, color);
            }
            catch { }
        }

        // The window is gone: its handle may be reused by a new window.
        public static void Forget(string windowKey)
        {
            ByWindow.Remove(windowKey);
        }

        private static Dictionary<string, string> Saved()
        {
            if (_byPath != null) return _byPath;
            _byPath = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                    if (k != null)
                        foreach (string name in k.GetValueNames())
                            _byPath[name] = Convert.ToString(k.GetValue(name));
            }
            catch { }
            return _byPath;
        }
    }

    // ---- Settings shared by the add-in and the Center app (HKCU\Software\OfficeTabsLite) ----

    public enum TabStyle { Rounded, Trapezoid, Flat }

    public class TabsSettings
    {
        public const string KeyPath = @"Software\OfficeTabsLite";
        public static readonly string[] Apps = { "Word", "Excel", "PowerPoint" };
        private const string AddinKey = @"Software\Microsoft\Office\{0}\Addins\OfficeTabsLite.Connect";

        public TabStyle Style = TabStyle.Rounded;
        public bool GroupTaskbar = true;      // one taskbar button, one window place for all files
        public bool ShowIcons = true;         // file-type icon on each tab
        public bool NewOnDoubleClick = true;  // double-click empty strip space opens a new file
        public bool HideSingleTab = false;    // hide the strip while only one file is open

        public static TabsSettings Load()
        {
            var s = new TabsSettings();
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (k == null) return s;
                    TabStyle style;
                    if (Enum.TryParse(Convert.ToString(k.GetValue("TabStyle", "Rounded")), true, out style)) s.Style = style;
                    s.GroupTaskbar = Convert.ToInt32(k.GetValue("GroupTaskbar", 1)) != 0;
                    s.ShowIcons = Convert.ToInt32(k.GetValue("ShowIcons", 1)) != 0;
                    s.NewOnDoubleClick = Convert.ToInt32(k.GetValue("NewOnDoubleClick", 1)) != 0;
                    s.HideSingleTab = Convert.ToInt32(k.GetValue("HideSingleTab", 0)) != 0;
                }
            }
            catch { }
            return s;
        }

        public void Save()
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(KeyPath))
            {
                k.SetValue("TabStyle", Style.ToString());
                k.SetValue("GroupTaskbar", GroupTaskbar ? 1 : 0, RegistryValueKind.DWord);
                k.SetValue("ShowIcons", ShowIcons ? 1 : 0, RegistryValueKind.DWord);
                k.SetValue("NewOnDoubleClick", NewOnDoubleClick ? 1 : 0, RegistryValueKind.DWord);
                k.SetValue("HideSingleTab", HideSingleTab ? 1 : 0, RegistryValueKind.DWord);
            }
        }

        public bool SameAs(TabsSettings o)
        {
            return o != null && o.Style == Style && o.GroupTaskbar == GroupTaskbar &&
                   o.ShowIcons == ShowIcons && o.NewOnDoubleClick == NewOnDoubleClick &&
                   o.HideSingleTab == HideSingleTab;
        }

        // Whether Office loads the add-in at startup for this app; null if it isn't installed for it.
        public static bool? IsEnabledFor(string app)
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(string.Format(AddinKey, app)))
            {
                if (k == null) return null;
                return Convert.ToInt32(k.GetValue("LoadBehavior", 0)) == 3;
            }
        }

        public static void SetEnabledFor(string app, bool enabled)
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(string.Format(AddinKey, app), true))
                if (k != null) k.SetValue("LoadBehavior", enabled ? 3 : 0, RegistryValueKind.DWord);
        }

        // Undoes what Office may have done to the add-in for one app: turns it back on and
        // clears the "disabled items" / "crashing add-ins" records that keep it from loading,
        // and asks Office never to disable it automatically for being slow.
        public static void ResetApp(string app)
        {
            SetEnabledFor(app, true);
            using (RegistryKey office = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office"))
            {
                if (office == null) return;
                foreach (string version in office.GetSubKeyNames())
                {
                    double ignored;
                    if (!double.TryParse(version, System.Globalization.NumberStyles.Float,
                                         System.Globalization.CultureInfo.InvariantCulture, out ignored)) continue;
                    string res = version + @"\" + app + @"\Resiliency";
                    foreach (string list in new[] { "DisabledItems", "CrashingAddinList" })
                        using (RegistryKey k = office.OpenSubKey(res + @"\" + list, true))
                        {
                            if (k == null) continue;
                            foreach (string name in k.GetValueNames())
                            {
                                byte[] data = k.GetValue(name) as byte[];
                                string text = data == null ? Convert.ToString(k.GetValue(name))
                                                           : System.Text.Encoding.Unicode.GetString(data);
                                if (text.IndexOf("officetabslite", StringComparison.OrdinalIgnoreCase) >= 0)
                                    k.DeleteValue(name, false);
                            }
                        }
                    using (RegistryKey k = Registry.CurrentUser.CreateSubKey(
                               @"Software\Microsoft\Office\" + res + @"\DoNotDisableAddinList"))
                        k.SetValue("OfficeTabsLite.Connect", 1, RegistryValueKind.DWord);
                }
            }
        }
    }

    // ---- The add-in ----

    [ComVisible(true)]
    [Guid("B3D9F2A4-5C7E-4A1B-9E6D-2F8A1C0B3D44")]
    [ProgId("OfficeTabsLite.Connect")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        internal const string StripProgId = "OfficeTabsLite.TabStrip";
        internal static Connect Current;

        private object _app;          // the host Application (Word/Excel/PowerPoint)
        private object _addInInst;    // the COMAddIn instance
        private dynamic _ribbon;      // IRibbonUI callback object
        private ICTPFactory _factory; // creates the per-window tab strip panes
        private Timer _timer;         // keeps the tab strips in sync with open windows
        private bool _barVisible = true;
        private readonly Dictionary<string, CustomTaskPane> _panes = new Dictionary<string, CustomTaskPane>();
        private readonly HashSet<string> _failed = new HashSet<string>();
        private readonly List<string> _order = new List<string>();   // tab order, by window key
        private readonly Dictionary<string, int> _missing = new Dictionary<string, int>();   // ticks a key has been gone
        private string _activeKey;                                     // window shown last sync
        private int _lastWindowCount;
        private int _tick;
        private Native.WINDOWPLACEMENT _refPlacement;                  // place of the window shown before
        private int _settleTicks;
        private int _settingsTick;
        private Native.ITaskbarList _taskbar;
        private readonly HashSet<IntPtr> _taskbarHidden = new HashSet<IntPtr>();
        private readonly HashSet<IntPtr> _noAnimation = new HashSet<IntPtr>();
        private List<string> _labels = new List<string>();
        private TabsSettings _settings = new TabsSettings();
        private string _defaultExt = ".docx";   // extension for unsaved files of this app
        private Color _accent = TabStrip.WordAccent;

        internal string DefaultExt { get { return _defaultExt; } }

        // ---------- IDTExtensibility2 ----------
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _app = Application;
            _addInInst = AddInInst;
            Current = this;

            string name = "";
            try { name = (string)((dynamic)Application).Name; } catch { }
            if (name.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0) _defaultExt = ".xlsx";
            else if (name.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0) _defaultExt = ".pptx";
            _accent = TabStrip.AccentFor(_defaultExt);
            _settings = TabsSettings.Load();
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            StopTimer();
            RemoveAllPanes();
            RestoreTaskbar();
            _factory = null;
            _ribbon = null;
            _addInInst = null;
            _app = null;
            if (Current == this) Current = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { StopTimer(); }

        // ---------- ICustomTaskPaneConsumer ----------
        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            _factory = CTPFactoryInst;
            if (_timer == null)
            {
                _timer = new Timer();
                _timer.Interval = 200;
                _timer.Tick += delegate { try { Sync(); } catch { } };
                _timer.Start();
            }
        }

        // ---------- IRibbonExtensibility ----------
        public string GetCustomUI(string RibbonID)
        {
            return
@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='otlTab' label='Tabs'>
        <group id='otlBarGroup' label='Tab Bar'>
          <toggleButton id='otlBar' label='Tab Bar' size='large' imageMso='WindowSwitchWindowsMenuWord'
                        getPressed='GetBarPressed' onAction='OnToggleBar' />
          <button id='otlSettings' label='Settings' size='large' imageMso='PropertySheet' onAction='OnSettings' />
        </group>
        <group id='otlGroup' label='Open Documents'>
          <dropDown id='otlDocs' label='Document'
                    getItemCount='GetDocCount'
                    getItemLabel='GetDocLabel'
                    getSelectedItemIndex='GetSelectedDoc'
                    onAction='OnSelectDoc'
                    sizeString='WWWWWWWWWWWWWWWWWWWWWWWWWW' />
          <button id='otlPrev' label='Previous' imageMso='ReviewPreviousComment' onAction='OnPrev' />
          <button id='otlNext' label='Next' imageMso='ReviewNextComment' onAction='OnNext' />
          <button id='otlRefresh' label='Refresh' imageMso='Refresh' onAction='OnRefresh' />
          <button id='otlNew' label='New' imageMso='FileNew' onAction='OnNew' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // ---------- Ribbon callbacks ----------
        public void OnRibbonLoad(object ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public bool GetBarPressed(object control)
        {
            return _barVisible;
        }

        public void OnToggleBar(object control, bool pressed)
        {
            _barVisible = pressed;
            if (pressed) Sync(); else { RemoveAllPanes(); RestoreTaskbar(); }
        }

        // Opens the Center app, installed next to this DLL.
        public void OnSettings(object control)
        {
            try
            {
                string dir = Path.GetDirectoryName(typeof(Connect).Assembly.Location);
                Process.Start(Path.Combine(dir, "OfficeTabsLite-Center.exe"));
            }
            catch { }
        }

        public int GetDocCount(object control)
        {
            _labels = CollectLabels();
            return _labels.Count;
        }

        public string GetDocLabel(object control, int index)
        {
            if (index >= 0 && index < _labels.Count) return _labels[index];
            return "(document)";
        }

        public int GetSelectedDoc(object control)
        {
            return ActiveIndex();
        }

        public void OnSelectDoc(object control, string selectedId, int selectedIndex)
        {
            ActivateAt(selectedIndex);
            Invalidate();
        }

        public void OnPrev(object control) { Step(-1); }
        public void OnNext(object control) { Step(+1); }
        public void OnRefresh(object control) { Invalidate(); }

        public void OnNew(object control)
        {
            try { ((dynamic)_app).Documents.Add(); }
            catch {
                try { ((dynamic)_app).Workbooks.Add(); }
                catch {
                    try { ((dynamic)_app).Presentations.Add(); } catch { }
                }
            }
            Invalidate();
        }

        // ---------- tab strip (called by TabStrip) ----------
        internal void ActivateTab(string key)
        {
            object w = FindWindow(key);
            if (w == null) return;
            if (_barVisible && _activeKey != null && _activeKey != key)
                CopyPlacement(RootOf(_activeKey), RootOf(key));
            try { ((dynamic)w).Activate(); } catch { }
            Sync();
        }

        internal void CloseTab(string key)
        {
            object w = FindWindow(key);
            if (w == null) return;
            // Close through Office's own command so it asks about unsaved changes as usual.
            try
            {
                ((dynamic)w).Activate();
                ((dynamic)_app).CommandBars.ExecuteMso("FileClose");
            }
            catch { }
            Sync();
        }

        // Office draws a new window blank (no ribbon, a loading spinner) while it builds it,
        // with a fade-in on top. A snapshot of the current window covers that and fades out
        // once the new window's tab strip is ready.
        internal void NewDocument()
        {
            FreezeFrame cover = null;
            if (_settings.GroupTaskbar && _activeKey != null) cover = FreezeFrame.Cover(RootOf(_activeKey));
            OnNew(null);
            Sync();
            if (cover != null) cover.RevealWhenReady();
        }

        // ---------- tab strip sync ----------
        private void Sync()
        {
            if (_factory == null || _app == null) return;

            // Pick up changes made in the Center app (every ~5 s).
            if (++_settingsTick % 25 == 0)
            {
                TabsSettings latest = TabsSettings.Load();
                if (!latest.SameAs(_settings))
                {
                    if (_settings.GroupTaskbar && !latest.GroupTaskbar) RestoreTaskbar();
                    _settings = latest;
                }
            }

            var items = new List<TabItem>();
            var wins = new Dictionary<string, object>();
            foreach (object w in GetWindows())
            {
                string key = KeyOf(w);
                if (key == null || wins.ContainsKey(key)) continue;
                wins[key] = w;
                var t = new TabItem();
                t.Key = key;
                t.Caption = CaptionOf(w);
                t.Ext = ExtOf(w, _defaultExt);
                dynamic doc = DocOf(w);
                t.Path = PathOf(doc);
                t.Modified = IsModified(doc);
                t.Color = TabColorStore.Get(key, t.Path);
                items.Add(t);
            }

            // Keep tabs in a stable order (Excel lists windows most-recently-active first,
            // which would reshuffle the tabs on every switch): known tabs keep their place,
            // new ones go to the end. A key that drops out for a moment keeps its place; one
            // gone for over a second is forgotten (with its color), because Windows reuses window
            // handles and a new window must not inherit a closed one's place or color.
            foreach (string k in new List<string>(_order))
            {
                if (wins.ContainsKey(k)) { _missing.Remove(k); continue; }
                int gone;
                _missing.TryGetValue(k, out gone);
                if (++gone > 5) { _order.Remove(k); _missing.Remove(k); TabColorStore.Forget(k); }
                else _missing[k] = gone;
            }
            List<TabItem> fresh = items.FindAll(t => !_order.Contains(t.Key));
            fresh.Sort((a, b) => string.Compare(a.Caption, b.Caption, StringComparison.CurrentCultureIgnoreCase));
            foreach (TabItem t in fresh) _order.Add(t.Key);
            items.Sort((a, b) => _order.IndexOf(a.Key).CompareTo(_order.IndexOf(b.Key)));

            foreach (string key in new List<string>(_panes.Keys))
                if (!wins.ContainsKey(key)) RemovePane(key);

            if (!_barVisible) return;

            if (_settings.GroupTaskbar) GroupWindows(wins);

            foreach (KeyValuePair<string, object> kv in wins)
            {
                CustomTaskPane pane;
                if (!_panes.TryGetValue(kv.Key, out pane))
                {
                    if (_failed.Contains(kv.Key)) continue;
                    pane = CreatePane(kv.Value);
                    if (pane == null) { _failed.Add(kv.Key); continue; }
                    _panes[kv.Key] = pane;
                }
                TabStrip strip = StripOf(pane);
                if (strip == null) continue;
                FitHeight(pane, strip);
                strip.Configure(_settings, _accent);
                strip.SetTabs(items, kv.Key);
                bool show = !(_settings.HideSingleTab && wins.Count == 1);
                try { if (pane.Visible != show) pane.Visible = show; } catch { }
            }
        }

        // ---------- document helpers ----------
        // The document behind a window: Word Window.Document, PowerPoint Window.Presentation,
        // Excel Window.Parent (the workbook).
        private dynamic DocOf(dynamic w)
        {
            try
            {
                if (_defaultExt == ".docx") return w.Document;
                if (_defaultExt == ".pptx") return w.Presentation;
                return w.Parent;
            }
            catch { return null; }
        }

        // Full path of a saved file; null while it has never been saved.
        private static string PathOf(dynamic doc)
        {
            try
            {
                if (doc == null || string.IsNullOrEmpty((string)doc.Path)) return null;
                return (string)doc.FullName;
            }
            catch { return null; }
        }

        // Saved is a bool in Word and Excel and an MsoTriState (-1 = true) in PowerPoint.
        private static bool IsModified(dynamic doc)
        {
            try
            {
                if (doc == null) return false;
                object saved = doc.Saved;
                return saved is bool ? !(bool)saved : Convert.ToInt32(saved) == 0;
            }
            catch { return false; }
        }

        // ---------- tab menu actions (called by TabStrip) ----------
        internal void SaveTab(string key)
        {
            object w = FindWindow(key);
            if (w == null) return;
            dynamic doc = DocOf(w);
            if (PathOf(doc) != null)
            {
                try { doc.Save(); } catch { }
            }
            else
            {
                // Never saved: let Office show its own Save As.
                try { ((dynamic)w).Activate(); ((dynamic)_app).CommandBars.ExecuteMso("FileSave"); } catch { }
            }
            Sync();
        }

        internal void SaveAll()
        {
            foreach (object w in GetWindows())
            {
                string key = KeyOf(w);
                if (key != null && IsModified(DocOf(w))) SaveTab(key);
            }
        }

        internal void CloseTabs(List<string> keys)
        {
            foreach (string key in keys) CloseTab(key);
        }

        // New tab order from a drag: the dragged strip's visible order goes first.
        internal void ReorderTabs(List<string> keys)
        {
            _order.RemoveAll(k => keys.Contains(k));
            _order.InsertRange(0, keys);
            Sync();
        }

        internal void SetTabColor(string key, string path, string color)
        {
            TabColorStore.Set(key, path, color);
            Sync();
        }

        // Make the document windows behave like one tabbed window: the active window takes
        // the previous one's place and size, only it keeps a taskbar button, and minimizing it
        // minimizes the whole group. Windows are never hidden, so nothing can get lost.
        private void GroupWindows(Dictionary<string, object> wins)
        {
            // Switching tabs should be instant: no zoom/fade animation when a window is
            // shown, maximized or restored.
            foreach (string key in wins.Keys) DisableAnimation(RootOf(key));

            // Minimizing the shown window makes Windows activate the next document window;
            // instead, minimize the whole group and keep the shown window as the active one.
            IntPtr shownRoot = (_activeKey != null && wins.ContainsKey(_activeKey)) ? RootOf(_activeKey) : IntPtr.Zero;
            if (shownRoot != IntPtr.Zero && Native.IsIconic(shownRoot))
            {
                bool minimizedNow = false;
                foreach (string key in wins.Keys)
                {
                    IntPtr root = RootOf(key);
                    if (root != IntPtr.Zero && root != shownRoot && !Native.IsIconic(root))
                    {
                        Native.ShowWindow(root, Native.SW_SHOWMINNOACTIVE);
                        minimizedNow = true;
                    }
                }
                if (minimizedNow || ++_tick % 25 == 0)
                    foreach (string key in wins.Keys)
                    {
                        IntPtr root = RootOf(key);
                        if (root != IntPtr.Zero) SetTaskbarButton(root, root == shownRoot);
                    }
                return;
            }

            string active = null;
            try { active = KeyOf(((dynamic)_app).ActiveWindow); } catch { }
            if (active == null || !wins.ContainsKey(active)) return;
            IntPtr activeRoot = RootOf(active);
            if (activeRoot == IntPtr.Zero) return;

            // The newly shown window takes the previous one's place. Office may still resize a
            // window right after it appears (e.g. leaving the start screen), so keep re-applying
            // that place for a few seconds.
            bool changed = active != _activeKey || wins.Count != _lastWindowCount;
            if (active != _activeKey && shownRoot != IntPtr.Zero && GetPlacement(shownRoot, out _refPlacement))
                _settleTicks = 15;
            if (_settleTicks > 0)
            {
                _settleTicks--;
                ApplyPlacement(activeRoot, _refPlacement);
            }
            _activeKey = active;
            _lastWindowCount = wins.Count;

            if (!changed && ++_tick % 25 != 0) return;
            foreach (string key in wins.Keys)
            {
                IntPtr root = RootOf(key);
                if (root != IntPtr.Zero) SetTaskbarButton(root, root == activeRoot);
            }
        }

        private void DisableAnimation(IntPtr root)
        {
            if (root == IntPtr.Zero || !_noAnimation.Add(root)) return;
            int on = 1;
            try { Native.DwmSetWindowAttribute(root, Native.DWMWA_TRANSITIONS_FORCEDISABLED, ref on, 4); } catch { }
        }

        private static IntPtr RootOf(string key)
        {
            long value;
            if (key == null || !long.TryParse(key, out value)) return IntPtr.Zero;
            var hwnd = new IntPtr(value);
            if (!Native.IsWindow(hwnd)) return IntPtr.Zero;
            IntPtr root = Native.GetAncestor(hwnd, Native.GA_ROOT);
            return root == IntPtr.Zero ? hwnd : root;
        }

        private static void CopyPlacement(IntPtr from, IntPtr to)
        {
            Native.WINDOWPLACEMENT p;
            if (from != to && GetPlacement(from, out p)) ApplyPlacement(to, p);
        }

        // Reads a window's placement as "normal" or "maximized" (a minimized window reports
        // the state it would restore to).
        private static bool GetPlacement(IntPtr hwnd, out Native.WINDOWPLACEMENT p)
        {
            p = new Native.WINDOWPLACEMENT();
            p.length = Marshal.SizeOf(p);
            if (hwnd == IntPtr.Zero || !Native.GetWindowPlacement(hwnd, ref p)) return false;
            if (p.showCmd != Native.SW_SHOWMAXIMIZED)
                p.showCmd = (p.showCmd == Native.SW_SHOWMINIMIZED && (p.flags & Native.WPF_RESTORETOMAXIMIZED) != 0)
                    ? Native.SW_SHOWMAXIMIZED : Native.SW_SHOWNORMAL;
            return true;
        }

        private static void ApplyPlacement(IntPtr hwnd, Native.WINDOWPLACEMENT p)
        {
            Native.WINDOWPLACEMENT current;
            if (hwnd == IntPtr.Zero) return;
            if (GetPlacement(hwnd, out current) && current.showCmd == p.showCmd &&
                current.rcNormalPosition.Equals(p.rcNormalPosition)) return;
            Native.SetWindowPlacement(hwnd, ref p);
        }

        private void SetTaskbarButton(IntPtr root, bool show)
        {
            if (_taskbar == null)
            {
                try { _taskbar = (Native.ITaskbarList)new Native.TaskbarList(); _taskbar.HrInit(); }
                catch { _taskbar = null; return; }
            }
            try
            {
                if (show)
                {
                    if (_taskbarHidden.Remove(root)) _taskbar.AddTab(root);
                }
                else
                {
                    _taskbar.DeleteTab(root);
                    _taskbarHidden.Add(root);
                }
            }
            catch { }
        }

        private void RestoreTaskbar()
        {
            if (_taskbar != null)
                foreach (IntPtr root in _taskbarHidden)
                    if (Native.IsWindow(root)) { try { _taskbar.AddTab(root); } catch { } }
            _taskbarHidden.Clear();
            _activeKey = null;
        }

        // The pane height includes Office's own title bar and its units vary with DPI, so nudge
        // it a little on each sync until the strip itself is StripHeight pixels tall.
        // Once one pane has the right height, new panes start at that height, so a new window's
        // strip appears complete instead of resizing (and flickering) over the first ticks.
        private static int _fittedPaneHeight;

        private static void FitHeight(CustomTaskPane pane, TabStrip strip)
        {
            // A brand-new pane has not been laid out yet (height 0); resizing it then would
            // overshoot and make the strip jump on the next tick.
            if (strip.Height == 0) return;
            int diff = TabStrip.StripHeight - strip.Height;
            if (Math.Abs(diff) <= 2)
            {
                try { if (_fittedPaneHeight == 0) _fittedPaneHeight = pane.Height; } catch { }
                return;
            }
            try { pane.Height = Math.Max(20, pane.Height + diff); } catch { }
        }

        private CustomTaskPane CreatePane(object window)
        {
            try
            {
                CustomTaskPane p = _factory.CreateCTP(StripProgId, "Tabs", window);
                p.DockPosition = MsoCTPDockPosition.msoCTPDockPositionTop;
                p.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                p.Height = _fittedPaneHeight > 0 ? _fittedPaneHeight : 64;
                p.Visible = true;
                return p;
            }
            catch { return null; }
        }

        private static TabStrip StripOf(CustomTaskPane pane)
        {
            try { return pane.ContentControl as TabStrip; } catch { return null; }
        }

        private void RemovePane(string key)
        {
            CustomTaskPane pane;
            if (_panes.TryGetValue(key, out pane))
            {
                try { pane.Delete(); } catch { }
                _panes.Remove(key);
            }
        }

        private void RemoveAllPanes()
        {
            foreach (string key in new List<string>(_panes.Keys)) RemovePane(key);
            _failed.Clear();
        }

        private void StopTimer()
        {
            if (_timer != null) { _timer.Stop(); _timer.Dispose(); _timer = null; }
        }

        // ---------- helpers (all defensive; must never throw into Office) ----------
        private List<object> GetWindows()
        {
            var list = new List<object>();
            try
            {
                dynamic app = _app;
                dynamic wins = app.Windows;
                int n = (int)wins.Count;
                for (int i = 1; i <= n; i++) list.Add(wins[i]);
            }
            catch { }
            return list;
        }

        private object FindWindow(string key)
        {
            foreach (object w in GetWindows())
                if (KeyOf(w) == key) return w;
            return null;
        }

        private static string KeyOf(dynamic w)
        {
            try { long h = Convert.ToInt64(w.Hwnd); if (h > 0) return h.ToString(); } catch { }
            // PowerPoint's document windows expose no handle: find their frame window
            // (class PPTFrameClass, title "<caption> - PowerPoint") in this process.
            string caption = null;
            try { caption = CaptionOf(w); } catch { }
            if (caption == null) return null;
            // No frame yet (window still opening): skip it this round rather than giving it a
            // temporary key, which would make its tab jump and its pane be recreated.
            IntPtr frame = FrameByCaption(caption);
            return frame != IntPtr.Zero ? frame.ToInt64().ToString() : null;
        }

        private static readonly uint OwnPid = (uint)Process.GetCurrentProcess().Id;

        private static IntPtr FrameByCaption(string caption)
        {
            IntPtr found = IntPtr.Zero;
            var text = new System.Text.StringBuilder(512);
            Native.EnumWindows(delegate (IntPtr h, IntPtr l)
            {
                uint pid;
                Native.GetWindowThreadProcessId(h, out pid);
                if (pid != OwnPid) return true;
                // Only the document frame: PowerPoint also has hidden helper windows with
                // the same title, which would make the key flip between windows.
                text.Length = 0;
                Native.GetClassName(h, text, text.Capacity);
                if (text.ToString() != "PPTFrameClass") return true;
                text.Length = 0;
                Native.GetWindowText(h, text, text.Capacity);
                string t = text.ToString();
                if (t == caption || t.StartsWith(caption + " - ", StringComparison.Ordinal)) { found = h; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        // Extension of the window's file (Word: Document, PowerPoint: Presentation,
        // Excel: the window's parent Workbook); unsaved files get the app's default.
        private static string ExtOf(dynamic w, string fallback)
        {
            string name = null;
            try { name = (string)w.Document.Name; } catch { }
            if (name == null) try { name = (string)w.Presentation.Name; } catch { }
            if (name == null) try { name = (string)w.Parent.Name; } catch { }
            string ext = null;
            try { ext = Path.GetExtension(name ?? ""); } catch { }
            return string.IsNullOrEmpty(ext) ? fallback : ext.ToLowerInvariant();
        }

        private List<string> CollectLabels()
        {
            var labels = new List<string>();
            foreach (var w in GetWindows()) labels.Add(CaptionOf(w));
            return labels;
        }

        private static string CaptionOf(dynamic w)
        {
            try { string c = (string)w.Caption; if (!string.IsNullOrEmpty(c)) return c; } catch { }
            try { return (string)w.Presentation.Name; } catch { }
            try { return (string)w.Document.Name; } catch { }
            return "(document)";
        }

        private int ActiveIndex()
        {
            try
            {
                dynamic app = _app;
                string active = CaptionOf(app.ActiveWindow);
                var wins = GetWindows();
                for (int i = 0; i < wins.Count; i++)
                    if (CaptionOf(wins[i]) == active) return i;
            }
            catch { }
            return 0;
        }

        private void ActivateAt(int index)
        {
            try
            {
                var wins = GetWindows();
                if (index >= 0 && index < wins.Count)
                    ((dynamic)wins[index]).Activate();
            }
            catch { }
        }

        private void Step(int delta)
        {
            var wins = GetWindows();
            int n = wins.Count;
            if (n == 0) return;
            int cur = ActiveIndex();
            int next = ((cur + delta) % n + n) % n;
            ActivateAt(next);
            Invalidate();
        }

        private void Invalidate()
        {
            try { if (_ribbon != null) _ribbon.Invalidate(); } catch { }
        }
    }

    // ---- Snapshot that hides a new window while Office builds it ----

    // A topmost, click-through, non-activating layered window showing a picture of the current
    // document window. Being layered, Windows shows the picture itself from the first frame
    // (no unpainted flash) and fades it by alpha alone. It always goes away within MaxMs.
    internal sealed class FreezeFrame : NativeWindow
    {
        private const int MinMs = 250;     // keep at least this long (Office starts drawing late)
        private const int SettleMs = 40;   // after the new strip appears, let Office finish painting
        private const int FadeMs = 120;
        private const int MaxMs = 2000;
        private readonly DateTime _start = DateTime.UtcNow;
        private Timer _timer;
        private DateTime _fadeStart = DateTime.MinValue;
        private IntPtr _bitmap;            // GDI bitmap with the picture
        private Rectangle _bounds;

        [StructLayout(LayoutKind.Sequential)] private struct SIZE { public int cx, cy; }
        [StructLayout(LayoutKind.Sequential)]
        private struct BLENDFUNCTION { public byte BlendOp, BlendFlags, SourceConstantAlpha, AlphaFormat; }

        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd, out Native.RECT rect);
        [DllImport("user32.dll")] private static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, ref Native.POINT dst,
            ref SIZE size, IntPtr hdcSrc, ref Native.POINT src, int key, ref BLENDFUNCTION blend, int flags);
        [DllImport("user32.dll")] private static extern IntPtr GetDC(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr hwnd, IntPtr dc);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateCompatibleDC(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern bool DeleteDC(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr dc, IntPtr obj);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr obj);

        public static FreezeFrame Cover(IntPtr window)
        {
            try
            {
                Native.RECT r;
                if (window == IntPtr.Zero || Native.IsIconic(window) || !GetWindowRect(window, out r)) return null;
                var bounds = Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom);
                // only the part on screen can be captured
                bounds.Intersect(SystemInformation.VirtualScreen);
                if (bounds.Width < 50 || bounds.Height < 50) return null;
                IntPtr hbitmap;
                using (var shot = new Bitmap(bounds.Width, bounds.Height))
                {
                    using (Graphics g = Graphics.FromImage(shot))
                        g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size);
                    hbitmap = shot.GetHbitmap();
                }
                var f = new FreezeFrame();
                f.Open(bounds, hbitmap);
                return f;
            }
            catch { return null; }
        }

        private void Open(Rectangle bounds, IntPtr hbitmap)
        {
            const int WS_POPUP = unchecked((int)0x80000000);
            const int WS_EX_TOPMOST = 0x8, WS_EX_TRANSPARENT = 0x20, WS_EX_TOOLWINDOW = 0x80,
                      WS_EX_LAYERED = 0x80000, WS_EX_NOACTIVATE = 0x08000000;
            _bitmap = hbitmap;
            _bounds = bounds;
            var cp = new CreateParams();
            cp.Caption = "";
            cp.X = bounds.X; cp.Y = bounds.Y; cp.Width = bounds.Width; cp.Height = bounds.Height;
            cp.Style = WS_POPUP;
            cp.ExStyle = WS_EX_TOPMOST | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_LAYERED | WS_EX_NOACTIVATE;
            CreateHandle(cp);
            Render(255);                                   // picture first ...
            Native.ShowWindow(Handle, 4 /* SW_SHOWNOACTIVATE */);   // ... then show it
        }

        private void Render(byte alpha)
        {
            IntPtr screen = GetDC(IntPtr.Zero);
            IntPtr mem = CreateCompatibleDC(screen);
            IntPtr old = SelectObject(mem, _bitmap);
            try
            {
                var dst = new Native.POINT(); dst.X = _bounds.X; dst.Y = _bounds.Y;
                var size = new SIZE(); size.cx = _bounds.Width; size.cy = _bounds.Height;
                var src = new Native.POINT();
                var blend = new BLENDFUNCTION(); blend.SourceConstantAlpha = alpha;   // AC_SRC_OVER, constant alpha
                UpdateLayeredWindow(Handle, screen, ref dst, ref size, mem, ref src, 0, ref blend, 2 /* ULW_ALPHA */);
            }
            finally
            {
                SelectObject(mem, old);
                DeleteDC(mem);
                ReleaseDC(IntPtr.Zero, screen);
            }
        }

        public void RevealWhenReady()
        {
            _timer = new Timer();
            _timer.Interval = 15;
            _timer.Tick += delegate { try { Step(); } catch { Finish(); } };
            _timer.Start();
        }

        private void Step()
        {
            double age = (DateTime.UtcNow - _start).TotalMilliseconds;
            if (_fadeStart == DateTime.MinValue)
            {
                bool stripReady = TabStrip.LastStripShown > _start &&
                                  (DateTime.UtcNow - TabStrip.LastStripShown).TotalMilliseconds >= SettleMs;
                if ((age >= MinMs && stripReady) || age >= MaxMs) _fadeStart = DateTime.UtcNow;
                return;
            }
            double t = (DateTime.UtcNow - _fadeStart).TotalMilliseconds / FadeMs;
            if (t >= 1) { Finish(); return; }
            Render((byte)(255 * (1 - t)));
        }

        private void Finish()
        {
            if (_timer != null) { _timer.Stop(); _timer.Dispose(); _timer = null; }
            try { DestroyHandle(); } catch { }
            if (_bitmap != IntPtr.Zero) { DeleteObject(_bitmap); _bitmap = IntPtr.Zero; }
        }
    }

    // ---- The tab strip shown at the top of each document window ----

    [ComVisible(true)]
    [Guid("6E0C2D8A-3F41-4B7C-A9D2-5B1E7F3C8A61")]
    [ProgId("OfficeTabsLite.TabStrip")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TabStrip : UserControl
    {
        internal const int StripHeight = 34;

        public static readonly Color WordAccent = Color.FromArgb(43, 87, 154);
        public static readonly Color ExcelAccent = Color.FromArgb(33, 115, 70);
        public static readonly Color PowerPointAccent = Color.FromArgb(196, 67, 37);

        private static readonly Color StripTop    = Color.FromArgb(242, 242, 242);
        private static readonly Color StripBottom = Color.FromArgb(230, 230, 230);
        private static readonly Color Border      = Color.FromArgb(188, 188, 188);
        private static readonly Color CloseHover  = Color.FromArgb(232, 17, 35);
        private const int Slant = 12;   // width of the slanted edge of trapezoid tabs
        private static readonly Dictionary<string, Bitmap> IconCache = new Dictionary<string, Bitmap>();

        // New-tab animation: a tab grows from nothing to full width over GrowMs after its window
        // first appears. The first-seen times are shared by every strip in the Office process,
        // so the animation also plays in the strip of the newly opened window.
        private const double GrowMs = 280;
        private static readonly Dictionary<string, DateTime> FirstSeen = new Dictionary<string, DateTime>();
        private readonly Timer _animTimer = new Timer();
        private bool _shownFull;   // this strip has been painted at its full height at least once

        // "+" plays its animation in the strip the user is looking at (a placeholder tab grows
        // in) and only then opens the file, whose window appears with its tab already in place.
        // Tabs first seen before NoAnimUntil are shown at full width straight away.
        private static DateTime NoAnimUntil = DateTime.MinValue;
        private TabItem _placeholder;

        // drag to reorder
        private string _dragKey;     // tab under the left button, may become a drag
        private Point _dragStart;
        private bool _dragging;
        private int _dropIndex = -1; // slot the dragged tab would land in

        // When a strip was last painted complete for the first time (a new window became ready).
        internal static DateTime LastStripShown = DateTime.MinValue;

        private List<TabItem> _items = new List<TabItem>();
        private string _mine;
        private TabStyle _style = TabStyle.Rounded;
        private bool _showIcons = true;
        private bool _newOnDoubleClick = true;
        private bool _preview;          // shown inside the Center app: clicks only move the selection
        private Color _accent = WordAccent;
        private readonly List<Rectangle> _tabRects = new List<Rectangle>();
        private readonly List<Rectangle> _closeRects = new List<Rectangle>();
        private Rectangle _plusRect;
        private int _hoverTab = -1;
        private bool _hoverClose;
        private bool _hoverPlus;
        private readonly Font _tabFont = new Font("Segoe UI", 9f);
        private readonly Font _activeFont = new Font("Segoe UI", 9f, FontStyle.Bold);
        private readonly ToolTip _tip = new ToolTip();
        private int _tipTab = -1;

        public TabStrip()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            BackColor = StripBottom;
            _animTimer.Interval = 15;
            _animTimer.Tick += delegate
            {
                try
                {
                    if (!AnyGrowing()) _animTimer.Stop();
                    Relayout();
                }
                catch { _animTimer.Stop(); }
            };
        }

        // 0..1 with an ease-out curve: fast at first, settling gently at full width.
        private static double GrowthOf(string key)
        {
            DateTime seen;
            if (key == null || !FirstSeen.TryGetValue(key, out seen)) return 1;
            double t = (DateTime.UtcNow - seen).TotalMilliseconds / GrowMs;
            if (t >= 1) return 1;
            if (t <= 0) return 0;
            return 1 - Math.Pow(1 - t, 3);
        }

        private bool AnyGrowing()
        {
            foreach (TabItem t in _items)
                if (GrowthOf(t.Key) < 1) return true;
            return false;
        }

        // Records new tabs and starts animating them.
        private void NoteNewTabs()
        {
            DateTime now = DateTime.UtcNow;
            DateTime start = now < NoAnimUntil ? DateTime.MinValue : now;
            foreach (TabItem t in _items)
                if (t.Key != null && !FirstSeen.ContainsKey(t.Key)) FirstSeen[t.Key] = start;
            if (FirstSeen.Count > 300)
                foreach (string k in new List<string>(FirstSeen.Keys))
                    if ((now - FirstSeen[k]).TotalMinutes > 1) FirstSeen.Remove(k);
            if (AnyGrowing() && !_animTimer.Enabled) _animTimer.Start();
        }

        // A new window's strip only becomes visible after Office has created and sized its
        // pane, which can take longer than the animation itself. So when a strip is first
        // shown at full height, tabs that appeared moments ago start growing from that moment.
        private void RestartRecentTabs()
        {
            DateTime now = DateTime.UtcNow;
            foreach (TabItem t in _items)
            {
                DateTime seen;
                if (t.Key != null && FirstSeen.TryGetValue(t.Key, out seen) && (now - seen).TotalMilliseconds < 2000)
                    FirstSeen[t.Key] = now;
            }
            if (AnyGrowing() && !_animTimer.Enabled) _animTimer.Start();
        }

        public static Color AccentFor(string ext)
        {
            ext = (ext ?? "").ToLowerInvariant();
            if (ext.StartsWith(".xl") || ext == ".csv") return ExcelAccent;
            if (ext.StartsWith(".pp") || ext.StartsWith(".pot")) return PowerPointAccent;
            return WordAccent;
        }

        internal void Configure(TabsSettings s, Color accent)
        {
            if (s.Style == _style && s.ShowIcons == _showIcons &&
                s.NewOnDoubleClick == _newOnDoubleClick && accent == _accent) return;
            _style = s.Style;
            _showIcons = s.ShowIcons;
            _newOnDoubleClick = s.NewOnDoubleClick;
            _accent = accent;
            Relayout();
        }

        // Used by the Center app to show a live sample of the chosen look.
        public void ShowPreview(string style, bool showIcons)
        {
            TabStyle parsed;
            _style = Enum.TryParse(style, true, out parsed) ? parsed : TabStyle.Rounded;
            _showIcons = showIcons;
            _preview = true;
            if (_items.Count == 0)
            {
                _items = new List<TabItem> { Sample("1", "Report", ".docx"), Sample("2", "Budget", ".xlsx"),
                                             Sample("3", "Lecture slides", ".pptx") };
                _items[1].Modified = true;   // shows the unsaved-changes dot
                _mine = "1";
                NoteNewTabs();
            }
            Relayout();
        }

        private static TabItem Sample(string key, string caption, string ext)
        {
            var t = new TabItem();
            t.Key = key; t.Caption = caption; t.Ext = ext;
            return t;
        }

        // Shows a "New ..." tab growing in, then asks Office for the new file once it is done.
        // host may be null (tests): then only the animation plays.
        private void AnimateThenNew(Connect host)
        {
            if (_placeholder != null) return;   // already opening one
            string ext = host != null ? host.DefaultExt : ".docx";
            Color accent = AccentFor(ext);
            _placeholder = new TabItem();
            _placeholder.Key = "pending:" + DateTime.UtcNow.Ticks;
            _placeholder.Caption = accent == ExcelAccent ? "New workbook"
                                 : accent == PowerPointAccent ? "New presentation" : "New document";
            _placeholder.Ext = ext;
            _items = new List<TabItem>(_items);
            _items.Add(_placeholder);
            _mine = _placeholder.Key;
            NoteNewTabs();
            Relayout();

            var wait = new Timer();
            wait.Interval = (int)GrowMs;
            wait.Tick += delegate
            {
                wait.Stop();
                wait.Dispose();
                NoAnimUntil = DateTime.UtcNow.AddMilliseconds(1500);
                _placeholder = null;
                if (host != null) { try { host.NewDocument(); } catch { } }
            };
            wait.Start();
        }

        internal void SetTabs(List<TabItem> items, string mine)
        {
            if (_placeholder != null)
            {
                // keep the growing placeholder until the real file has been opened
                items = new List<TabItem>(items);
                items.Add(_placeholder);
                mine = _placeholder.Key;
            }
            bool same = mine == _mine && items.Count == _items.Count;
            for (int i = 0; same && i < items.Count; i++)
                same = items[i].SameAs(_items[i]);
            if (same) return;
            _items = new List<TabItem>(items);
            _mine = mine;
            NoteNewTabs();
            Relayout();
        }

        // Recompute hit areas now (the mouse can act before the next paint) and repaint.
        private void Relayout()
        {
            _hoverTab = -1; _hoverClose = false; _tipTab = -1;
            try { using (Graphics g = CreateGraphics()) LayoutTabs(g); } catch { }
            Invalidate();
        }

        // ---------- layout ----------
        private void LayoutTabs(Graphics g)
        {
            _tabRects.Clear();
            _closeRects.Clear();
            int top = 5, height = Math.Max(16, Height - top);
            int slant = _style == TabStyle.Trapezoid ? Slant : 0;
            int extra = (_showIcons ? 24 : 0) + 44 + slant;
            int gap = 3;
            int available = Math.Max(0, Width - 12 - 40);
            int n = _items.Count;

            var widths = new int[n];
            int total = 0;
            for (int i = 0; i < n; i++)
            {
                Size s = TextRenderer.MeasureText(g, _items[i].Caption, _activeFont);
                widths[i] = Math.Max(110, Math.Min(250, s.Width + extra));
                total += widths[i] + gap;
            }
            if (n > 0 && total > available)
            {
                int each = Math.Max(70, available / n - gap);
                for (int i = 0; i < n; i++) widths[i] = each;
            }
            for (int i = 0; i < n; i++)
                widths[i] = Math.Max(1, (int)Math.Round(widths[i] * GrowthOf(_items[i].Key)));

            int x = 6;
            for (int i = 0; i < n; i++)
            {
                var r = new Rectangle(x, top, widths[i], height);
                _tabRects.Add(r);
                int closeRight = r.Right - 7 - slant;
                _closeRects.Add(new Rectangle(closeRight - 16, r.Top + (r.Height - 16) / 2, 16, 16));
                x += widths[i] + gap;
            }
            _plusRect = new Rectangle(x + 5, top + (height - 22) / 2, 28, 22);
        }

        // Outline of one tab; the bottom is left open so it merges into the strip.
        private GraphicsPath TabPath(Rectangle r)
        {
            var p = new GraphicsPath();
            int L = r.Left, T = r.Top, R = r.Right, B = r.Bottom;
            if (_style == TabStyle.Rounded)
            {
                const int rad = 8;
                p.AddLine(L, B, L, T + rad);
                p.AddArc(L, T, 2 * rad, 2 * rad, 180, 90);
                p.AddLine(L + rad, T, R - rad, T);
                p.AddArc(R - 2 * rad, T, 2 * rad, 2 * rad, 270, 90);
                p.AddLine(R, T + rad, R, B);
            }
            else if (_style == TabStyle.Trapezoid)
            {
                const int rad = 5;
                p.AddLine(L, B, L, T + rad);
                p.AddArc(L, T, 2 * rad, 2 * rad, 180, 90);
                p.AddLine(L + rad, T, R - Slant, T);
                p.AddLine(R - Slant, T, R, B);
            }
            else
            {
                p.AddLine(L, B, L, T);
                p.AddLine(L, T, R, T);
                p.AddLine(R, T, R, B);
            }
            return p;
        }

        private static Color Mix(Color a, Color b, double t)
        {
            return Color.FromArgb((int)(a.R + (b.R - a.R) * t), (int)(a.G + (b.G - a.G) * t),
                                  (int)(a.B + (b.B - a.B) * t));
        }

        // Small icon Windows shows for this file type (the real Word/Excel/PowerPoint icon).
        private static Bitmap IconFor(string ext)
        {
            if (string.IsNullOrEmpty(ext)) return null;
            Bitmap bmp;
            if (IconCache.TryGetValue(ext, out bmp)) return bmp;
            bmp = null;
            try
            {
                var info = new Native.SHFILEINFO();
                Native.SHGetFileInfo("file" + ext, Native.FILE_ATTRIBUTE_NORMAL, ref info, (uint)Marshal.SizeOf(info),
                                     Native.SHGFI_ICON | Native.SHGFI_SMALLICON | Native.SHGFI_USEFILEATTRIBUTES);
                if (info.hIcon != IntPtr.Zero)
                {
                    using (Icon icon = Icon.FromHandle(info.hIcon)) bmp = icon.ToBitmap();
                    Native.DestroyIcon(info.hIcon);
                }
            }
            catch { }
            IconCache[ext] = bmp;
            return bmp;
        }

        // ---------- painting ----------
        // Every handler below is guarded: an exception escaping a task pane control
        // would surface inside Office as a .NET error dialog.
        protected override void OnPaint(PaintEventArgs e)
        {
            try { PaintStrip(e.Graphics); } catch { }
        }

        private void PaintStrip(Graphics g)
        {
            if (!_shownFull && Height >= StripHeight - 4 && _items.Count > 0)
            {
                _shownFull = true;
                LastStripShown = DateTime.UtcNow;
                RestartRecentTabs();
            }
            LayoutTabs(g);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var client = new Rectangle(0, 0, Math.Max(1, Width), Math.Max(1, Height));
            using (var bg = new LinearGradientBrush(client, StripTop, StripBottom, LinearGradientMode.Vertical))
                g.FillRectangle(bg, client);
            using (var border = new Pen(Border))
                g.DrawLine(border, 0, Height - 1, Width, Height - 1);

            // inactive tabs first, the active one last so it sits on top of the bottom line
            int activeIndex = -1;
            for (int i = 0; i < _items.Count; i++)
            {
                if (_items[i].Key == _mine) { activeIndex = i; continue; }
                PaintTab(g, i, false);
            }
            if (activeIndex >= 0) PaintTab(g, activeIndex, true);

            PaintPlus(g);

            if (_dragging && _dropIndex >= 0 && _tabRects.Count > 0)
            {
                int x = _dropIndex < _tabRects.Count ? _tabRects[_dropIndex].Left - 2
                                                    : _tabRects[_tabRects.Count - 1].Right + 2;
                using (var mark = new Pen(_accent, 3f)) g.DrawLine(mark, x, 4, x, Height - 3);
                using (var tri = new SolidBrush(_accent))
                    g.FillPolygon(tri, new[] { new Point(x - 5, 2), new Point(x + 5, 2), new Point(x, 8) });
            }
        }

        private void FinishDrag()
        {
            string key = _dragKey;
            int drop = _dropIndex;
            _dragKey = null; _dragging = false; _dropIndex = -1;
            Cursor = Cursors.Default;
            int from = _items.FindIndex(t => t.Key == key);
            if (from < 0 || drop < 0) { Invalidate(); return; }
            var list = new List<TabItem>(_items);
            TabItem moved = list[from];
            list.RemoveAt(from);
            if (drop > from) drop--;
            list.Insert(Math.Min(drop, list.Count), moved);
            _items = list;
            Relayout();
            Connect host = Connect.Current;
            if (_preview || host == null) return;
            List<string> keys = list.ConvertAll(t => t.Key);
            keys.RemoveAll(k => k.StartsWith("pending:"));
            Defer(delegate { host.ReorderTabs(keys); });
        }

        // ---------- tab menu (right click) ----------
        private void ShowTabMenu(int tab, Point at)
        {
            TabItem item = _items[tab];
            string key = item.Key;
            Connect host = Connect.Current;
            bool live = host != null && !_preview && !key.StartsWith("pending:");
            List<string> all = _items.ConvertAll(t => t.Key);
            all.RemoveAll(k => k.StartsWith("pending:"));
            List<string> others = all.FindAll(k => k != key);
            List<string> right = all.GetRange(Math.Min(tab + 1, all.Count), Math.Max(0, all.Count - tab - 1));

            var menu = new ContextMenuStrip();
            AddItem(menu, "New", live, delegate { AnimateThenNew(host); });
            menu.Items.Add(new ToolStripSeparator());
            AddItem(menu, "Save", live, delegate { host.SaveTab(key); });
            AddItem(menu, "Save All", live, delegate { host.SaveAll(); });
            menu.Items.Add(new ToolStripSeparator());
            AddItem(menu, "Close", live, delegate { host.CloseTab(key); });
            AddItem(menu, "Close Other Tabs", live && others.Count > 0, delegate { host.CloseTabs(others); });
            AddItem(menu, "Close Tabs to the Right", live && right.Count > 0, delegate { host.CloseTabs(right); });
            AddItem(menu, "Close All", live, delegate { host.CloseTabs(all); });
            menu.Items.Add(new ToolStripSeparator());
            string path = item.Path;
            AddItem(menu, "Copy Full Path", path != null, delegate { Clipboard.SetText(path); });
            AddItem(menu, "Open Containing Folder", path != null,
                    delegate { Process.Start("explorer.exe", "/select,\"" + path + "\""); });
            menu.Items.Add(new ToolStripSeparator());

            var colors = new ToolStripMenuItem("Tab Color");
            var none = new ToolStripMenuItem("None");
            none.Checked = item.Color == null;
            none.Click += delegate { Defer(delegate { ApplyColor(item, null); }); };
            colors.DropDownItems.Add(none);
            foreach (KeyValuePair<string, Color> c in TabColors)
            {
                string name = c.Key;
                var swatch = new Bitmap(14, 14);
                using (Graphics g = Graphics.FromImage(swatch))
                {
                    using (var b = new SolidBrush(c.Value)) g.FillRectangle(b, 0, 0, 14, 14);
                    g.DrawRectangle(Pens.Gray, 0, 0, 13, 13);
                }
                var mi = new ToolStripMenuItem(name, swatch);
                mi.Checked = item.Color == name;
                mi.Click += delegate { Defer(delegate { ApplyColor(item, name); }); };
                colors.DropDownItems.Add(mi);
            }
            menu.Items.Add(colors);

            menu.Closed += delegate { BeginInvoke(new MethodInvoker(menu.Dispose)); };
            menu.Show(this, at);
        }

        private void AddItem(ContextMenuStrip menu, string text, bool enabled, MethodInvoker action)
        {
            var mi = new ToolStripMenuItem(text);
            mi.Enabled = enabled;
            mi.Click += delegate { Defer(action); };
            menu.Items.Add(mi);
        }

        private void ApplyColor(TabItem item, string color)
        {
            Connect host = Connect.Current;
            if (_preview || host == null) { item.Color = color; Invalidate(); return; }
            host.SetTabColor(item.Key, item.Path, color);
        }

        private void PaintTab(Graphics g, int i, bool active)
        {
            TabItem item = _items[i];
            Rectangle r = _tabRects[i];
            if (r.Width < 24) return;     // still growing in: too narrow to draw yet
            if (!active) r.Height -= 1;   // stop above the strip's bottom line
            else r.Height += 1;
            Color accent = _preview ? AccentFor(item.Ext) : _accent;
            bool hover = i == _hoverTab;

            Color top, bottom, outline;
            Color tint;
            if (item.Color != null && TabColors.TryGetValue(item.Color, out tint))
            {
                // colored tab: stronger when active, lighter otherwise
                accent = Mix(tint, Color.Black, 0.35);
                if (active) { top = Mix(tint, Color.White, 0.25); bottom = Mix(tint, Color.White, 0.65); }
                else if (hover) { top = Mix(tint, Color.White, 0.55); bottom = Mix(tint, Color.White, 0.75); }
                else { top = Mix(tint, Color.White, 0.65); bottom = Mix(tint, Color.White, 0.80); }
                outline = Mix(tint, Border, 0.4);
            }
            else if (active) { top = Mix(accent, Color.White, 0.80); bottom = Color.White; outline = Mix(accent, Border, 0.45); }
            else if (hover) { top = Color.FromArgb(253, 253, 253); bottom = Color.FromArgb(242, 242, 242); outline = Border; }
            else { top = Color.FromArgb(236, 236, 236); bottom = Color.FromArgb(222, 222, 222); outline = Border; }

            using (GraphicsPath path = TabPath(r))
            {
                using (var fill = new LinearGradientBrush(new Rectangle(r.X, r.Y, r.Width, r.Height + 1), top, bottom,
                                                          LinearGradientMode.Vertical))
                    g.FillPath(fill, path);
                using (var pen = new Pen(outline)) g.DrawPath(pen, path);
            }
            if (active)
                using (var bar = new Pen(accent, 2f))
                    g.DrawLine(bar, r.Left + 7, r.Top + 1, r.Right - 7 - (_style == TabStyle.Trapezoid ? Slant : 0), r.Top + 1);

            // While the tab grows, keep its contents inside it.
            g.SetClip(new Rectangle(r.Left, r.Top, r.Width, r.Height + 1));
            try { PaintTabContents(g, i, item, r, active, hover); }
            finally { g.ResetClip(); }
        }

        private void PaintTabContents(Graphics g, int i, TabItem item, Rectangle r, bool active, bool hover)
        {
            int textLeft = r.Left + 10;
            if (_showIcons)
            {
                Bitmap icon = IconFor(item.Ext);
                if (icon != null)
                {
                    g.DrawImage(icon, new Rectangle(r.Left + 10, r.Top + (r.Height - 16) / 2, 16, 16));
                    textLeft += 22;
                }
            }

            Rectangle c = _closeRects[i];
            var textRect = new Rectangle(textLeft, r.Top, Math.Max(0, c.Left - 4 - textLeft), r.Height);
            TextRenderer.DrawText(g, item.Caption, active ? _activeFont : _tabFont, textRect,
                active ? Color.FromArgb(30, 30, 30) : Color.FromArgb(70, 70, 70),
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);

            // Unsaved changes show as a dot in the close button's place; hovering the tab
            // turns it back into the close button.
            bool closeHot = hover && _hoverClose;
            if (item.Modified && !hover)
            {
                using (var dot = new SolidBrush(Color.FromArgb(80, 80, 80)))
                    g.FillEllipse(dot, c.Left + 4, c.Top + 4, c.Width - 8, c.Height - 8);
                return;
            }
            if (closeHot)
                using (var hb = new SolidBrush(CloseHover)) g.FillEllipse(hb, c);
            if (active || hover)
                using (var xp = new Pen(closeHot ? Color.White : Color.FromArgb(95, 95, 95), 1.5f))
                {
                    g.DrawLine(xp, c.Left + 5, c.Top + 5, c.Right - 5, c.Bottom - 5);
                    g.DrawLine(xp, c.Right - 5, c.Top + 5, c.Left + 5, c.Bottom - 5);
                }
        }

        // Tab colors offered in the tab menu (name -> fill color).
        internal static readonly Dictionary<string, Color> TabColors = new Dictionary<string, Color>
        {
            { "Red", Color.FromArgb(239, 120, 120) },
            { "Orange", Color.FromArgb(255, 176, 92) },
            { "Yellow", Color.FromArgb(250, 214, 92) },
            { "Green", Color.FromArgb(124, 199, 128) },
            { "Blue", Color.FromArgb(110, 168, 238) },
            { "Purple", Color.FromArgb(186, 128, 214) },
        };

        private void PaintPlus(Graphics g)
        {
            Rectangle p = _plusRect;
            using (var path = new GraphicsPath())
            {
                const int rad = 5;
                path.AddArc(p.Left, p.Top, 2 * rad, 2 * rad, 180, 90);
                path.AddArc(p.Right - 2 * rad, p.Top, 2 * rad, 2 * rad, 270, 90);
                path.AddArc(p.Right - 2 * rad, p.Bottom - 2 * rad, 2 * rad, 2 * rad, 0, 90);
                path.AddArc(p.Left, p.Bottom - 2 * rad, 2 * rad, 2 * rad, 90, 90);
                path.CloseFigure();
                using (var b = new SolidBrush(_hoverPlus ? Color.White : Color.FromArgb(238, 238, 238)))
                    g.FillPath(b, path);
                using (var pen = new Pen(_hoverPlus ? Color.FromArgb(150, 150, 150) : Border))
                    g.DrawPath(pen, path);
            }
            using (var pp = new Pen(Color.FromArgb(85, 85, 85), 1.6f))
            {
                int cx = p.Left + p.Width / 2, cy = p.Top + p.Height / 2;
                g.DrawLine(pp, cx - 5, cy, cx + 5, cy);
                g.DrawLine(pp, cx, cy - 5, cx, cy + 5);
            }
        }

        // ---------- mouse ----------
        private int HitTab(Point p)
        {
            int n = Math.Min(_tabRects.Count, Math.Min(_closeRects.Count, _items.Count));
            for (int i = 0; i < n; i++)
                if (_tabRects[i].Contains(p)) return i;
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            try { MouseMoved(e); } catch { }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            try
            {
                int tab = HitTab(e.Location);
                _dragKey = (e.Button == MouseButtons.Left && tab >= 0 && !_closeRects[tab].Contains(e.Location))
                           ? _items[tab].Key : null;
                _dragStart = e.Location;
                _dragging = false;
                _dropIndex = -1;
            }
            catch { _dragKey = null; }
        }

        // Where a dragged tab would land: the slot before the first tab whose middle is
        // right of the mouse.
        private int DropIndexAt(int x)
        {
            for (int i = 0; i < _tabRects.Count; i++)
                if (x < _tabRects[i].Left + _tabRects[i].Width / 2) return i;
            return _tabRects.Count;
        }

        private void MouseMoved(MouseEventArgs e)
        {
            if (_dragKey != null && (e.Button & MouseButtons.Left) != 0)
            {
                if (!_dragging && Math.Abs(e.X - _dragStart.X) > 6) _dragging = true;
                if (_dragging)
                {
                    int drop = DropIndexAt(e.X);
                    if (drop != _dropIndex) { _dropIndex = drop; Invalidate(); }
                    Cursor = Cursors.SizeWE;
                    return;
                }
            }
            int tab = HitTab(e.Location);
            bool close = tab >= 0 && _closeRects[tab].Contains(e.Location);
            bool plus = _plusRect.Contains(e.Location);
            if (tab != _hoverTab || close != _hoverClose || plus != _hoverPlus)
            {
                _hoverTab = tab; _hoverClose = close; _hoverPlus = plus;
                Invalidate();
            }
            if (tab != _tipTab)
            {
                _tipTab = tab;
                if (tab >= 0) _tip.SetToolTip(this, _items[tab].Path ?? _items[tab].Caption);
                else _tip.SetToolTip(this, null);
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (!_dragging) Cursor = Cursors.Default;
            _hoverTab = -1; _hoverClose = false; _hoverPlus = false;
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            try { MouseReleased(e); } catch { }
        }

        private void MouseReleased(MouseEventArgs e)
        {
            int tab = HitTab(e.Location);
            if (_dragging)
            {
                FinishDrag();
                return;
            }
            _dragKey = null;
            if (e.Button == MouseButtons.Right && tab >= 0)
            {
                ShowTabMenu(tab, e.Location);
                return;
            }
            if (_preview)
            {
                if (e.Button != MouseButtons.Left) return;
                if (_plusRect.Contains(e.Location) && _items.Count < 6)
                {
                    // demo of the new-tab animation
                    string key = "new" + _items.Count;
                    string[] exts = { ".docx", ".xlsx", ".pptx" };
                    _items.Add(Sample(key, "New file " + (_items.Count - 2), exts[_items.Count % 3]));
                    _mine = key;
                    NoteNewTabs();
                    Relayout();
                }
                else if (tab >= 0) { _mine = _items[tab].Key; Invalidate(); }
                return;
            }

            Connect host = Connect.Current;
            if (host == null) return;
            if (e.Button == MouseButtons.Left)
            {
                if (_plusRect.Contains(e.Location)) { AnimateThenNew(host); return; }
                if (tab < 0) return;
                string key = _items[tab].Key;
                if (_closeRects[tab].Contains(e.Location)) Defer(delegate { host.CloseTab(key); });
                else if (key != _mine) Defer(delegate { host.ActivateTab(key); });
            }
            else if (e.Button == MouseButtons.Middle && tab >= 0)
            {
                string key = _items[tab].Key;
                Defer(delegate { host.CloseTab(key); });
            }
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            try
            {
                Connect host = Connect.Current;
                if (host != null && !_preview && _newOnDoubleClick && e.Button == MouseButtons.Left &&
                    HitTab(e.Location) < 0 && !_plusRect.Contains(e.Location))
                    AnimateThenNew(host);
            }
            catch { }
        }

        // Run after the current click finishes: closing a document can destroy this very pane.
        private void Defer(MethodInvoker action)
        {
            try { BeginInvoke(new MethodInvoker(delegate { try { action(); } catch { } })); } catch { }
        }
    }
}
