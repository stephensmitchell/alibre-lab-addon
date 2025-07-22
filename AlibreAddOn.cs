using AlibreAddOn;
using AlibreX;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;
using IStream = System.Runtime.InteropServices.ComTypes.IStream;

namespace AlibreAddOnAssembly
{
    public static class StringConstants
    {
        public const string V_STR_1 = "GenericAddon";
        public const string V_STR_2 = "Failed to initialize: ";
        public const string V_STR_3 = "Generic Add-on Loaded";
        public const string V_STR_4 = "Load Status";
        public const string V_STR_5 = "Main Tools";
        public const string V_STR_6 = "Utilities";
        public const string V_STR_7 = "Import File";
        public const string V_STR_8 = "Processes a file and imports it.";
        public const string V_STR_9 = "About";
        public const string V_STR_10 = "Shows information about this add-on.";
        public const string V_STR_11 = @"Icons\icon.ico";
        public const string V_STR_12 = "Select file to process";
        public const string V_STR_13 = "Data files (*.dat)|*.dat";
        public const string V_STR_14 = ".out";
        public const string V_STR_15 = "bin";
        public const string V_STR_16 = "processor.exe";
        public const string V_STR_17 = "File Processor Add-on\n\nThis utility uses an external executable to process files.";
        public const string V_STR_18 = "Executable Missing";
        public const string V_STR_19 = "Process Execution Error";
        public const string V_STR_20 = "Executable not found at the expected path:\n{0}";
        public const string V_STR_21 = "\"{0}\" \"{1}\"";
        public const string V_STR_22 = "File import failed (API error):\n{0}\n\nFile: {1}";
        public const string V_STR_23 = "processor.exe failed to handle '{0}'.\n\nExit Code: {1}\n\nError Output:\n{2}";
        public const string V_STR_24 = "An unexpected error occurred during the file processing operation:\n{0}";
    }

    public static class AlibreAddOn
    {
        private static IADRoot _root { get; set; }
        private static UIHandler _ui;

        public static void AddOnLoad(IntPtr h, IAutomationHook hook, IntPtr unused)
        {
            _root = (IADRoot)hook.Root;
            _ui = new UIHandler(_root);
            MessageBox.Show(
                StringConstants.V_STR_3,
                StringConstants.V_STR_4,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        public static void AddOnUnload(IntPtr h, bool f, ref bool c, int r1, int r2)
        {
            _ui = null;
            _root = null;
        }

        public static IAlibreAddOn GetAddOnInterface()
        {
            return _ui;
        }

        public static void AddOnInvoke(IntPtr h, IntPtr hook, string sName, bool licensed, int r1, int r2) { }
    }

    public class UIHandler : IAlibreAddOn
    {
        private readonly IADRoot _appRoot;
        private readonly UIManager _uiManager;

        public UIHandler(IADRoot root)
        {
            _appRoot = root;
            try
            {
                _uiManager = new UIManager(this);
            }
            catch (Exception e)
            {
                Debug.WriteLine(StringConstants.V_STR_2 + e.Message);
            }
        }

        public bool UseDedicatedRibbonTab() => true;
        public string RibbonTabName() => _uiManager.TabLabel;
        public bool HasRibbonPanels(string sId) => _uiManager.Panels.Any();
        public Array RibbonPanels(string sId) => _uiManager.Panels.Select(p => p.ID).ToArray();
        public string RibbonPanelName(int pId) => _uiManager.GetPanel(pId)?.Name;
        public Array RibbonPanelItems(int pId, string sId) => _uiManager.GetPanel(pId)?.Buttons.Select(b => b.ID).ToArray();
        public string RibbonPanelItemText(int pId, int iId, string sId) => _uiManager.GetButton(pId, iId)?.Label;
        public string RibbonPanelItemToolTip(int pId, int iId, string sId) => _uiManager.GetButton(pId, iId)?.Hint;
        public string RibbonPanelItemIcon(int pId, int iId, string sId) => _uiManager.GetButton(pId, iId)?.Icon;
        public ADDONMenuStates RibbonPanelItemState(int pId, int iId, string sId) => ADDONMenuStates.ADDON_MENU_ENABLED;

        public IAlibreAddOnCommand InvokeRibbonPanelItem(int pId, int iId, string sId)
        {
            var session = _appRoot.Sessions.Item(sId);
            return _uiManager.GetButton(pId, iId)?.Action?.Invoke(session);
        }

        public IAlibreAddOnCommand ShowInfoDialog(IADSession s)
        {
            MessageBox.Show(StringConstants.V_STR_17, StringConstants.V_STR_9, MessageBoxButton.OK, MessageBoxImage.Information);
            return null;
        }

        public IAlibreAddOnCommand ExecuteFileConversion(IADSession s)
        {
            var fileDialog = new OpenFileDialog
            {
                Title = StringConstants.V_STR_12,
                Filter = StringConstants.V_STR_13,
                CheckFileExists = true,
                Multiselect = false
            };

            if (fileDialog.ShowDialog() != DialogResult.OK)
                return null;

            string path1 = fileDialog.FileName;
            string path2 = Path.ChangeExtension(path1, StringConstants.V_STR_14);
            string dir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)!;
            string path3 = Path.Combine(dir1, StringConstants.V_STR_15, StringConstants.V_STR_16);

            if (!File.Exists(path3))
            {
                MessageBox.Show(string.Format(StringConstants.V_STR_20, path3), StringConstants.V_STR_18, MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            try
            {
                var pInfo = new ProcessStartInfo(path3, string.Format(StringConstants.V_STR_21, path1, path2))
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var p = Process.Start(pInfo))
                {
                    string errOut = p.StandardError.ReadToEnd();
                    p.WaitForExit();

                    if (p.ExitCode == 0 && File.Exists(path2))
                    {
                        // _appRoot.Import(path2); 
                    }
                    else
                    {
                        string errMessage = string.Format(StringConstants.V_STR_23, Path.GetFileName(path1), p.ExitCode, errOut);
                        MessageBox.Show(errMessage, StringConstants.V_STR_19, MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format(StringConstants.V_STR_24, e.Message), StringConstants.V_STR_19, MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return null;
        }

        public int RootMenuItem => 0;
        public bool HasSubMenus(int mID) => false;
        public Array SubMenuItems(int mID) => null;
        public string MenuItemText(int mID) => null;
        public string MenuItemToolTip(int mID) => null;
        public string MenuIcon(int mID) => null;
        public bool PopupMenu(int mID) => false;
        public ADDONMenuStates MenuItemState(int mID, string sID) => ADDONMenuStates.ADDON_MENU_UNCHECKED;
        public IAlibreAddOnCommand InvokeCommand(int mID, string sID) => null;
        public bool HasPersistentDataToSave(string sID) => false;
        public void SaveData(IStream pcd, string sID) { }
        public void LoadData(IStream pcd, string sID) { }
        public void setIsAddOnLicensed(bool licensed) { }
        public void LoadData(global::AlibreAddOn.IStream pcd, string sID) { throw new NotImplementedException(); }
        public void SaveData(global::AlibreAddOn.IStream pcd, string sID) { throw new NotImplementedException(); }
    }

    public class UIButton
    {
        public int ID { get; }
        public string Label { get; }
        public string Hint { get; }
        public string Icon { get; }
        public Func<IADSession, IAlibreAddOnCommand> Action { get; }

        public UIButton(int id, string text, string tip, string icon, Func<IADSession, IAlibreAddOnCommand> cmd)
        {
            ID = id;
            Label = text;
            Hint = tip;
            Icon = icon;
            Action = cmd;
        }
    }

    public class UIPanel
    {
        public int ID { get; }
        public string Name { get; }
        public List<UIButton> Buttons { get; }

        public UIPanel(int id, string name)
        {
            ID = id;
            Name = name;
            Buttons = new List<UIButton>();
        }

        public void Add(UIButton b) => Buttons.Add(b);
    }
    public class UIManager
    {
        public string TabLabel { get; }
        public List<UIPanel> Panels { get; }
        private readonly Dictionary<int, UIPanel> _panelMap;
        private readonly Dictionary<int, Dictionary<int, UIButton>> _buttonMap;

        public UIManager(UIHandler ui)
        {
            TabLabel = StringConstants.V_STR_5;
            Panels = new List<UIPanel>();
            _panelMap = new Dictionary<int, UIPanel>();
            _buttonMap = new Dictionary<int, Dictionary<int, UIButton>>();

            BuildLayout(ui);
            Register();
        }

        private void BuildLayout(UIHandler ui)
        {
            var p1 = new UIPanel(101, StringConstants.V_STR_6);
            var b1 = new UIButton(201, StringConstants.V_STR_7, StringConstants.V_STR_8, StringConstants.V_STR_11, ui.ExecuteFileConversion);
            var b2 = new UIButton(202, StringConstants.V_STR_9, StringConstants.V_STR_10, StringConstants.V_STR_11, ui.ShowInfoDialog);
            p1.Add(b1);
            p1.Add(b2);
            Panels.Add(p1);
        }

        private void Register()
        {
            foreach (var p in Panels)
            {
                _panelMap[p.ID] = p;
                _buttonMap[p.ID] = new Dictionary<int, UIButton>();
                foreach (var b in p.Buttons)
                {
                    _buttonMap[p.ID][b.ID] = b;
                }
            }
        }
        public UIPanel GetPanel(int pId) =>
            _panelMap.TryGetValue(pId, out var p) ? p : null;

        public UIButton GetButton(int pId, int bId) =>
            _buttonMap.TryGetValue(pId, out var pButtons) && pButtons.TryGetValue(bId, out var b) ? b : null;
    }
}