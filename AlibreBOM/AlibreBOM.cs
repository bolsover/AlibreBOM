using System;
using System.IO;
using System.Security.Cryptography;
using System.Windows.Forms;
using AlibreAddOn;
using AlibreX;
using SpreadsheetLight;

namespace AlibreBOM
{
    public class AlibreBOM : IAlibreAddOn
    {
        private const int MenuIdRoot = 401;
        private const int MenuIdSample = 402;

        private readonly int[] _menuIdsBase = new int[]
        {
            MenuIdSample
        };

        private IADRoot _alibreRoot;
        private IntPtr _parentWinHandle;
        private readonly bool _useSvgIcons;

        public AlibreBOM(IADRoot alibreRoot, IntPtr parentWinHandle)
        {
            _alibreRoot = alibreRoot;
            _parentWinHandle = parentWinHandle;
            string version = _alibreRoot.Version.Replace("PRODUCTVERSION ", "");
            string[] versionarr = version.Split(',');
            int majorVersion = int.Parse(versionarr[0]);
            _useSvgIcons = majorVersion > 25;
        }

        #region Menus

        /// <summary>
        /// Returns the menu ID of the add-on's root menu item
        /// </summary>
        public int RootMenuItem => MenuIdRoot;


        /// <summary>
        /// Description("Returns Whether the given Menu ID has any sub menus")
        /// </summary>
        /// <param name="menuId"></param>
        /// <returns></returns>
        public bool HasSubMenus(int menuId)
        {
            //   return false;
            return menuId == MenuIdRoot;
        }

        /// <summary>
        /// Returns the ID's of sub menu items under a popup menu item; the menu ID of a 'leaf' menu becomes its command ID
        /// </summary>
        /// <param name="menuId"></param>
        /// <returns></returns>
        public Array SubMenuItems(int menuId)
        {
            return _menuIdsBase;
        }

        /// <summary>
        /// Returns the display name of a menu item; a menu item with text of a single dash (“-“) is a separator
        /// </summary>
        /// <param name="menuId"></param>
        /// <returns></returns>
        public string MenuItemText(int menuId)
        {
            return "Write .xlsx BOM";
        }

        /// <summary>
        /// Returns True if input menu item has sub menus // seems odd given name of method
        /// </summary>
        /// <param name="menuId"></param>
        /// <returns></returns>
        public bool PopupMenu(int menuId)
        {
            return true;
        }

        /// <summary>
        /// Returns property bits providing information about the state of a menu item
        /// ADDON_MENU_ENABLED = 1,
        /// ADDON_MENU_GRAYED = 2,
        /// ADDON_MENU_CHECKED = 3,
        /// ADDON_MENU_UNCHECKED = 4,
        /// </summary>
        /// <param name="menuId"></param>
        /// <param name="sessionIdentifier"></param>
        /// <returns></returns>
        public ADDONMenuStates MenuItemState(int menuId, string sessionIdentifier)
        {
            var session = _alibreRoot.Sessions.Item(sessionIdentifier);

            switch (session)
            {
                case IADAssemblySession: return ADDONMenuStates.ADDON_MENU_ENABLED;
                //case IADPartSession: return ADDONMenuStates.ADDON_MENU_ENABLED;
            }


            return ADDONMenuStates.ADDON_MENU_GRAYED;
        }

        /// <summary>
        /// Returns a tool tip string if input menu ID is that of a 'leaf' menu item
        /// </summary>
        /// <param name="menuId"></param>
        /// <returns></returns>
        public string MenuItemToolTip(int menuId)
        {
            return "Write XLSX BOM";
        }

        /// <summary>
        /// Returns the icon name (with extension) for a menu item; the icon will be searched under the folder where the add-on's .adc file is present
        /// </summary>
        /// <param name="menuId"></param>
        /// <returns></returns>
        public string MenuIcon(int menuId)
        {
            return _useSvgIcons ? "3DPrint.svg" : "3DPrint.ico";
        }

        /// <summary>
        /// Returns True if AddOn has updated Persistent Data
        /// </summary>
        /// <param name="sessionIdentifier"></param>
        /// <returns></returns>
        public bool HasPersistentDataToSave(string sessionIdentifier)
        {
            return false;
        }

        /// <summary>
        /// Invokes the add-on command identified by menu ID; returning the add-on command interface is optional
        /// </summary>
        /// <param name="menuId"></param>
        /// <param name="sessionIdentifier"></param>
        /// <returns></returns>
        public IAlibreAddOnCommand InvokeCommand(int menuId, string sessionIdentifier)
        {
            var session = _alibreRoot.Sessions.Item(sessionIdentifier);
            return DoBOM(session);
        }

        #endregion

        #region Export

        private IAlibreAddOnCommand DoBOM(IADSession currentSession)
        {
            IADOccurrence adOccurrence = ((IADAssemblySession) currentSession).RootOccurrence;
            Console.WriteLine(adOccurrence.NestLevel + " " + adOccurrence.Name + " " + adOccurrence.Configuration.Name);
            SLDocument sl = new SLDocument();
            sl.SetCellValue("A1", "Level");
            sl.SetCellValue("B1", "Name");
            sl.SetCellValue("C1", "Configuration");
            sl.SetCellValue("D1", "Path");
            sl.SetCellValue("E1", "Description");
            sl.SetCellValue("F1", "Number");
            sl.SetCellValue("G1", "Material");
            sl.SetCellValue("A2", adOccurrence.NestLevel);
            sl.SetCellValue("B2", adOccurrence.Name);
            sl.SetCellValue("C2", adOccurrence.Configuration.Name);
            sl.SetCellValue("D2", adOccurrence.DesignSession.FilePath);
            sl.SetCellValue("E2", adOccurrence.DesignSession.DesignProperties.Description);
            sl.SetCellValue("F2", adOccurrence.DesignSession.DesignProperties.Number);
            sl.SetCellValue("G2", adOccurrence.DesignSession.DesignProperties.Material);
            row = 2;
            DIEnum diEnumParent = adOccurrence.Occurrences.Enum;

            WalkChildren(sl, diEnumParent);

            SaveExcelFile(sl, currentSession);


            return null;
        }

        private void SaveExcelFile(SLDocument sl, IADSession currentSession)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel|*.xlsx";
            saveFileDialog1.Title = "Save Excel BOM";
            saveFileDialog1.FileName = currentSession.Name + "BOM.xlsx";
            DialogResult dr = saveFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                sl.SaveAs(saveFileDialog1.FileName);


                System.Diagnostics.Process.Start(Path.GetFullPath(saveFileDialog1.FileName));
            }
        }

        private int row;

        private void WalkChildren(SLDocument sl, DIEnum diEnumParent)
        {
            while (diEnumParent.HasMoreElements())
            {
                row++;
                IADOccurrence occurrence = (IADOccurrence) diEnumParent.NextElement();
                sl.SetCellValue("A" + row, occurrence.NestLevel);
                sl.SetCellValue("B" + row, occurrence.Name);
                sl.SetCellValue("C" + row, occurrence.Configuration.Name);
                sl.SetCellValue("D" + row, occurrence.DesignSession.FilePath);
                sl.SetCellValue("E" + row, occurrence.DesignSession.DesignProperties.Description);
                sl.SetCellValue("F" + row, occurrence.DesignSession.DesignProperties.Number);
                sl.SetCellValue("G" + row, occurrence.DesignSession.DesignProperties.Material);
                if (occurrence.Occurrences.Enum.HasMoreElements())
                {
                    WalkChildren(sl, occurrence.Occurrences.Enum);
                }
            }
        }

        #endregion


        /// <summary>
        /// Loads Data from AddOn
        /// </summary>
        /// <param name="pCustomData"></param>
        /// <param name="sessionIdentifier"></param>
        public void LoadData(IStream pCustomData, string sessionIdentifier)
        {
        }

        /// <summary>
        /// Saves Data to AddOn
        /// </summary>
        /// <param name="pCustomData"></param>
        /// <param name="sessionIdentifier"></param>
        public void SaveData(IStream pCustomData, string sessionIdentifier)
        {
        }

        /// <summary>
        /// Sets the IsLicensed bit for the tightly coupled Add-on
        /// </summary>
        /// <param name="isLicensed"></param>
        public void setIsAddOnLicensed(bool isLicensed)
        {
        }

        /// <summary>
        /// Returns True if the AddOn needs to use a Dedicated Ribbon Tab
        /// </summary>
        /// <returns></returns>
        public bool UseDedicatedRibbonTab()
        {
            return true;
        }
    }
}