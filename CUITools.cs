using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Customization;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using pgDBTools;
using Westmark;
using wMenu = Westmark.CADNames.CADMenu;

namespace Westmark.CAD
{
    public static class CUI
    {
        // Used to support the saving and updating of CUI settings.
        [DllImport("acad.exe", EntryPoint = "acedCmd")]
        private static extern int acedCmd(System.IntPtr vlist);

        static CustomizationSection cs = null;
        public static bool AddPanelFromDB(string WorkspaceName, string PanelName, string[] TabNames)
        {
            return AddPanelFromDB(WorkspaceName, PanelName, TabNames, false);
        }
        public static bool AddPanelFromDB(string WorkspaceName, string PanelName, string[] TabNames, bool UpdateExisting)
        {
            if (cs == null) cs = GetCS();
            bool res = false;
            if (WorkspaceExists(cs, WorkspaceName))
            {
                bool pnlExists = PanelExists(PanelName);
                if (pnlExists && !UpdateExisting)
                {
                    // Panel already exists
                    // Don't add or update
                    // UpdatePanelFromDB is for updating
                }
                else
                {
                    if (pnlExists)
                    {
                        // need to delete it
                        cs.MenuGroup.RibbonRoot.RibbonPanelSources.Remove(GetPanelByName(PanelName));
                        cs.Save();                       
                    }
                    // get data for panel with this name
                    pgDB odb = new pgDB();
                    string sql = "SELECT * " +
                        "FROM " + wMenu.DB_TABLE_ACAD_PANEL_DEF + " " +
                        "WHERE " + wMenu.DB_FIELD_PNL_TITLE + " = '" + PanelName + "'";
                    System.Data.DataTable dt = odb.pgDataTable("eng", sql);
                    if (dt.Rows.Count > 0)
                    {
                        // Create the panel definition
                        string newPanelElementID = dt.Rows[0][wMenu.DB_FIELD_ELEM_ID].ToString();
                        // new panel definition
                        RibbonPanelSource rpSource = new RibbonPanelSource(cs.MenuGroup.RibbonRoot);
                        rpSource.ElementID = newPanelElementID;
                        rpSource.Name = dt.Rows[0][wMenu.DB_FIELD_PNL_NAME].ToString();
                        rpSource.Text = dt.Rows[0][wMenu.DB_FIELD_PNL_TITLE].ToString();
                        rpSource.KeyTip = dt.Rows[0][wMenu.DB_FIELD_PNL_KEY_TIP].ToString();
                        
                        // get the buttons (items)
                        sql = "SELECT * " +
                            "FROM " + wMenu.DB_TABLE_ACAD_PANEL_ITEMS + " " +
                            "WHERE " + wMenu.DB_FIELD_PANEL_ID + " = " + dt.Rows[0][wMenu.DB_FIELD_ID] + " " +
                            "ORDER BY " + wMenu.DB_FIELD_ROW_INDEX + ", " + wMenu.DB_FIELD_SUBPANEL_INDEX + ", " +
                            wMenu.DB_FIELD_SUBROW_INDEX + ", " + wMenu.DB_FIELD_ITEM_ORDER;
                        
                        dt = odb.pgDataTable("eng", sql);

                        if (dt.Rows.Count > 0)
                        {
                            int row = -1; // (int)dt.Rows[0][wMenu.DB_FIELD_ROW_INDEX];
                            int subRow = 0;
                            int subPanel = 0;
                            // new panel row
                            RibbonRow ribRow = new RibbonRow();
                            RibbonRowPanel ribRowPanel = null;
                            for (int j = 0; j < dt.Rows.Count; j++)
                            {
                                string eid = dt.Rows[j][wMenu.DB_FIELD_ELEM_ID].ToString().ToUpper();
                                if (eid.Equals(wMenu.SLIDE_OUT.ToUpper())) // check for slide-out marker (ONLY ONE PER PANEL)
                                {
                                    RibbonPanelBreak ribPanelBreak = new RibbonPanelBreak(rpSource);
                                    rpSource.Items.Add(ribPanelBreak);
                                }
                                else
                                {
                                    // add a new row as needed
                                    if (row < (int)dt.Rows[j][wMenu.DB_FIELD_ROW_INDEX])
                                    {
                                        // add the row to the panel
                                        ribRow = new RibbonRow();
                                        rpSource.Items.Add((RibbonItem)ribRow);
                                        row = (int)dt.Rows[j][wMenu.DB_FIELD_ROW_INDEX];
                                        subRow = 0;
                                        subPanel = 0;
                                    }
                                    if (subPanel < (int)dt.Rows[j][wMenu.DB_FIELD_SUBPANEL_INDEX])
                                    {
                                        // add a sub panel
                                        subPanel = (int)dt.Rows[j][wMenu.DB_FIELD_SUBPANEL_INDEX];
                                        subRow = 0;
                                        ribRowPanel = new RibbonRowPanel((RibbonItem)ribRow);
                                        ribRow.Items.Add((RibbonItem)ribRowPanel);
                                        ribRow = new RibbonRow((RibbonItem)ribRowPanel);
                                        ribRowPanel.Items.Add((RibbonItem)ribRow);
                                    }
                                    else if (subRow < (int)dt.Rows[j][wMenu.DB_FIELD_SUBROW_INDEX])
                                    {
                                        // add a sub row
                                        subRow = (int)dt.Rows[j][wMenu.DB_FIELD_SUBROW_INDEX];
                                        if (ribRowPanel == null)
                                        {
                                            ribRow = new RibbonRow();
                                            rpSource.Items.Add((RibbonItem)ribRow);
                                        }
                                        else
                                        {
                                            ribRow = new RibbonRow((RibbonItem)ribRowPanel);
                                            ribRowPanel.Items.Add((RibbonItem)ribRow);
                                        }
                                    }
                                    // new button
                                    RibbonCommandButton ribCmdBtn = CreateButtonFromDB(ribRow, dt.Rows[j]); // new RibbonCommandButton(ribRow);
                                    // add the button to the row
                                    ribRow.Items.Add((RibbonItem)ribCmdBtn);

                                }

                            }
                        }
                        
                        // add the panel to the ribbon
                        cs.MenuGroup.RibbonRoot.RibbonPanelSources.Add(rpSource);

                        // Add it to each tab 
                        int tabCount = TabNames.GetUpperBound(0) + 1;
                        if (tabCount > 0)
                        {
                            for (int i = 0; i < tabCount; i++)
                            {
                                string TabName = TabNames[i];
                                if (WorkspaceHasTab(cs, WorkspaceName, TabName))
                                {
                                    RibbonTabSource tab = GetTab(cs, TabName);
                                    if (TabHasPanel(tab, PanelName))
                                    {
                                        // Has already been added
                                    }
                                    else
                                    {
                                        // add the panel to this tab
                                        RibbonPanelSourceReference newPanel = new RibbonPanelSourceReference(tab);
                                        newPanel.PanelId = rpSource.ElementID;
                                        tab.Items.Add(newPanel);
                                    }
                                }
                            }
                        }
                        saveCui(true, false);
                    }
                }
            }
            return res;
        }

        private static RibbonCommandButton CreateButtonFromDB(RibbonRow ribRow, DataRow dataRow)
        {
            RibbonCommandButton cmdButton = new RibbonCommandButton(ribRow);
            string buttonDefID = dataRow[wMenu.DB_FIELD_ELEM_ID].ToString();
            cmdButton.Text = dataRow[wMenu.DB_FIELD_BUTTON_NAME].ToString();
            cmdButton.TooltipTitle = dataRow[wMenu.DB_FIELD_TOOL_TIP].ToString();
            string style = dataRow[wMenu.DB_FIELD_BUTTON_STYLE].ToString();
            RibbonButtonStyle bstyle = RibbonButtonStyle.SmallWithoutText;
            switch (style)
            {
                case wMenu.BUTTON_STYLE_LARGE_TXT:
                    bstyle = RibbonButtonStyle.LargeWithText;
                    break;
                case wMenu.BUTTON_STYLE_LARGE_TXT_HOR:
                    bstyle = RibbonButtonStyle.LargeWithHorizontalText;
                    break;
                case wMenu.BUTTON_STYLE_SMALL:
                    bstyle = RibbonButtonStyle.SmallWithoutText;
                    break;
                case wMenu.BUTTON_STYLE_SMALL_TXT:
                    bstyle = RibbonButtonStyle.SmallWithText;
                    break;
            }
            cmdButton.ButtonStyle = bstyle;

            if (cs == null) cs = GetCS();

            string macID = GetMacroIDFromName(dataRow[wMenu.DB_FIELD_MACRO_NAME].ToString());
            bool macroExists = MacroExists(macID);

            if (!ButtonExists(buttonDefID))
            {
                if (!macroExists)
                {
                    MacroGroup macGroup = TryGetUserMacroGroup(cs);
                    string name = dataRow[wMenu.DB_FIELD_MACRO_NAME].ToString();
                    string command = dataRow[wMenu.DB_FIELD_MACRO].ToString();
                    string help = string.Empty;

                    string smallImage = dataRow[wMenu.DB_FIELD_SMALL_IMG].ToString();
                    string largeImage = dataRow[wMenu.DB_FIELD_LARGE_IMG].ToString();
                    string imgPath = Westmark.CADNames.CADFilePaths.ICON_DIRECTORY_LOCAL;
                    string imgPathSvr = Westmark.CADNames.CADFilePaths.ICON_DIRECTORY_SERVER;

                    #region Image file checks

                    //if (!File.Exists(imgPath + smallImage))
                    //{
                        if (File.Exists(imgPathSvr + smallImage))
                        {
                            try
                            {
                                File.Copy(imgPathSvr + smallImage, imgPath + smallImage, true);
                            }
                            catch (System.Exception ex)
                            {
                            }
                        }
                        else
                        {
                            smallImage = Westmark.CADNames.CADFilePaths.BLANK_BUTTON_16_FILE;
                            if (File.Exists(imgPathSvr + smallImage))
                            {
                                try
                                {
                                    File.Copy(imgPathSvr + smallImage, imgPath + smallImage, true);
                                }
                                catch (System.Exception ex)
                                {
                                }
                            }
                        }
                    //}
                    //else
                    //{
                    //    // image exists but go ahead and overwrite it

                    //}
                    //if (!File.Exists(imgPath + largeImage))
                    //{
                        if (File.Exists(imgPathSvr + largeImage))
                        {
                            try
                            {
                                File.Copy(imgPathSvr + largeImage, imgPath + largeImage, true);
                            }
                            catch (System.Exception ex)
                            {
                            }
                        }
                        else
                        {
                            largeImage = Westmark.CADNames.CADFilePaths.BLANK_BUTTON_32_FILE;
                            if (File.Exists(imgPathSvr + largeImage))
                            {
                                try
                                {
                                    File.Copy(imgPathSvr + largeImage, imgPath + largeImage, true);
                                }
                                catch (System.Exception ex)
                                {
                                }
                            }
                        }
                    //}

                    #endregion

                    MenuMacro menMac = macGroup.CreateMenuMacro(
                        name,
                        command,
                        string.Empty,
                        help,
                        MacroType.Any,
                        smallImage,
                        largeImage,
                        string.Empty);
                    macID = menMac.ElementID;
                }
                cmdButton.MacroID = macID;
            }
            else
            {
                // button already exists
                cmdButton.ElementID = buttonDefID;
            }
            return cmdButton;
        }

        private static MacroGroup TryGetUserMacroGroup(CustomizationSection cs)
        {
            MacroGroup resMG = cs.MenuGroup.MacroGroups[0];
            MacroGroupCollection mgc = cs.MenuGroup.MacroGroups;
            foreach (MacroGroup mg in mgc)
            {
                if (mg.Name == "User")
                {
                    resMG = mg;
                    break;
                }
            }
            return resMG;
        }

        private static string GetMacroIDFromName(string MacroName)
        {
            string res = string.Empty;
            if (cs == null) cs = GetCS();
            MenuMacroCollection mmc = cs.MenuGroup.MacroGroups[0].MenuMacros;
            foreach (MenuMacro mm in mmc)
            {
                if (mm.macro.Name == MacroName)
                {
                    res = mm.ElementID;
                    break;
                }
            }
            return res;
        }

        private static void DeleteMacro(string macID)
        {
            if (cs == null) cs = GetCS();
            MenuMacro mm = null;
            mm = cs.getMenuMacro(macID);
            if (mm != null)
            {
                cs.MenuGroup.MacroGroups[0].MenuMacros.Remove(mm);
            }
        }

        private static bool MacroExists(string macroID)
        {
            bool res = false;
            if (cs == null) cs = GetCS();
            MenuMacro mm = null;
            mm = cs.getMenuMacro(macroID);
            res = (mm != null);
            return res;
        }

        private static bool ButtonExists(string buttonDefID)
        {
            bool res = false;
            if (cs == null) cs = GetCS();
            RibbonRoot rr = cs.MenuGroup.RibbonRoot;
            RibbonItem ri = rr.FindButton(buttonDefID);
            res = (ri != null);
            return res;
        }

        private static bool TabHasPanel(RibbonTabSource tab, string PanelName)
        {
            bool res = false;
            if (cs == null) cs = GetCS();
            RibbonRoot root = cs.MenuGroup.RibbonRoot;
            RibbonPanelSourceCollection pnlColl = root.RibbonPanelSources;
            RibbonPanelSource rps = null;
            string pnlID = string.Empty;
            foreach (RibbonPanelSource pnl in pnlColl)
            {
                if (pnl.Text == PanelName)
                {
                    rps = pnl;
                    pnlID = rps.ElementID;
                    break;
                }
            }
            if (rps != null)
            {
                RibbonPanelSourceReference rpsref = tab.Find(pnlID);
                if (rpsref != null)
                {
                    res = true;
                }
            }
            return res;
        }

        private static RibbonTabSource GetTab(CustomizationSection cs, string TabName)
        {
            RibbonTabSource res = null;
            RibbonTabSourceCollection rtsc = cs.MenuGroup.RibbonRoot.RibbonTabSources;
            foreach (RibbonTabSource rtsrc in rtsc)
            {
                if (rtsrc.Text == TabName)
                {
                    res = rtsrc;
                    break;
                }
            }
            return res;
        }



        public static bool WorkspaceExists(string WorkspaceName)
        {
            bool res = false;
            if (cs == null) cs = GetCS();
            res = (-1 != cs.Workspaces.IndexOfWorkspaceName(WorkspaceName));
            return res;
        }
        public static bool WorkspaceExists(CustomizationSection cs, string WorkspaceName)
        {
            bool res = false;
            res = (-1 != cs.Workspaces.IndexOfWorkspaceName(WorkspaceName));
            return res;
        }

        public static bool WorkspaceHasTab(string WorkspaceName,string TabName)
        {
            bool res = false;
            if (WorkspaceExists(WorkspaceName))
            {
                CustomizationSection cs = GetCS();
                Workspace ws = cs.Workspaces[cs.Workspaces.IndexOfWorkspaceName(WorkspaceName)];
                WSRibbonRoot wsrr = ws.WorkspaceRibbonRoot;
                RibbonTabSource rts = null;
                string tabEID = string.Empty;
                RibbonTabSourceCollection rtsc = cs.MenuGroup.RibbonRoot.RibbonTabSources;
                foreach (RibbonTabSource rtsrc in rtsc)
                {
                    if (rtsrc.Text == TabName)
                    {
                        rts = rtsrc;
                        tabEID = rts.ElementID;
                        break;
                    }
                }
                if (rts != null)
                {
                    WSRibbonTabSourceReference wsrtsref = wsrr.FindTabReference(cs.MenuGroupName, tabEID);
                    if (wsrtsref != null)
                    {
                        res = true;
                    }
                }
            }
            return res;
        }
        public static bool WorkspaceHasTab(CustomizationSection cs, string WorkspaceName, string TabName)
        {
            bool res = false;
            if (WorkspaceExists(WorkspaceName))
            {
                Workspace ws = cs.Workspaces[cs.Workspaces.IndexOfWorkspaceName(WorkspaceName)];
                WSRibbonRoot wsrr = ws.WorkspaceRibbonRoot;
                RibbonTabSource rts = null;
                string tabEID = string.Empty;
                RibbonTabSourceCollection rtsc = cs.MenuGroup.RibbonRoot.RibbonTabSources;
                foreach (RibbonTabSource rtsrc in rtsc)
                {
                    if (rtsrc.Text == TabName)
                    {
                        rts = rtsrc;
                        tabEID = rts.ElementID;
                        break;
                    }
                }
                if (rts != null)
                {
                    WSRibbonTabSourceReference wsrtsref = wsrr.FindTabReference(cs.MenuGroupName, tabEID);
                    if (wsrtsref != null)
                    {
                        res = true;
                    }
                }
            }
            return res;
        }

        public static bool TabExists(string TabName)
        {
            bool res = false;
            CustomizationSection cs = GetCS();
            RibbonTabSourceCollection rtsc = cs.MenuGroup.RibbonRoot.RibbonTabSources;
            foreach (RibbonTabSource rtsrc in rtsc)
            {
                if (rtsrc.Text == TabName)
                {
                    res = true;
                    break;
                }
            }
            return res;
        }
        public static bool TabExists(CustomizationSection cs, string TabName)
        {
            bool res = false;
            RibbonTabSourceCollection rtsc = cs.MenuGroup.RibbonRoot.RibbonTabSources;
            foreach (RibbonTabSource rtsrc in rtsc)
            {
                if (rtsrc.Text == TabName)
                {
                    res = true;
                    break;
                }
            }
            return res;
        }
        
        public static bool PanelExists(string PanelName)
        {
            bool res = false;
            if (cs == null) cs = GetCS();
            RibbonRoot ribRoot = cs.MenuGroup.RibbonRoot;
            RibbonPanelSourceCollection panels = ribRoot.RibbonPanelSources;
            foreach (RibbonPanelSource pan in panels)
            {
                if (pan.Text == PanelName)
                {
                    res = true;
                    break;
                }
            }
            return res;
        }
        public static RibbonPanelSource GetPanelByName(string PanelName)
        {
            RibbonPanelSource res = null;
            if (cs == null) cs = GetCS();
            RibbonRoot ribRoot = cs.MenuGroup.RibbonRoot;
            RibbonPanelSourceCollection panels = ribRoot.RibbonPanelSources;
            foreach (RibbonPanelSource pan in panels)
            {
                if (pan.Text == PanelName)
                {
                    res = pan;
                    break;
                }
            }
            return res;
        }

        private static CustomizationSection GetCS()
        {
            string mainCUIFile = (string)Application.GetSystemVariable("MENUNAME");
            mainCUIFile += ".cuix";
            return new CustomizationSection(mainCUIFile);
        }

        public static void saveCui(bool SetCurrentWorkspace, bool LoadPartials)
        {
            // Save all Changes made to the CUI file in this session. 
            // If changes were made to the Main CUI file - save it
            // If changes were made to teh Partial CUI files need to save them too

            if (cs.IsModified)
                cs.Save();

            string wsCurrent = GetCurrentWorkspaceName();

            // Here we unload and reload the main CUI file so the changes to the CUI file could take effect immediately.
            // To do this we P/Invoke into acedCmd() in order to synchronously call the CUIUNLOAD / CUILOAD command, disarming and
            // rearming the file dialog during the process.

            ResultBuffer rb = new ResultBuffer();
            // RTSTR = 5005
            rb.Add(new TypedValue(5005, "FILEDIA"));
            rb.Add(new TypedValue(5005, "0"));
            // start the insert command
            acedCmd(rb.UnmanagedObject);

            //CUIUNLOAD
            string cuiMenuGroup = cs.MenuGroup.Name;
            rb = new ResultBuffer();
            rb.Add(new TypedValue(5005, "_CUIUNLOAD"));
            rb.Add(new TypedValue(5005, cuiMenuGroup));
            acedCmd(rb.UnmanagedObject);

            //CUILOAD
            string cuiFileName = cs.CUIFileName;
            rb = new ResultBuffer();
            rb.Add(new TypedValue(5005, "_CUILOAD"));
            rb.Add(new TypedValue(5005, cuiFileName));
            acedCmd(rb.UnmanagedObject);

            if (LoadPartials)
            {
                // Load all the partial CUIX files, They are not loaded when the main cuix file
                // is loaded with command line version of CUILOAD
                CustomizationSection[] partials;
                //int numPartialFiles;
                string fileNameWithQuotes;
                partials = new CustomizationSection[cs.PartialCuiFiles.Count];
                ////int i = 0;
                foreach (string fileName in cs.PartialCuiFiles)
                {
                    if (File.Exists(fileName))
                    {
                        fileNameWithQuotes = "\"" + fileName + "\"";
                        //Application.DocumentManager.MdiActiveDocument.SendStringToExecute("cuiload " + fileNameWithQuotes + "\n", false, false, true);

                        rb = new ResultBuffer();
                        rb.Add(new TypedValue(5005, "_CUILOAD"));
                        rb.Add(new TypedValue(5005, fileNameWithQuotes));
                        acedCmd(rb.UnmanagedObject);
                    }
                }
            }

            if (SetCurrentWorkspace)
            {
                // Set the current workspace
                rb = new ResultBuffer();
                rb.Add(new TypedValue(5005, "_WSCURRENT"));
                string wsWithQuotes = "\"" + wsCurrent + "\"";
                rb.Add(new TypedValue(5005, wsCurrent));
                acedCmd(rb.UnmanagedObject);
            }


            //FILEDIA back on
            rb = new ResultBuffer();
            rb.Add(new TypedValue(5005, "FILEDIA"));
            rb.Add(new TypedValue(5005, "1"));
            acedCmd(rb.UnmanagedObject);
        }
        
        // Save all the open CUI Files that have been modified
        public static void saveCui(bool bSetWorkspace)
        {
            saveCui(bSetWorkspace, false);
            //Document doc = Application.DocumentManager.MdiActiveDocument;
            //// Save Changes to the CUIx file 
            //if (cs == null) cs = GetCS();
            //if (cs.IsModified)
            //    cs.Save();

            //// Here we unload and reload the main CUIx file so the changes
            //// to the CUIx file can take effect immediately.
            //string wsCurrent = GetCurrentWorkspaceName();
            //object fd = Application.GetSystemVariable("FILEDIA");
            //Application.SetSystemVariable("FILEDIA", 0);

            //string flName = cs.CUIFileBaseName;
            //RunCommand(doc, "_.CUIUNLOAD " + flName + " ");
            //RunCommand(doc, "_.CUILOAD " + flName + " ");
            //RunCommand(doc, "_FILEDIA " + fd.ToString() + " ");
            //RunCommand(doc, "_.WSCURRENT " + wsCurrent + "\n");
            //if (bSetWorkspace)
            //{
            //    RunCommand(doc, "_.WSCURRENT " + wsCurrent + "\n");
            //    RunCommand(doc, "_.RIBBON ");
            //}
            //Application.SetSystemVariable("FILEDIA", fd);
        }

        private static void RunCommand(Document doc, String cmdString)
        {
            doc.SendStringToExecute(cmdString, false, false, true);
        }

        public static string GetCurrentWorkspaceName()
        {
            string res = Application.GetSystemVariable("WSCURRENT").ToString();
            return res;
        }

        internal static System.Collections.Specialized.StringCollection GetAllPanelItemNames(string PanelName)
        {
            System.Collections.Specialized.StringCollection res = new System.Collections.Specialized.StringCollection();
            RibbonPanelSource rps = GetPanelByName(PanelName);
            if (rps != null)
            {
                RibbonItemCollection ric = rps.Items;
                string mess = string.Empty;
                foreach (RibbonItem ri in ric)
                {
                    GetItemNames(ref res, ri);
                }
                //for (int i = 0; i < res.Count; i++)
                //{
                //    mess += res[i] + Environment.NewLine;
                //}
                //System.Windows.Forms.MessageBox.Show(mess);
            }

            return res;
        }
        private static void GetItemNames(ref System.Collections.Specialized.StringCollection Names, RibbonItem RibItem)
        {
            // if its a row then add any buttons you find
            if (RibItem.GetType() == typeof(RibbonRow))
            {
                RibbonRow rr = RibItem as RibbonRow;
                if (rr != null)
                {
                    RibbonItemCollection rric = rr.Items;
                    foreach (RibbonItem rri in rric)
                    {
                        if (rri.GetType() == typeof(RibbonCommandButton))
                        {
                            RibbonCommandButton rb = rri as RibbonCommandButton;
                            if (rb.Text.Trim().Length > 0)
                            {
                                if (!Names.Contains(rb.Text.Trim()))
                                {
                                    Names.Add(rb.Text);
                                }
                            }
                        }
                        else
                        {
                            GetItemNames(ref Names, rri);
                        }
                    }
                }
            }
            else // if it's a panel then send it's items back through GetItemNames
            {
                if (RibItem.GetType() == typeof(RibbonRowPanel))
                {
                    RibbonRowPanel rps = RibItem as RibbonRowPanel;
                    foreach (RibbonItem ri in rps.Items)
                    {
                        GetItemNames(ref Names, ri);
                    }
                }
            }
        }
    }
}