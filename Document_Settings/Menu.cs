using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM.Framework;

namespace SBOAddonProject_Setting
{
    class Menu
    {
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "SBOAddonProject_Setting";
            oCreationPackage.String = "Document Setting";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception e)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("SBOAddonProject_Setting");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;

                ////Please replace following 2 "Form1" with real form class in current project
                oCreationPackage.UniqueID = "SBOAddonProject_Setting.Sale_Document";
                oCreationPackage.String = "Sale Document";
                oMenus.AddEx(oCreationPackage);

            }
            catch (Exception er)
            { //  Menu already exists
               
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short ,true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                ////Please replace following 3 "Form1" with real form class in current project
                if (pVal.BeforeAction && pVal.MenuUID == "SBOAddonProject_Setting.Sale_Document")
                {
                    Document_Settings_b1f activeForm = new Document_Settings_b1f();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
