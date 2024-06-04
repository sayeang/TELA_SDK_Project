using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM.Framework;

namespace SBOAddOn_TSA
{
    class Menu
    {
        //private SAPbouiCOM.Application SBO_Application;

        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;

            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            //oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

            //Declare Varible sPath that store directory 
            //string sPath = null;
            //sPath = System.Environment.CurrentDirectory + @"\bin\Debug\Biz-Logo16x16.bmp";
            //sPath = sPath.Remove(sPath.Length - 9, 9);
        
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "BIZD";
            oCreationPackage.String = "BIZ Add-on";
            //oCreationPackage.Image = sPath;
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

            try
            {
                //==========================================================
                //// Get the menu collection of the newly added pop-up item
                //oMenuItem = Application.SBO_Application.Menus.Item("BIZD");
                //oMenus = oMenuItem.SubMenus;

                ////Create s sub menu
                //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //oCreationPackage.UniqueID = "BIZD.Form1";
                //oCreationPackage.String = "Exemption";

                //oMenus.AddEx(oCreationPackage);

                //==========================================================

                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("BIZD");
                oMenus = oMenuItem.SubMenus;

                //Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "BIZD1";
                oCreationPackage.String = "Setup";
                oMenus.AddEx(oCreationPackage);

                oMenuItem = Application.SBO_Application.Menus.Item("BIZD1");
                oMenus = oMenuItem.SubMenus;

                //Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("BIZD1");
                oMenus = oMenuItem.SubMenus;

                //create the payroll menu item for master data
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BIZD.Form1";
                oCreationPackage.String = "General Authorization";

                oMenus.AddEx(oCreationPackage);
                //==========================================================
            }
            catch (Exception ex)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BIZD.Form1")
                {
                    NewForm_1_b1f actvieForm = new NewForm_1_b1f();
                    actvieForm.Show();                    
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
