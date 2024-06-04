using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace Copy_Form
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    //If you want to use an add-on identifier for the development license, you can specify an add-on identifier string as the second parameter.
                    //oApp = new Application(args[0], "XXXXX");
                    oApp = new Application(args[0]);
                }
                //Menu MyMenu = new Menu();
                //MyMenu.AddMenuItems();
                //oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                App = Application.SBO_Application;
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public Program()
        {
        }
        public static SAPbouiCOM.Form oForm;
        public static SAPbouiCOM.Form sForm;
        public static SAPbouiCOM.Item oItem;
        public static SAPbouiCOM.Item oNewItem;
        public static SAPbouiCOM.ChooseFromListCollection oCFLs = null;
        public static SAPbouiCOM.Conditions oCons = null;
        public static SAPbouiCOM.Condition oCon = null;
        public static SAPbouiCOM.ChooseFromList oCFL = null;
        public static SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
        public static SAPbouiCOM.Button oButton = null;
        public static SAPbouiCOM.EditText cItem = null;
        public static SAPbouiCOM.DBDataSource oDbDataSource = null;
        public static SAPbouiCOM.DataTable oDataTable,oDataTable1 = null;
        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbouiCOM.Matrix oMatrix = null;
        public static SAPbobsCOM.Recordset RecordSet = null;
        public static SAPbouiCOM.Application App = null;
        public static SAPbouiCOM.Item mItem = null;
        public static SAPbouiCOM.Columns oColumns = null;
        public static SAPbouiCOM.Column oColumn = null;
        public static SAPbobsCOM.BranchParams getlistParams = null;

        public static void AddBtn()
        {
            oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
            oCFLCreationParams.MultiSelection = true;
            oCFLCreationParams.ObjectType = "17";
            oCFLCreationParams.UniqueID = "CFL_3";
            oCFLs = oForm.ChooseFromLists;
            oCFL = oCFLs.Add(oCFLCreationParams);

            oNewItem = oForm.Items.Item("2");
            oItem = oForm.Items.Add("btnCopy", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Top = oNewItem.Top;
            oItem.Height = oNewItem.Height;
            oItem.Left = oNewItem.Left + oNewItem.Width + 5;
            oItem.Width = oNewItem.Width + 20;
            oButton = (SAPbouiCOM.Button)oItem.Specific;
            oButton.Caption = "Copy Form";
            oButton.ChooseFromListUID = "CFL_3";

           SAPbouiCOM.Item oNewItem1 = oForm.Items.Item("20");
            SAPbouiCOM.Item oNewItem2 = oForm.Items.Item("23");
            mItem = oForm.Items.Add("Matrix1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            mItem.Left = oNewItem1.Left+oNewItem1.Width+20;
            mItem.Width =oNewItem2.Left- mItem.Left;
            mItem.Top = oNewItem1.Top;
            mItem.Height = 200;
            oMatrix = ((SAPbouiCOM.Matrix)(mItem.Specific));
            oColumns = oMatrix.Columns;
            //*******************************************************************
            oColumn = oColumns.Add("No", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "No";
            oColumn.Width = mItem.Width * mItem.Width / 3;
            oColumn.Editable = true;
            //*******************************************************************
            oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "ItemCode";
            oColumn.Width = mItem.Width* mItem.Width/3;
            oColumn.Editable = true;
            //*******************************************************************
            oColumn = oColumns.Add("ItemName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "ItemName";
            oColumn.Width = mItem.Width* mItem.Width/3;
            oColumn.Editable = false;
            //*******************************************************************

            oButton.ChooseFromListAfter += OButton_ChooseFromListAfter;
            oButton.ChooseFromListBefore += OButton_ChooseFromListBefore;
        }
        //choose form list condition 
        public static void CFL_Condition()
        {
            oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item("CFL_3");
            cItem = (SAPbouiCOM.EditText)sForm.Items.Item("54").Specific;
            SAPbouiCOM.ConditionsClass emptyCon;
            emptyCon = new SAPbouiCOM.ConditionsClass();
            oCFL.SetConditions(emptyCon);
            oCons = oCFL.GetConditions();
            oCon = oCons.Add();
            oCon.BracketOpenNum = 2;
            oCon.Alias = "DocStatus";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "O";
            //oCon.BracketCloseNum = 1;
            //oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            //oCon = oCons.Add();
            //oCon.BracketOpenNum = 1;
            //oCon.Alias = "CardName";
            //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            //oCon.CondVal = cItem.Value.ToString();
            oCon.BracketCloseNum = 2;
            oCFL.SetConditions(oCons);
        }
        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if (pVal.FormType==10016 && pVal.ItemUID=="2" && pVal.Before_Action==true)
                {
                    oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    oForm.Close();
                }
                if ((pVal.FormType == 140 && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) && (pVal.Before_Action == true))
                {
                    oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    //Application.SBO_Application.MessageBox(oForm.UniqueID);
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) && (pVal.Before_Action == true))
                    {
                        oDataTable = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Add("SO");
                        oDataTable1 = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Add("DN");
                        sForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                        AddBtn();
                    }
                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
            }
            
        }
        
        private static void OButton_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                CFL_Condition();
            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
            }
        }

        private static void OButton_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new NotImplementedException();
            try
            {
                string Deliver;
                SAPbobsCOM.Documents Delivery = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                SAPbobsCOM.Documents SalesOrder = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                SAPbouiCOM.ISBOChooseFromListEventArg CFLEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                oCFL = oForm.ChooseFromLists.Item("CFL_3");
                for (int j = 0; j < CFLEvent.SelectedObjects.Rows.Count; j++)
                {
                    int SoDocEntry = 0;
                    SoDocEntry = Convert.ToInt32(CFLEvent.SelectedObjects.GetValue("DocEntry", j).ToString());
                    SoDocEntry = Convert.ToInt32(SoDocEntry);
                    Delivery.CardCode = CFLEvent.SelectedObjects.GetValue("CardCode", j).ToString();
                    Delivery.DocDate = DateTime.Today;//DateTime.Now;
                    Delivery.DocDueDate = DateTime.Today;//.AddDays(7);
                    Delivery.BPL_IDAssignedToInvoice =Convert.ToInt32(CFLEvent.SelectedObjects.GetValue("BPLId", j).ToString());
                    for (int i = 0; i < SalesOrder.Lines.Count; i++)
                        {
                            if (SalesOrder.GetByKey(SoDocEntry))
                            {
                                if (SalesOrder.Lines.LineStatus == 0)
                                {
                                    Delivery.Lines.BaseEntry = SoDocEntry;
                                    Delivery.Lines.BaseLine = i;
                                    Delivery.Lines.BaseType = 17;
                                    Delivery.Lines.SetCurrentLine(i);
                                    Delivery.Lines.Add();
                                }
                            }

                    }

                    int result = Delivery.Add();
                    if (result == 0)
                    {
                        oCompany.GetNewObjectCode(out Deliver);
                        Application.SBO_Application.SetStatusBarMessage($"Update successfully.{Delivery.Lines.BaseLine}");
                        return;
                    }
                    else
                    {
                        Application.SBO_Application.SetStatusBarMessage($"Failed to add delivery note: {oCompany.GetLastErrorDescription()}");
                    }
                }
                

            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
            }
            
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
