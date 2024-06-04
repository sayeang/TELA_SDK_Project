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
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
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
        public static SAPbouiCOM.DataTable oDataTable = null;
        public static void AddBtn()
        {
            oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
            oCFLCreationParams.MultiSelection = false;
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

            oButton.ChooseFromListAfter += OButton_ChooseFromListAfter;
            oButton.ChooseFromListBefore += OButton_ChooseFromListBefore;
        }
        //choose form list condition 
        public static void CFL_Condition()
        {
            //oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item("CFL_3");
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
                oCon.BracketCloseNum = 1;
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCon = oCons.Add();
                oCon.BracketOpenNum = 1;
                oCon.Alias = "CardName";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = cItem.Value.ToString();
                oCon.BracketCloseNum = 2;
                oCFL.SetConditions(oCons);
        }
        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if ((pVal.FormType == 140 && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) && (pVal.Before_Action == true))
                {
                    oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) && (pVal.Before_Action == true))
                    {
                        oDataTable = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Add("DN");
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
            SAPbobsCOM.Company oCompany=null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbobsCOM.Recordset RecordSet = null;
            try
            {
                    oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                    //SAPbobsCOM.Documents DN = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                    SAPbouiCOM.ISBOChooseFromListEventArg CFLEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                //oDbDataSource = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("DLN1");
                oDataTable.ExecuteQuery("SELECT A.\"ItemCode\",A.\"Dscription\" FROM \"DLN1\" A");
                    oCFL = oForm.ChooseFromLists.Item("CFL_3");
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    int DnDocEntry = 0;
                    DnDocEntry = Convert.ToInt32(CFLEvent.SelectedObjects.GetValue("DocEntry", 0).ToString());
                    string Query = "SELECT A.\"CardCode\",B.\"ItemCode\",B.\"Dscription\" FROM \"ORDR\" A LEFT OUTER JOIN \"RDR1\" B ON A.\"DocEntry\"=B.\"DocEntry\" WHERE A.\"DocEntry\"='" + DnDocEntry + "'";
                    RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    RecordSet.DoQuery(Query);
                    oMatrix.Columns.Item("1").DataBind.Bind("DN", "ItemCode");
                    oMatrix.Columns.Item("3").DataBind.Bind("DN", "Dscription");
                    Application.SBO_Application.MessageBox(DnDocEntry.ToString()+""+ RecordSet.Fields.Item("ItemCode").Value.ToString());
                    int i = 0;
                    while (!RecordSet.EoF)
                    {
                        oMatrix.FlushToDataSource();
                        //oDbDataSource.GetValue("ItemCode", i);
                        oDataTable.SetValue("ItemCode", i, RecordSet.Fields.Item("ItemCode").Value.ToString());
                        oDataTable.SetValue("Dscription", i, RecordSet.Fields.Item("Dscription").Value.ToString());
                        i = i + 1;
                        oMatrix.LoadFromDataSource();
                        oMatrix.AddRow(1, oMatrix.RowCount);
                        oMatrix.ClearRowData(oMatrix.RowCount);
                        RecordSet.MoveNext();
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet);
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
