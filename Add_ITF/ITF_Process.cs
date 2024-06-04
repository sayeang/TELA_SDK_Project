using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace Add_ITF
{
    class ITF_Process
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
                //Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
                Application.SBO_Application.StatusBar.SetSystemMessage("Add-On Connecting....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning, "", "", "", 0);
                oCompany = new SAPbobsCOM.Company();
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                Application.SBO_Application.StatusBar.SetSystemMessage("Connecting Add-on successfully.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        #region Declare
        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbobsCOM.Documents GoodIssue, GoodReceipt = null;
        public static SAPbobsCOM.StockTransfer Transfer = null;
        public static SAPbouiCOM.Form oForm = null;
        public static SAPbouiCOM.Item oItem, oNewItem = null;
        public static SAPbouiCOM.Button oButton = null;
        public static SAPbouiCOM.Matrix oMatrix = null;
        public static SAPbobsCOM.Recordset oRS_Ref, oRS = null;
        public static SAPbouiCOM.Conditions oCons = null;
        public static SAPbouiCOM.Condition oCon = null;
        public static SAPbouiCOM.ChooseFromList oCFL = null;
        public static SAPbouiCOM.EditText Atten_teminal = null;
        public static SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
        public static string GI_Entry = null;
        #endregion
        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.FormTypeEx == "940"
                && BusinessObjectInfo.Type == "67" && BusinessObjectInfo.ActionSuccess)
            {
                Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                oRS_Ref = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                Transfer = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                GoodIssue = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                GoodReceipt = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

                #region GI_Entry Transfer Copy to Goods Issue Copy 
                int offset = oForm.DataSources.DBDataSources.Item(0).Offset;
                string Entry = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", offset);
                int DocEntry = 0;
                DocEntry = int.Parse(Entry);
                if (Transfer.GetByKey(DocEntry))
                {
                    if (Transfer.UserFields.Fields.Item("U_tl_itr_type").Value.ToString() == "Store Request")
                    {
                        GoodIssue.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                        GoodIssue.UserFields.Fields.Item("U_tl_cardcode").Value = Transfer.CardCode;
                        GoodIssue.DocDate = DateTime.Today;
                        GoodIssue.DocDueDate = DateTime.Today;
                        GoodIssue.TaxDate = DateTime.Today;
                        GoodIssue.BPL_IDAssignedToInvoice = Transfer.BPLID;
                        GoodIssue.UserFields.Fields.Item("U_tl_whsdesc").Value = Transfer.ToWarehouse;
                        GoodIssue.UserFields.Fields.Item("U_tl_gitwhs").Value = Transfer.UserFields.Fields.Item("U_tl_gitwhs").Value.ToString();
                        GoodIssue.UserFields.Fields.Item("U_tl_gitstore").Value = Transfer.UserFields.Fields.Item("U_tl_gitstore").Value.ToString();
                        string BinAbsEntry = null;
                        //oRS.DoQuery($"SELECT A.\"ToBinCode\",B.\"AbsEntry\" FROM OWTR A " +
                        //    $"LEFT OUTER JOIN OBIN B ON A.\"ToBinCode\" = B.\"BinCode\" AND A.\"ToWhsCode\" = B.\"WhsCode\" WHERE A.\"DocEntry\" = {DocEntry} AND A.\"ToBinCode\" IS NOT NULL");

                        for (int i = 0; i < Transfer.Lines.Count; i++)
                        {
                            Transfer.Lines.SetCurrentLine(i);
                            //GoodReceipt.Lines.BaseEntry = DocEntry;
                            //GoodReceipt.Lines.BaseType = 60;
                            //GoodIssue.Lines.UoMEntry = Transfer.Lines.UoMEntry;
                            GoodIssue.Lines.BaseLine = Transfer.Lines.LineNum;
                            GoodIssue.Lines.ItemCode = Transfer.Lines.ItemCode;
                            GoodIssue.Lines.ItemDescription = Transfer.Lines.ItemDescription;
                            GoodIssue.Lines.Quantity = Transfer.Lines.Quantity;

                            GoodIssue.Lines.WarehouseCode = Transfer.ToWarehouse;
                            oRS.DoQuery($"SELECT A.\"DocEntry\",A.\"ItemCode\",B.\"AbsEntry\",A.\"FisrtBin\",A.\"LineNum\" FROM WTR1 A " +
                                $"LEFT OUTER JOIN OBIN B ON A.\"WhsCode\" = B.\"WhsCode\" and A.\"FisrtBin\" = B.\"BinCode\" " +
                                $"WHERE A.\"DocEntry\" = {DocEntry} And A.\"LineNum\" = {i}");
                            BinAbsEntry = oRS.Fields.Item("AbsEntry").Value.ToString();
                            if (oRS.RecordCount == 1)
                            {
                                GoodIssue.Lines.BinAllocations.BinAbsEntry = int.Parse(BinAbsEntry);
                                GoodIssue.Lines.BinAllocations.Quantity = Transfer.Lines.Quantity;
                            }
                            GoodIssue.Lines.Add();
                            Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                        int Add = GoodIssue.Add();
                        if (Add == 0)
                        {
                            oCompany.GetNewObjectCode(out GI_Entry);
                            if (Transfer.GetByKey(DocEntry))
                            {
                                oRS_Ref.DoQuery($"SELECT \"DocEntry\",\"LineNum\",\"RefDocEntr\",\"RefDocNum\",\"RefObjType\" From \"WTR21\" WHERE \"DocEntry\"={DocEntry}");
                                int LineNum = int.Parse(oRS_Ref.Fields.Item("LineNum").Value.ToString());
                                int RefEntry = 0;
                                for (int i = 0; i <= oRS_Ref.RecordCount; i++)
                                {
                                    Transfer.DocumentReferences.SetCurrentLine(i);
                                    if (i == oRS_Ref.RecordCount)
                                    {
                                        Transfer.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_GoodsIssue;
                                        Transfer.DocumentReferences.ReferencedDocEntry = int.Parse(GI_Entry);
                                        Transfer.DocumentReferences.Add();
                                    }
                                    else
                                    {
                                        int RefObj = int.Parse(oRS_Ref.Fields.Item("RefObjType").Value.ToString());
                                        Transfer.DocumentReferences.ReferencedObjectType = (SAPbobsCOM.ReferencedObjectTypeEnum)RefObj;
                                        RefEntry = int.Parse(oRS_Ref.Fields.Item("RefDocEntr").Value.ToString());
                                        Transfer.DocumentReferences.ReferencedDocEntry = RefEntry;
                                        Transfer.DocumentReferences.Add();
                                        oRS_Ref.MoveNext();
                                    }
                                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                }
                                Transfer.Update();
                                //Application.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID).Select();
                                //Application.SBO_Application.ActivateMenuItem("1304");
                            }
                            Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsIssue, "", GI_Entry);
                            Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                        }
                        else
                        {
                            Application.SBO_Application.SetStatusBarMessage($"{oCompany.GetLastErrorDescription()}");
                            Application.SBO_Application.MessageBox($"Goods Issue:{oCompany.GetLastErrorDescription()}");
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }
                }
                #endregion

                #region Goods Issue Copy To Goods Receipt 
                Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                if (GoodIssue.GetByKey(int.Parse(GI_Entry)))
                {
                    GoodReceipt.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                    GoodReceipt.UserFields.Fields.Item("U_tl_cardcode").Value = GoodIssue.UserFields.Fields.Item("U_tl_cardcode").Value.ToString();
                    GoodReceipt.DocDate = DateTime.Today;
                    GoodReceipt.DocDueDate = DateTime.Today;
                    GoodReceipt.TaxDate = DateTime.Today;
                    GoodReceipt.BPL_IDAssignedToInvoice = GoodIssue.BPL_IDAssignedToInvoice;
                    GoodReceipt.UserFields.Fields.Item("U_tl_whsdesc").Value = GoodIssue.UserFields.Fields.Item("U_tl_whsdesc").Value;
                    for (int i = 0; i < GoodIssue.Lines.Count; i++)
                    {
                        GoodIssue.Lines.SetCurrentLine(i);
                        //GoodReceipt.Lines.BaseEntry = DocEntry;
                        //GoodReceipt.Lines.BaseType = 60;
                        GoodReceipt.Lines.BaseLine = GoodIssue.Lines.BaseLine;
                        GoodReceipt.Lines.ItemCode = GoodIssue.Lines.ItemCode;
                        GoodReceipt.Lines.ItemDescription = GoodIssue.Lines.ItemDescription;
                        GoodReceipt.Lines.Quantity = GoodIssue.Lines.Quantity;
                        GoodReceipt.Lines.WarehouseCode = GoodIssue.UserFields.Fields.Item("U_tl_gitwhs").Value.ToString();//GoodIssue.Lines.WarehouseCode;
                        GoodReceipt.Lines.BinAllocations.BinAbsEntry =int.Parse(GoodIssue.UserFields.Fields.Item("U_tl_gitstore").Value.ToString());//GoodIssue.Lines.BinAllocations.BinAbsEntry;

                        GoodReceipt.Lines.BinAllocations.Quantity = GoodIssue.Lines.Quantity;
                        GoodReceipt.Lines.Add();
                        Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                    string GR_Entry = null;
                    int Add = GoodReceipt.Add();
                    if (Add == 0)
                    {
                        oCompany.GetNewObjectCode(out GR_Entry);
                        //Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                        if (GoodIssue.GetByKey(int.Parse(GI_Entry)))
                        {
                            oRS_Ref.DoQuery($"SELECT \"DocEntry\",\"LineNum\",\"RefDocEntr\",\"RefDocNum\",\"RefObjType\" From \"IGE21\" WHERE \"DocEntry\"={int.Parse(GI_Entry)}");
                            int LineNum = int.Parse(oRS_Ref.Fields.Item("LineNum").Value.ToString());
                            int RefEntry = 0;
                            for (int i = 0; i <= oRS_Ref.RecordCount; i++)
                            {
                                GoodIssue.DocumentReferences.SetCurrentLine(i);
                                if (i == oRS_Ref.RecordCount)
                                {
                                    GoodIssue.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_GoodsReceipt;
                                    GoodIssue.DocumentReferences.ReferencedDocEntry = int.Parse(GR_Entry);
                                    GoodIssue.DocumentReferences.Add();
                                }
                                else
                                {
                                    int RefObj = int.Parse(oRS_Ref.Fields.Item("RefObjType").Value.ToString());
                                    GoodIssue.DocumentReferences.ReferencedObjectType = (SAPbobsCOM.ReferencedObjectTypeEnum)RefObj;
                                    RefEntry = int.Parse(oRS_Ref.Fields.Item("RefDocEntr").Value.ToString());
                                    GoodIssue.DocumentReferences.ReferencedDocEntry = RefEntry;
                                    GoodIssue.DocumentReferences.Add();
                                    oRS_Ref.MoveNext();
                                }
                                Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                            GoodIssue.Update();
                            //Application.SBO_Application.Forms.GetFormByTypeAndCount(720, pVal.FormTypeCount).Select();
                            //Application.SBO_Application.ActivateMenuItem("1304");
                        }
                        Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsReceipt, oCompany.GetNewObjectType(), GR_Entry);
                        Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                    }
                    else
                    {
                        Application.SBO_Application.SetStatusBarMessage($"{oCompany.GetLastErrorDescription()}");
                        Application.SBO_Application.MessageBox($"Goods Receipt:{oCompany.GetLastErrorDescription()}");
                        return;
                    }
                }
            }
            #endregion
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
