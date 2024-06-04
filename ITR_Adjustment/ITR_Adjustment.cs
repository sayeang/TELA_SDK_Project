using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace ITR_Adjustment
{
    class ITR_Adjustment
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
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
                Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
                Application.SBO_Application.StatusBar.SetSystemMessage("Add-On Connecting....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning, "", "", "", 0);
                oCompany = new SAPbobsCOM.Company();
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                Application.SBO_Application.StatusBar.SetSystemMessage("Connecting Add-on successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        #region 'Declare'
        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbobsCOM.Documents GoodIssue, GoodReceipt = null;
        public static SAPbobsCOM.StockTransfer TransferRequest, Transfer, TransferRequest_Draft = null;
        //public static SAPbobsCOM.Documents TransferRequest_Draft = null;
        public static SAPbouiCOM.Form oForm = null;
        public static SAPbouiCOM.Item oItem, oNewItem = null;
        public static SAPbouiCOM.Button oButton = null;
        public static SAPbouiCOM.Matrix oMatrix = null;
        public static SAPbobsCOM.Recordset oRS_Ref, oRS, oRS1 = null;
        public static SAPbouiCOM.Conditions oCons = null;
        public static SAPbouiCOM.Condition oCon = null;
        public static SAPbouiCOM.ChooseFromList oCFL = null;
        public static SAPbouiCOM.EditText Atten_teminal = null;
        public static SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
        public static string GI_Entry, BinAbsEntry, GR_Entry = null;
        #endregion

        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.ActionSuccess &&
                    ((BusinessObjectInfo.FormTypeEx == "1250000940" && BusinessObjectInfo.Type == "1250000001") || (BusinessObjectInfo.FormTypeEx == "720" && BusinessObjectInfo.Type == "60")))
                {
                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    #region 'Connect to Company'
                    try
                    {
                        oCompany.Disconnect();

                        oCompany = new SAPbobsCOM.Company();
                        oCompany.Server = "HDB@192.168.1.11:30013";
                        oCompany.SLDServer = "192.168.1.11:40001";
                        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                        oCompany.CompanyDB = "TLTELA_DEVELOPER_TECH1";
                        oCompany.UserName = "manager";
                        oCompany.Password = "1234";
                        oCompany.DbUserName = "SYSTEM";
                        oCompany.DbPassword = "Biz@2022";
                        //oCompany.Connect();
                        int iCon = oCompany.Connect();
                        if (iCon != 0)
                        {
                            Application.SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
                            return;
                        }
                    
                    #endregion

                    #region 'Declare Object'
                    oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                    oRS_Ref = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    TransferRequest = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);
                    GoodIssue = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    GoodReceipt = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                    int offset = oForm.DataSources.DBDataSources.Item(0).Offset;
                    string Entry = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", offset);
                    int DocEntry = 0;
                    DocEntry = int.Parse(Entry);
                    #endregion

                    #region 'GIT Adjustment'
                    if (BusinessObjectInfo.FormTypeEx == "1250000940" && BusinessObjectInfo.Type == "1250000001")
                    {
                        #region Inventory TransferRequest Copy To Goods Issue
                        if (TransferRequest.GetByKey(DocEntry))
                        {
                            if (TransferRequest.UserFields.Fields.Item("U_tl_itr_type").Value.ToString() == "Adjustment")
                            {
                                Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                #region 'Goods Issue Header'
                                GoodIssue.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                                GoodIssue.UserFields.Fields.Item("U_tl_cardcode").Value = TransferRequest.CardCode;
                                GoodIssue.DocDate = DateTime.Today;
                                GoodIssue.DocDueDate = DateTime.Today;
                                GoodIssue.TaxDate = DateTime.Today;
                                GoodIssue.BPL_IDAssignedToInvoice = TransferRequest.BPLID;
                                GoodIssue.UserFields.Fields.Item("U_tl_whsdesc").Value = TransferRequest.ToWarehouse;
                                GoodIssue.UserFields.Fields.Item("U_tl_gitwhs").Value = TransferRequest.UserFields.Fields.Item("U_tl_gitwhs").Value.ToString();
                                GoodIssue.UserFields.Fields.Item("U_tl_gitstore").Value = TransferRequest.UserFields.Fields.Item("U_tl_gitstore").Value.ToString();
                                #endregion

                                #region 'Goods Issue Line'
                                for (int i = 0; i < TransferRequest.Lines.Count; i++)
                                {
                                    TransferRequest.Lines.SetCurrentLine(i);
                                    //GoodReceipt.Lines.BaseEntry = DocEntry;
                                    //GoodReceipt.Lines.BaseType = 60;
                                    //GoodIssue.Lines.UoMEntry = Transfer.Lines.UoMEntry;
                                    GoodIssue.Lines.BaseLine = TransferRequest.Lines.LineNum;
                                    GoodIssue.Lines.ItemCode = TransferRequest.Lines.ItemCode;
                                    GoodIssue.Lines.ItemDescription = TransferRequest.Lines.ItemDescription;
                                    GoodIssue.Lines.Quantity = TransferRequest.Lines.Quantity;

                                    GoodIssue.Lines.WarehouseCode = TransferRequest.ToWarehouse;
                                    BinAbsEntry = TransferRequest.UserFields.Fields.Item("U_tl_gitstore").Value.ToString();
                                    GoodIssue.Lines.BinAllocations.BinAbsEntry = int.Parse(BinAbsEntry);
                                    GoodIssue.Lines.BinAllocations.Quantity = TransferRequest.Lines.Quantity;
                                    GoodIssue.Lines.Add();
                                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                }
                                #endregion
                                int Add = GoodIssue.Add();
                                if (Add == 0)
                                {
                                    oCompany.GetNewObjectCode(out GI_Entry);
                                    if (TransferRequest.GetByKey(DocEntry))
                                    {
                                        oRS_Ref.DoQuery($"SELECT \"DocEntry\",\"LineNum\",\"RefDocEntr\",\"RefDocNum\",\"RefObjType\" From \"WTQ21\" WHERE \"DocEntry\"={DocEntry}");
                                        int LineNum = int.Parse(oRS_Ref.Fields.Item("LineNum").Value.ToString());
                                        int RefEntry = 0;
                                        for (int i = 0; i <= oRS_Ref.RecordCount; i++)
                                        {
                                            TransferRequest.DocumentReferences.SetCurrentLine(i);
                                            if (i == oRS_Ref.RecordCount)
                                            {
                                                TransferRequest.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_GoodsIssue;
                                                TransferRequest.DocumentReferences.ReferencedDocEntry = int.Parse(GI_Entry);
                                                TransferRequest.DocumentReferences.Add();
                                            }
                                            else
                                            {
                                                int RefObj = int.Parse(oRS_Ref.Fields.Item("RefObjType").Value.ToString());
                                                TransferRequest.DocumentReferences.ReferencedObjectType = (SAPbobsCOM.ReferencedObjectTypeEnum)RefObj;
                                                RefEntry = int.Parse(oRS_Ref.Fields.Item("RefDocEntr").Value.ToString());
                                                TransferRequest.DocumentReferences.ReferencedDocEntry = RefEntry;
                                                TransferRequest.DocumentReferences.Add();
                                                oRS_Ref.MoveNext();
                                            }
                                            Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        }
                                        TransferRequest.Close();
                                        TransferRequest.Update();
                                    }
                                    Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsIssue, "", GI_Entry);
                                    Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                                }
                                else
                                {
                                    Application.SBO_Application.SetStatusBarMessage($"{oCompany.GetLastErrorDescription()}");
                                    Application.SBO_Application.MessageBox($"Failed to create \"Goods Issue\" and \"Goods Receipt\".");
                                    Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_StockTransfersRequest, "", oCompany.GetNewObjectKey());
                                    oCompany.Disconnect();
                                    return;
                                }
                            }
                            else
                            {
                                Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_StockTransfersRequest, "", oCompany.GetNewObjectKey());
                                return;
                            }
                        }
                        #endregion
                        #region Goods Issue Copy To Goods Receipt 
                        if (GoodIssue.GetByKey(int.Parse(GI_Entry)))
                        {
                            Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
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
                                GoodReceipt.Lines.BinAllocations.BinAbsEntry = int.Parse(GoodIssue.UserFields.Fields.Item("U_tl_gitstore").Value.ToString());//GoodIssue.Lines.BinAllocations.BinAbsEntry;

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
                        #endregion
                    }
                    #endregion

                    #region 'Product Return'
                    if (BusinessObjectInfo.FormTypeEx == "720" && BusinessObjectInfo.Type == "60")
                    {
                        #region Goods Issue Copy To Goods Receipt 
                        if (GoodIssue.GetByKey(DocEntry))
                        {
                            if (GoodIssue.UserFields.Fields.Item("U_tl_gi_type").Value.ToString() == "Product Return")
                            {
                                Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                #region 'Goods Receipt Header'
                                GoodReceipt.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                                GoodReceipt.DocDate = DateTime.Today;
                                GoodReceipt.DocDueDate = DateTime.Today;
                                GoodReceipt.TaxDate = DateTime.Today;
                                GoodReceipt.BPL_IDAssignedToInvoice = GoodIssue.BPL_IDAssignedToInvoice;
                                GoodReceipt.UserFields.Fields.Item("U_tl_attn_ter").Value = GoodIssue.UserFields.Fields.Item("U_tl_attn_ter").Value;
                                #endregion

                                #region 'Goods Receipt Line'
                                for (int i = 0; i < GoodIssue.Lines.Count; i++)
                                {
                                    GoodIssue.Lines.SetCurrentLine(i);
                                    //GoodReceipt.Lines.BaseEntry = DocEntry;
                                    //GoodReceipt.Lines.BaseType = 60;
                                    //GoodReceipt.Lines.BaseLine = i;
                                    GoodReceipt.Lines.ItemCode = GoodIssue.Lines.ItemCode;
                                    GoodReceipt.Lines.ItemDescription = GoodIssue.Lines.ItemDescription;
                                    GoodReceipt.Lines.WarehouseCode = GoodIssue.UserFields.Fields.Item("U_tl_attn_ter").Value.ToString();
                                    GoodReceipt.Lines.Quantity = GoodIssue.Lines.Quantity;
                                    GoodReceipt.Lines.BinAllocations.BinAbsEntry = GoodIssue.Lines.BinAllocations.BinAbsEntry;
                                    GoodReceipt.Lines.BinAllocations.Quantity = GoodIssue.Lines.Quantity;
                                    GoodReceipt.Lines.Add();

                                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                }
                                #endregion

                                #region 'Add Goods Receipt'
                                int Add = GoodReceipt.Add();
                                if (Add == 0)
                                {
                                    oCompany.GetNewObjectCode(out GR_Entry);
                                    if (GoodIssue.GetByKey(DocEntry))
                                    {
                                        oRS_Ref.DoQuery($"SELECT \"DocEntry\",\"LineNum\",\"RefDocEntr\",\"RefDocNum\",\"RefObjType\" From \"IGE21\" WHERE \"DocEntry\"={DocEntry}");
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
                                    }
                                    Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsReceipt, oCompany.GetNewObjectType(), GR_Entry);
                                    Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                                }
                                else
                                {
                                    Application.SBO_Application.SetStatusBarMessage($"{oCompany.GetLastErrorDescription()}");
                                    Application.SBO_Application.MessageBox($"Failed to create \"Goods Receipt\".");
                                    Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsIssue, "", oCompany.GetNewObjectKey());
                                    return;
                                }
                                #endregion
                            }
                            else
                            {
                                Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_StockTransfersRequest, "", oCompany.GetNewObjectKey());
                                return;
                            }
                        }
                        #endregion
                    }
                        #endregion
                        oCompany.Disconnect();
                    }
                    catch (Exception ex)
                    {
                        Application.SBO_Application.MessageBox(ex.Message);
                        return;
                    }
                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
                return;
            }

        }

        #region 'Add Button'
        public static void AddBtn()
        {
            oNewItem = oForm.Items.Item("1250000073");
            oItem = oForm.Items.Add("btnCopy", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Top = oNewItem.Top;
            oItem.Height = oNewItem.Height;
            oItem.Left = oNewItem.Left - (oNewItem.Width) - 20;
            oItem.Width = oNewItem.Width;
            oButton = (SAPbouiCOM.Button)oItem.Specific;
            oButton.Caption = "Adjustment";
            oItem.Visible = true;
            oItem.Enabled = oNewItem.Enabled;//false;
            oItem.AffectsFormMode = true;
            //if (oItem.Enabled == true)
            //{
            //    oButton.ClickAfter += OButton_ClickAfter;
            //}
            //*******************************************************************
        }
        #endregion

        private static void OButton_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new NotImplementedException();
            try
            {
                oItem = (SAPbouiCOM.Item)oForm.Items.Item("btnCopy");
                if (oItem.Enabled == true)
                {
                    #region 'ITR CopyTo ITR'

                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    System.Windows.Forms.MessageBox.Show($"UserName: {oCompany.UserName}");
                    oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                    System.Windows.Forms.MessageBox.Show($"UserName: {oCompany.UserName}");
                    #region 'Assign Object'
                    oRS_Ref = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    TransferRequest = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);
                    TransferRequest_Draft = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransferDraft);
                    string xmlStockTransfer = string.Empty;
                    int offset = oForm.DataSources.DBDataSources.Item(0).Offset;
                    string Entry = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", offset);
                    int DocEntry = 0;
                    DocEntry = int.Parse(Entry);
                  var User=  oCompany.UserName.ToString();
                    #endregion

                    if (TransferRequest.GetByKey(DocEntry))
                    {
                        Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        #region 'TransferRequest Header'
                        TransferRequest_Draft.UserFields.Fields.Item("U_tl_itr_type").Value = "Adjustment";
                        TransferRequest_Draft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest;
                        TransferRequest_Draft.DocDate = DateTime.Today;
                        TransferRequest_Draft.TaxDate = DateTime.Today;
                        TransferRequest_Draft.CardCode = TransferRequest.CardCode;
                        TransferRequest_Draft.FromWarehouse= TransferRequest.FromWarehouse;
                        TransferRequest_Draft.ToWarehouse = TransferRequest.ToWarehouse;
                        #endregion

                        #region 'TransferRequest Line'
                        oRS.DoQuery($"SELECT \"LineStatus\",\"LineNum\" FROM \"WTQ1\" WHERE \"DocEntry\" = {DocEntry} and \"LineStatus\"='C' ORDER BY \"LineNum\"");
                        for (int i = 0; i < oRS.RecordCount; i++)
                        {
                            int LineNum = int.Parse(oRS.Fields.Item("LineNum").Value.ToString());
                            TransferRequest.Lines.SetCurrentLine(LineNum - i);
                            TransferRequest.Lines.Delete();
                            oRS.MoveNext();
                        }
                        for (int i = 0; i < TransferRequest.Lines.Count; i++)
                        {
                            TransferRequest.Lines.SetCurrentLine(i);
                            TransferRequest_Draft.Lines.BaseLine = TransferRequest.Lines.LineNum;
                            TransferRequest_Draft.Lines.BaseType =SAPbobsCOM.InvBaseDocTypeEnum.InventoryTransferRequest;
                            TransferRequest_Draft.Lines.BaseEntry = TransferRequest.Lines.DocEntry;
                            oRS.DoQuery($"SELECT \"OpenQty\" FROM WTQ1 WHERE \"DocEntry\"={DocEntry} and \"LineNum\"={TransferRequest.Lines.LineNum}");
                            double OpenQty = double.Parse(oRS.Fields.Item("OpenQty").Value.ToString());
                            TransferRequest_Draft.Lines.Quantity = OpenQty;
                            TransferRequest_Draft.Lines.Add();
                        }
                        #endregion

                        #region 'Add Transfer Request'
                        int Add = TransferRequest_Draft.Add();
                        if (Add != 0)
                        {
                            Application.SBO_Application.SetStatusBarMessage($"Failed to add Inventory Transfer Request:[{oCompany.GetLastErrorDescription()}]");
                            return;
                        }
                        else
                        {
                            oCompany.GetNewObjectCode(out string Inventory);
                            string obj=oCompany.GetNewObjectType();
                            Application.SBO_Application.OpenForm((SAPbouiCOM.BoFormObjectEnum)112, "", Inventory);
                            Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                        }
                        #endregion
                    }
                    #endregion
                }
                else
                {
                    return;
                }

            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
                return;
            }
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if (pVal.FormType == 1250000940 && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD && pVal.Before_Action)
                {
                    oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) && pVal.Before_Action)
                    {
                        AddBtn();
                    }
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "btnCopy" && !pVal.Action_Success && !pVal.InnerEvent)
                    {
                        try
                        {
                           // SAPbouiCOM.EditText txtAdj = (SAPbouiCOM.EditText)oForm.Items.Item("U_tl_isadjust").Specific;
                           // txtAdj.Value = "Y";
                            oItem = (SAPbouiCOM.Item)oForm.Items.Item("btnCopy");
                            oButton = (SAPbouiCOM.Button)oItem.Specific;
                            oButton.ClickAfter += OButton_ClickAfter;
                        }
                        catch (Exception er)
                        {
                            Application.SBO_Application.MessageBox(er.Message);
                            return;
                        }
                    }
                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage(er.Message);
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
