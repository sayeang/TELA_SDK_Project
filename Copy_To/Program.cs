
using SAPbouiCOM.Framework;
using System;

namespace Copy_To
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        #region static void Main(string[] args)
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
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
                Application.SBO_Application.StatusBar.SetSystemMessage("Add-On Connecting....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning, "", "", "", 0);
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                Application.SBO_Application.StatusBar.SetSystemMessage("Add-On Connected Successfully.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
                oApp.Run();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                System.Windows.Forms.MessageBox.Show($"{ex.Message}");
            }
        }
        #endregion
        #region Declare
        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbobsCOM.Documents GoodIssue, GoodReceipt = null;
        public static SAPbobsCOM.StockTransfer TransferRequest,Transfer = null;
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
        #endregion
        public Program()
        {
            try
            {
                
            }
            catch (Exception er)
            {
                System.Windows.Forms.MessageBox.Show(er.Message);
            }
            
        }
        #region  SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if (((pVal.FormType == 720 || pVal.FormType == 1250000940 || pVal.FormType == 940) && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) && (pVal.Before_Action == true))
                {
                    oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) && (pVal.Before_Action == true))
                    {
                        AddBtn();
                        if (pVal.FormType == 1250000940)
                        {
                            oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                            Add_Attenterminal();
                        }
                    }

                    if (pVal.FormType == 81)
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("19").Specific;
                        try
                        {
                            oItem = (SAPbouiCOM.Item)oForm.Items.Item("btnCopy");
                        }
                        catch (Exception)
                        {
                            AddBtn();
                        }
                        SAPbouiCOM.Item Picked = (SAPbouiCOM.Item)oForm.Items.Item("8");
                        SAPbouiCOM.StaticText BtnPick = (SAPbouiCOM.StaticText)oForm.Items.Item("8").Specific;
                        if (oForm.PaneLevel == 4)
                        {
                            oItem.Visible = true;
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("19").Specific;
                            if (pVal.ColUID == "1" && pVal.ItemUID == "19" && pVal.Row <= oMatrix.RowCount)
                            {
                                oForm.Freeze(true);
                                for (int i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    SAPbouiCOM.CheckBox Check = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                                    if (Check.Checked)
                                    {
                                        oItem.Enabled = true;
                                        break;
                                    }
                                    else
                                    {
                                        oItem.Enabled = false;
                                    }
                                }
                                oForm.Freeze(false);
                            }
                        }
                        else
                        {
                            oItem.Enabled = false;
                            oItem.Visible = false;
                        }

                    }
                    else
                    {
                        try
                        {
                            if ((pVal.FormType == 720 || pVal.FormType == 1250000940 || pVal.FormType == 940))
                            {
                                oItem = (SAPbouiCOM.Item)oForm.Items.Item("btnCopy");
                                oItem.Enabled = true;
                            }
                        }
                        catch (Exception)
                        {
                            AddBtn();
                        }
                    }
                    
                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
            }
        }
        #endregion
        #region Atten_teminal_ChooseFromListAfter
        private static void Atten_teminal_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new NotImplementedException();
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg CFLValue = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                Atten_teminal.Value = CFLValue.SelectedObjects.GetValue("WhsCode", 0).ToString();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }
        #endregion
        #region Atten_teminal_ChooseFromListBefore
        private static void Atten_teminal_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item(Atten_teminal.ChooseFromListUID);
                SAPbouiCOM.ConditionsClass emptyCon;
                emptyCon = new SAPbouiCOM.ConditionsClass();
                oCFL.SetConditions(emptyCon);
                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "U_tl_attn_ter";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "Y";
                oCFL.SetConditions(oCons);
            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
            }
        }
        #endregion
        #region Add Atten_terminal
        public static void Add_Attenterminal()
        {
            try
            {
                oItem = (SAPbouiCOM.Item)oForm.Items.Add("txtAtten", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Width = oForm.Items.Item("1470000099").Width;
                oItem.Height = oForm.Items.Item("1470000099").Height;
                oItem.Left = oForm.Items.Item("1470000099").Left;
                oItem.Top = oForm.Items.Item("1470000099").Top + oForm.Items.Item("1470000099").Height + 5;
                SAPbouiCOM.StaticText txtAttn = (SAPbouiCOM.StaticText)oItem.Specific;
                txtAttn.Caption = "Atten terminal";

                oItem = (SAPbouiCOM.Item)oForm.Items.Add("Attn", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Width = oForm.Items.Item("1470000101").Width;
                oItem.Height = oForm.Items.Item("1470000101").Height;
                oItem.Left = oForm.Items.Item("1470000101").Left;
                oItem.Top = oForm.Items.Item("1470000101").Top + oForm.Items.Item("1470000101").Height + 5;
                Atten_teminal = (SAPbouiCOM.EditText)oItem.Specific;
                Atten_teminal.DataBind.SetBound(true, "OWTQ", "U_tl_attn_ter");

                oItem = (SAPbouiCOM.Item)oForm.Items.Add("link", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                oItem.Width = oForm.Items.Item("1470000102").Width;
                oItem.Height = oForm.Items.Item("1470000102").Height;
                oItem.Left = oForm.Items.Item("1470000102").Left;
                oItem.Top = oForm.Items.Item("1470000102").Top + oForm.Items.Item("1470000101").Height + 5;
                SAPbouiCOM.LinkedButton link_Attn = (SAPbouiCOM.LinkedButton)oItem.Specific;
                link_Attn.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Warehouses;
                link_Attn.Item.LinkTo = "Attn";

                oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFL_1";
                oCFL = oForm.ChooseFromLists.Add(oCFLCreationParams);
                Atten_teminal.ChooseFromListUID = oCFLCreationParams.UniqueID;
                Atten_teminal.ChooseFromListAlias = "WhsCode";
                Atten_teminal.ChooseFromListBefore += Atten_teminal_ChooseFromListBefore;
                Atten_teminal.ChooseFromListAfter += Atten_teminal_ChooseFromListAfter;
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage(er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        #endregion
        #region Add Button()
        public static void AddBtn()
        {
            
            oNewItem = oForm.Items.Item("2");
            oItem = oForm.Items.Add("btnCopy", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Top = oNewItem.Top;
            oItem.Height = oNewItem.Height;
            oItem.Left = oNewItem.Left + oNewItem.Width + 5;
            oItem.Width = oNewItem.Width + 20;
            oButton = (SAPbouiCOM.Button)oItem.Specific;
            oButton.Caption = "Copy To";
            oItem.Visible = true;
            oItem.Enabled = false;
            oButton.ClickAfter += OButton_ClickAfter;
            //if (oItem.Enabled == true)
            //{
            //    oButton.ClickAfter += OButton_ClickAfter;
            //}
            //*******************************************************************
        }
        #endregion
        #region  OButton_ClickAfter()
        private static void OButton_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new NotImplementedException();
            try
            {
                oRS_Ref = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS1 = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oItem = (SAPbouiCOM.Item)oForm.Items.Item("btnCopy");
                #region Inventory Transfer Request Copy to Inventory transfer Request 
                if (oForm.TypeEx == "1250000940")
                {
                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    TransferRequest = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);
                    oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                    oCompany.XMLAsString = true;
                    string xmlStockTransfer = string.Empty;
                    int offset = oForm.DataSources.DBDataSources.Item(0).Offset;
                    string Entry = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", offset);
                    int DocEntry = 0;
                    DocEntry = int.Parse(Entry);
                    if (TransferRequest.GetByKey(DocEntry))
                    {
                        //Convert your strock transfer in a xml
                        xmlStockTransfer = TransferRequest.GetAsXML();
                    }
                    if (!string.IsNullOrEmpty(xmlStockTransfer))
                    {
                        //Intialize a new stock transfer through your xml
                        Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        TransferRequest = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObjectFromXML(xmlStockTransfer, 0);
                        //Change the fields that you want.
                        //TransferRequest.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_InventoryTransferRequest;
                        oRS_Ref.DoQuery($"SELECT \"DocEntry\",\"LineNum\",\"RefDocEntr\",\"RefDocNum\",\"RefObjType\" From \"WTQ21\" WHERE \"DocEntry\"={DocEntry}");
                        for (int i = 0; i < oRS_Ref.RecordCount; i++)
                        {
                            TransferRequest.DocumentReferences.Delete();
                            oRS_Ref.MoveNext();
                        }
                        TransferRequest.FromWarehouse = TransferRequest.UserFields.Fields.Item("U_tl_attn_ter").Value.ToString();
                        TransferRequest.ToWarehouse = TransferRequest.UserFields.Fields.Item("U_tl_attn_ter").Value.ToString();

                        TransferRequest.DueDate = DateTime.Today;
                        TransferRequest.DocDate = DateTime.Today;
                        TransferRequest.TaxDate = DateTime.Today;
                        #region Transfer.Line
                        //TransferRequest.PriceList = TransferRequest.PriceList;
                        for (int i = 0; i < TransferRequest.Lines.Count; i++)
                        {
                            TransferRequest.Lines.SetCurrentLine(i);
                            TransferRequest.Lines.FromWarehouseCode = TransferRequest.FromWarehouse;
                            TransferRequest.Lines.WarehouseCode = TransferRequest.ToWarehouse;
                            //TransferRequest.Lines.Add();
                        }
                        //Add the new transfer.
                        #endregion
                        if (TransferRequest.Add() != 0)
                        {
                            Application.SBO_Application.SetStatusBarMessage($"Failed to add Inventory Transfer Request:[{oCompany.GetLastErrorDescription()}]");
                        }
                        else
                        {
                            oCompany.GetNewObjectCode(out string Inventory);
                            if (TransferRequest.GetByKey(DocEntry))
                            {
                                oRS_Ref.DoQuery($"SELECT \"DocEntry\",\"LineNum\",\"RefDocEntr\",\"RefDocNum\",\"RefObjType\" From \"WTQ21\" WHERE \"DocEntry\"={DocEntry}");
                                for (int i = 0; i <= oRS_Ref.RecordCount; i++)
                                {
                                    TransferRequest.DocumentReferences.SetCurrentLine(i);
                                    if (i == oRS_Ref.RecordCount)
                                    {
                                        TransferRequest.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_InventoryTransferRequest;
                                        TransferRequest.DocumentReferences.ReferencedDocEntry = int.Parse(Inventory);
                                        TransferRequest.DocumentReferences.Add();
                                    }
                                    else
                                    {
                                        int RefObj = int.Parse(oRS_Ref.Fields.Item("RefObjType").Value.ToString());
                                        TransferRequest.DocumentReferences.ReferencedObjectType = (SAPbobsCOM.ReferencedObjectTypeEnum)RefObj;
                                        TransferRequest.DocumentReferences.ReferencedDocEntry = int.Parse(oRS_Ref.Fields.Item("RefDocEntr").Value.ToString());
                                        TransferRequest.DocumentReferences.Add();
                                        oRS_Ref.MoveNext();
                                    }
                                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                }
                                oRS1.DoQuery($"UPDATE \"OWTQ\" SET \"DocStatus\" = '{"C"}' WHERE \"DocEntry\" = {DocEntry}");
                                TransferRequest.Update();
                            }
                            Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_StockTransfersRequest, oCompany.GetNewObjectType(), Inventory);
                            Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                        }
                    }
                }
                #endregion
                #region Inventory Transfer Copy to Goods Issue Copy 
                if (oForm.TypeEx == "940")
                {
                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    GoodIssue = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    Transfer = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                    int offset = oForm.DataSources.DBDataSources.Item(0).Offset;
                    string Entry = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", offset);
                    int DocEntry = 0;
                    DocEntry = int.Parse(Entry);
                    SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oForm.Items.Item("7").Specific;
                    if (Transfer.GetByKey(DocEntry))
                    {
                        GoodIssue.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                        GoodIssue.DocDate = DateTime.Today;
                        GoodIssue.DocDueDate = DateTime.Today;
                        GoodIssue.TaxDate = DateTime.Today;
                        GoodIssue.BPL_IDAssignedToInvoice = Transfer.BPLID;
                        GoodIssue.UserFields.Fields.Item("U_tl_whsdesc").Value = Transfer.ToWarehouse;
                        string BinAbsEntry = null;
                        oRS.DoQuery($"SELECT A.\"ToBinCode\",B.\"AbsEntry\" FROM OWTR A " +
                            $"LEFT OUTER JOIN OBIN B ON A.\"ToBinCode\" = B.\"BinCode\" AND A.\"ToWhsCode\" = B.\"WhsCode\" WHERE A.\"DocEntry\" = {DocEntry} AND A.\"ToBinCode\" IS NOT NULL");
                        BinAbsEntry = oRS.Fields.Item("AbsEntry").Value.ToString();
                        for (int i = 0; i < Transfer.Lines.Count; i++)
                        {
                            //if (GoodIssue.Lines.LineStatus == 0)
                            //{
                            Transfer.Lines.SetCurrentLine(i);
                            //GoodReceipt.Lines.BaseEntry = DocEntry;
                            //GoodReceipt.Lines.BaseType = 60;
                            GoodIssue.Lines.BaseLine = i;
                            GoodIssue.Lines.UoMEntry = Transfer.Lines.UoMEntry;
                            GoodIssue.Lines.ItemCode = Transfer.Lines.ItemCode;
                            GoodIssue.Lines.ItemDescription = GoodIssue.Lines.ItemDescription;
                            GoodIssue.Lines.WarehouseCode = Transfer.ToWarehouse;
                            GoodIssue.Lines.Quantity = Transfer.Lines.Quantity;
                            //GoodReceipt.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO;
                            //GoodReceipt.Lines.BinAllocations.SetCurrentLine(i);
                            if (oRS.RecordCount == 1)
                            {
                                GoodIssue.Lines.BinAllocations.BinAbsEntry = int.Parse(BinAbsEntry);
                                //GoodReceipt.Lines.BinAllocations.BaseLineNumber = i;
                                GoodIssue.Lines.BinAllocations.Quantity = Transfer.Lines.Quantity;
                            }
                            GoodIssue.Lines.Add();
                            Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                        string Inventory = null;
                        //GoodIssue.Update();
                        int Add = GoodIssue.Add();
                        if (Add == 0)
                        {
                            oCompany.GetNewObjectCode(out Inventory);
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
                                        Transfer.DocumentReferences.ReferencedDocEntry = int.Parse(Inventory);
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
                                Application.SBO_Application.Forms.GetFormByTypeAndCount(940, pVal.FormTypeCount).Select();
                                Application.SBO_Application.ActivateMenuItem("1304");
                            }
                            Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsIssue, "", Inventory);
                            Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                        }
                        else
                        {
                            Application.SBO_Application.SetStatusBarMessage($"{oCompany.GetLastErrorDescription()}");
                        }
                    }
                    oItem.Enabled = false;
                }
                #endregion
                #region Goods Issue Copy To Goods Receipt 
                if (oForm.TypeEx == "720" && oItem.Enabled)
                {
                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    GoodIssue = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    GoodReceipt = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                    int offset = oForm.DataSources.DBDataSources.Item(0).Offset;
                    string Entry = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", offset);
                    int DocEntry = 0;
                    DocEntry = int.Parse(Entry);
                    SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oForm.Items.Item("7").Specific;
                    if (GoodIssue.GetByKey(DocEntry))
                    {
                        GoodReceipt.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                        GoodReceipt.DocDate = DateTime.Today;
                        GoodReceipt.DocDueDate = DateTime.Today;
                        GoodReceipt.TaxDate = DateTime.Today;
                        GoodReceipt.BPL_IDAssignedToInvoice = GoodIssue.BPL_IDAssignedToInvoice;
                        GoodReceipt.UserFields.Fields.Item("U_tl_whsdesc").Value = GoodIssue.UserFields.Fields.Item("U_tl_whsdesc").Value;
                        for (int i = 0; i < GoodIssue.Lines.Count; i++)
                        {
                            //if (GoodIssue.Lines.LineStatus == 0)
                            //{
                            GoodIssue.Lines.SetCurrentLine(i);
                            //GoodReceipt.Lines.BaseEntry = DocEntry;
                            //GoodReceipt.Lines.BaseType = 60;
                            GoodReceipt.Lines.BaseLine = i;
                            GoodReceipt.Lines.ItemCode = GoodIssue.Lines.ItemCode;
                            GoodReceipt.Lines.ItemDescription = GoodIssue.Lines.ItemDescription;
                            GoodReceipt.Lines.WarehouseCode = GoodIssue.Lines.WarehouseCode;
                            GoodReceipt.Lines.Quantity = GoodIssue.Lines.Quantity;
                            GoodReceipt.Lines.BinAllocations.BinAbsEntry = GoodIssue.Lines.BinAllocations.BinAbsEntry;
                            
                            GoodReceipt.Lines.BinAllocations.Quantity = GoodIssue.Lines.Quantity;
                            //GoodIssue.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close;
                            GoodReceipt.Lines.Add();
                            //}
                            //else
                            //{
                            //Application.SBO_Application.SetStatusBarMessage($"Line {i} was closed.");
                            ////System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvTransfer);
                            //}
                            Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                        string Inventory = null;
                        int Add = GoodReceipt.Add();
                        if (Add == 0)
                        {
                            oCompany.GetNewObjectCode(out Inventory);
                            //Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
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
                                        GoodIssue.DocumentReferences.ReferencedDocEntry = int.Parse(Inventory);
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
                                Application.SBO_Application.Forms.GetFormByTypeAndCount(720, pVal.FormTypeCount).Select();
                                Application.SBO_Application.ActivateMenuItem("1304");
                            }
                            Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsReceipt, oCompany.GetNewObjectType(), Inventory);
                            Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                        }
                        else
                        {
                            Application.SBO_Application.SetStatusBarMessage($"{oCompany.GetLastErrorDescription()}");
                        }
                    }
                }
                #endregion
                #region Pick Pack Type Transfer Request Copy To Goods Issue 
                if (oForm.TypeEx == "81" && oItem.Enabled == true)
                {
                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    SAPbobsCOM.PickLists oPickLists = (SAPbobsCOM.PickLists)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);
                    SAPbobsCOM.Documents GoodIssue = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    SAPbobsCOM.StockTransfer TransferRequest = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);

                    string data = null;
                    string Ref = ",";
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        SAPbouiCOM.CheckBox Check = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                        if (Check.Caption == "Y")
                        {
                            SAPbouiCOM.EditText Pack_No = (SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText DocNo = (SAPbouiCOM.EditText)oMatrix.Columns.Item("12").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText col_ItemCode = (SAPbouiCOM.EditText)oMatrix.Columns.Item("9").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText col_row = (SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(i).Specific;
                            data += ", " + Pack_No.Value;
                            int DocEntry = 0;
                            DocEntry = int.Parse(Pack_No.Value);
                            if (oPickLists.GetByKey(DocEntry))
                            {
                                oRS.DoQuery($"select \"DocEntry\" from \"OWTQ\" where \"DocNum\" = {DocNo.Value}");
                                int Entry = 0;
                                Entry = int.Parse(oRS.Fields.Item("DocEntry").Value.ToString());
                                if (TransferRequest.GetByKey(Entry))
                                {
                                    if (DocNo.Value == TransferRequest.DocNum.ToString())
                                    {
                                        int PickEntry = 0;
                                        PickEntry = int.Parse(col_row.Value.ToString()) - 1;

                                        //int row = 0;
                                        //row = int.Parse(oRS1.Fields.Item("LineNum").Value.ToString());
                                        TransferRequest.Lines.SetCurrentLine(PickEntry);
                                        GoodIssue.DocDate = DateTime.Today;
                                        GoodIssue.DocDueDate = DateTime.Today;
                                        GoodIssue.TaxDate = DateTime.Today;
                                        //GoodIssue.UserFields.Fields.Item("").Value=
                                        #region Add References
                                        //Add DocumentReferences********************************************************************
                                        //string[] RefArr = Ref.Substring(1).Split(',');
                                        //bool exists= Array.Exists(RefArr, elemen => elemen == oPickLists.Lines.OrderEntry.ToString());
                                        //if (!exists)
                                        //{
                                        //    //GoodIssue.DocumentReferences.SetCurrentLine(CheckCount);
                                        //    GoodIssue.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_InventoryTransferRequest;
                                        //    GoodIssue.DocumentReferences.ReferencedDocEntry = oPickLists.Lines.OrderEntry;
                                        //    GoodIssue.DocumentReferences.Add();
                                        //}
                                        //**********************************************************************************************
                                        #endregion
                                        GoodIssue.BPL_IDAssignedToInvoice = TransferRequest.BPLID;
                                        GoodIssue.Lines.ItemCode = TransferRequest.Lines.ItemCode;
                                        GoodIssue.Lines.WarehouseCode = TransferRequest.UserFields.Fields.Item("U_tl_attn_ter").Value.ToString();
                                        oRS1.DoQuery($"SELECT A.\"ItemCode\",A.\"WhsCode\",B.\"AbsEntry\",B.\"BinCode\" FROM \"OITW\" A " +
                                            $"LEFT OUTER JOIN \"OBIN\" B ON A.\"DftBinAbs\" = B.\"AbsEntry\" " +
                                            $"WHERE A.\"ItemCode\" = '{GoodIssue.Lines.ItemCode}' AND A.\"WhsCode\"='{TransferRequest.UserFields.Fields.Item("U_tl_attn_ter").Value.ToString()}'");
                                        GoodIssue.Lines.Quantity = TransferRequest.Lines.Quantity;
                                        GoodIssue.Lines.BinAllocations.BinAbsEntry = int.Parse(oRS1.Fields.Item("AbsEntry").Value.ToString());
                                        GoodIssue.Lines.BinAllocations.Quantity = TransferRequest.Lines.Quantity;
                                        //GoodIssue.Lines.BaseLine = TransferRequest.Lines.LineNum;
                                        oRS1.DoQuery($"UPDATE \"PKL1\" SET \"PickStatus\" = '{"C"}' WHERE \"AbsEntry\" = {DocEntry} AND \"OrderLine\" ={PickEntry}");
                                        GoodIssue.Lines.Add();
                                        Ref += "," + oPickLists.Lines.OrderEntry.ToString();
                                    }
                                }
                                Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                    }
                    #region GoodIssue.Add()
                    int Add = GoodIssue.Add();
                    if (Add == 0)
                    {
                        oCompany.GetNewObjectCode(out string Inventory);
                        string[] RefArr = Ref.Substring(1).Split(',');
                        int ITR_Entry = 0;
                        string En = ",";
                        foreach (string item in RefArr)
                        {
                            if (item != "")
                            {
                                string[] ArrEn = En.Substring(1).Split(',');
                                bool exists = Array.Exists(ArrEn, elemen => elemen == item);
                                if (!exists)
                                {
                                    ITR_Entry = int.Parse(item);
                                    if (TransferRequest.GetByKey(ITR_Entry))
                                    {
                                        oRS_Ref.DoQuery($"SELECT \"DocEntry\",\"LineNum\",\"RefDocEntr\",\"RefDocNum\",\"RefObjType\" From \"WTQ21\" WHERE \"DocEntry\"={ITR_Entry}");
                                        for (int i = 0; i <= oRS_Ref.RecordCount; i++)
                                        {
                                            TransferRequest.DocumentReferences.SetCurrentLine(i);
                                            if (i == oRS_Ref.RecordCount)
                                            {
                                                TransferRequest.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_GoodsIssue;
                                                TransferRequest.DocumentReferences.ReferencedDocEntry = int.Parse(Inventory);
                                                TransferRequest.DocumentReferences.Add();
                                            }
                                            else
                                            {
                                                int RefObj = int.Parse(oRS_Ref.Fields.Item("RefObjType").Value.ToString());
                                                TransferRequest.DocumentReferences.ReferencedObjectType = (SAPbobsCOM.ReferencedObjectTypeEnum)RefObj;
                                                TransferRequest.DocumentReferences.ReferencedDocEntry = int.Parse(oRS_Ref.Fields.Item("RefDocEntr").Value.ToString());
                                                TransferRequest.DocumentReferences.Add();
                                                oRS_Ref.MoveNext();
                                            }
                                        }
                                        TransferRequest.Update();
                                    }
                                    En += "," + item;
                                }
                            }
                            Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                        Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsIssue, "", Inventory);
                        Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
                    }
                    else
                    {
                        Application.SBO_Application.SetStatusBarMessage($"Failed to add Goods Issue: {oCompany.GetLastErrorDescription()}");
                    }
                    //System.Windows.Forms.MessageBox.Show($"Pack No: {data.Substring(1)}");
                    #endregion
                }
                #endregion
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage(er.Message);
            }
        }
        #endregion
        #region SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
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
        #endregion
    }
}
