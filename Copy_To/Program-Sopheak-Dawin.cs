using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace Copy_To
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
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
                Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            //Application.SBO_Application.SetStatusBarMessage($"Code :{BusinessObjectInfo.ObjectKey}");
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if (((pVal.FormType == 720 || pVal.FormType==1250000940 || pVal.FormType==81) && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) && (pVal.Before_Action == true))
                {
                    oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) && (pVal.Before_Action == true))
                    {
                        //sForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                        AddBtn();
                    }
                    if (pVal.FormType == 81)
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("19").Specific;
                        oItem = (SAPbouiCOM.Item)oForm.Items.Item("btnCopy");
                        SAPbouiCOM.Item Picked = (SAPbouiCOM.Item)oForm.Items.Item("8");
                        SAPbouiCOM.StaticText BtnPick = (SAPbouiCOM.StaticText)oForm.Items.Item("8").Specific;
                        //Application.SBO_Application.StatusBar.SetText($"{oForm.PaneLevel}");
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
                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
            }
        }

        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbouiCOM.Form oForm,sForm = null;
        public static SAPbouiCOM.Item oItem, oNewItem = null;
        public static SAPbouiCOM.Button oButton = null;
        public static SAPbouiCOM.Matrix oMatrix = null;
        public static SAPbobsCOM.Recordset oRS_Ref = null;
        public static void AddBtn()
        {
            oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            oRS_Ref= (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oNewItem = oForm.Items.Item("2");
            //SAPbouiCOM.Item oNewItem1 = oForm.Items.Item("13");
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

        private static void OButton_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new NotImplementedException();
            try
            {
                oItem = (SAPbouiCOM.Item)oForm.Items.Item("btnCopy");
                //Goods Issue Copy To Goods Receipt 
                if (oForm.TypeEx=="720")
                {
                    SAPbobsCOM.Documents GoodIssue = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    SAPbobsCOM.Documents GoodReceipt = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                    int offset = oForm.DataSources.DBDataSources.Item(0).Offset;
                    string Entry = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", offset);
                    int DocEntry = 0;
                    DocEntry = int.Parse(Entry);
                    SAPbouiCOM.EditText DocNum = (SAPbouiCOM.EditText)oForm.Items.Item("7").Specific;
                    if (GoodIssue.GetByKey(DocEntry))
                    {
                        //GoodReceipt.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_GoodsIssue;
                        //GoodReceipt.DocumentReferences.ReferencedDocEntry = DocEntry;
                        //GoodReceipt.DocumentReferences.Add();
                        GoodReceipt.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                        GoodReceipt.DocDate = DateTime.Today;
                        GoodReceipt.DocDueDate = DateTime.Today;
                        GoodReceipt.TaxDate = DateTime.Today;
                        GoodReceipt.BPL_IDAssignedToInvoice = GoodIssue.BPL_IDAssignedToInvoice;
                        GoodReceipt.UserFields.Fields.Item("U_tl_whsdesc").Value = "WH11";
                        
                        for (int i = 0; i < GoodIssue.Lines.Count; i++)
                        {
                            //if (GoodIssue.Lines.LineStatus == 0)
                            //{
                                GoodIssue.Lines.SetCurrentLine(i);
                                //GoodReceipt.Lines.BaseEntry = DocEntry;
                                //GoodReceipt.Lines.BaseType = 60;
                                GoodReceipt.Lines.BaseLine = i;
                                GoodReceipt.Lines.UoMEntry = GoodIssue.Lines.UoMEntry;
                                GoodReceipt.Lines.ItemCode = GoodIssue.Lines.ItemCode;
                                GoodReceipt.Lines.ItemDescription = GoodIssue.Lines.ItemDescription;
                                GoodReceipt.Lines.WarehouseCode = "WH03";
                                GoodReceipt.Lines.Quantity = GoodIssue.Lines.Quantity;
                                //GoodReceipt.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO;
                                //GoodReceipt.Lines.BinAllocations.SetCurrentLine(i);
                                GoodReceipt.Lines.BinAllocations.BinAbsEntry = 102;
                                //GoodReceipt.Lines.BinAllocations.BaseLineNumber = i;
                                GoodReceipt.Lines.BinAllocations.Quantity = GoodIssue.Lines.Quantity;
                                //GoodIssue.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close;
                                GoodReceipt.Lines.Add();
                            //}
                            //else
                            //{
                            //Application.SBO_Application.SetStatusBarMessage($"Line {i} was closed.");
                            ////System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvTransfer);
                            //}
                            Application.SBO_Application.StatusBar.SetSystemMessage("Operating...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                        string Inventory = null;
                        //GoodIssue.Update();
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
                                }
                                GoodIssue.Update();
                                Application.SBO_Application.Forms.GetFormByTypeAndCount(720, pVal.FormTypeCount).Select();
                                Application.SBO_Application.ActivateMenuItem("1304");
                            }
                            Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsReceipt, oCompany.GetNewObjectType(),Inventory);
                            Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                        }
                        else
                        {
                            Application.SBO_Application.SetStatusBarMessage($"{oCompany.GetLastErrorDescription()}");
                        }
                    }
                }
                //Inventory Transfer Request Copy to Inventory transfer Request 
                if (oForm.TypeEx=="1250000940")
                {
                    SAPbobsCOM.StockTransfer TransferRequest = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);
                    oCompany.XmlExportType =SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
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
                        TransferRequest = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObjectFromXML(xmlStockTransfer, 0);
                        //Change the fields that you want.
                        TransferRequest.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_InventoryTransferRequest;
                        TransferRequest.DocumentReferences.ReferencedDocEntry = DocEntry;
                        TransferRequest.DocumentReferences.Add();
                        TransferRequest.DueDate = DateTime.Today;
                        TransferRequest.DocDate = DateTime.Today;
                        TransferRequest.TaxDate = DateTime.Today;
                        TransferRequest.PriceList = TransferRequest.PriceList;
                        for (int i = 0; i < TransferRequest.Lines.Count; i++)
                        {
                            TransferRequest.Lines.SetCurrentLine(i);
                            TransferRequest.Lines.FromWarehouseCode = TransferRequest.Lines.FromWarehouseCode;
                            TransferRequest.Lines.WarehouseCode = TransferRequest.Lines.WarehouseCode;
                        }
                        //Add the new transfer.
                        if (TransferRequest.Add() != 0)
                        {
                            //oCompany.GetNewObjectCode(out Inventory);
                            Application.SBO_Application.SetStatusBarMessage($"Failed to add Inventory Transfer Request:[{oCompany.GetLastErrorDescription()}]");
                        }
                        else
                        {

                            Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success,"","");
                            Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_StockTransfersRequest, oCompany.GetNewObjectType(), oCompany.GetNewObjectKey());

                        }
                    }
                }
                //Pick Pack Type Transfer Request Copy To Goods Issue 
                if (oForm.TypeEx=="81" && oItem.Enabled==true)
                {
                    SAPbobsCOM.PickLists oPickLists = (SAPbobsCOM.PickLists)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);
                    SAPbobsCOM.Documents GoodIssue = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    SAPbobsCOM.StockTransfer TransferRequest = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);
                    SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oRS1 = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string data = null;
                    string Ref =",";
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        SAPbouiCOM.CheckBox Check = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                        if (Check.Caption=="Y")
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
                                Entry=int.Parse(oRS.Fields.Item("DocEntry").Value.ToString());
                                    if (TransferRequest.GetByKey(Entry))
                                    {
                                    if (DocNo.Value == TransferRequest.DocNum.ToString())
                                    {
                                        int PickEntry = int.Parse(col_row.Value.ToString());
                                        oRS1.DoQuery($"SELECT B.\"AbsEntry\",A.\"PickStatus\",A.\"DocEntry\",A.\"ItemCode\",A.\"LineNum\",B.\"PickEntry\"" +
                                                    $"FROM \"WTQ1\" A RIGHT OUTER JOIN \"PKL1\" B ON  A.\"LineNum\" = B.\"OrderLine\" and A.\"DocEntry\" = B.\"OrderEntry\"" +
                                                    $" WHERE B.\"AbsEntry\"={DocEntry} AND A.\"DocEntry\"={Entry} And A.\"ItemCode\"='{col_ItemCode.Value}' AND B.\"PickEntry\"={PickEntry}" +
                                                    $"");
                                        int row = 0;
                                        row = int.Parse(oRS1.Fields.Item("LineNum").Value.ToString());
                                        TransferRequest.Lines.SetCurrentLine(row);
                                        GoodIssue.DocDate = DateTime.Today;
                                        GoodIssue.DocDueDate = DateTime.Today;
                                        GoodIssue.TaxDate = DateTime.Today;
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
                                        GoodIssue.BPL_IDAssignedToInvoice = TransferRequest.BPLID;
                                        GoodIssue.Lines.WarehouseCode = "WS11";
                                        GoodIssue.Lines.ItemCode = TransferRequest.Lines.ItemCode;
                                        //GoodIssue.Lines.BaseLine = oPickLists.Lines.LineNumber;
                                        GoodIssue.Lines.Add();
                                        Ref += "," + oPickLists.Lines.OrderEntry.ToString();
                                    }
                                }
                            }
                        }
                        Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning); 
                    }
                    string Inventory = null;
                    int Add = GoodIssue.Add();
                    if (Add == 0)
                    {
                        oCompany.GetNewObjectCode(out Inventory);
                        
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
                            Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                        Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
                        Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_GoodsIssue, "",Inventory);
                    }
                    else
                    {
                        Application.SBO_Application.SetStatusBarMessage($"Failed to add Goods Issue: {oCompany.GetLastErrorDescription()}");
                    }
                    //System.Windows.Forms.MessageBox.Show($"Pack No: {data.Substring(1)}");
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
