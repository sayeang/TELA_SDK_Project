using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace Add_ITR
{
    class ITR_Process
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
                Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
                Application.SBO_Application.StatusBar.SetSystemMessage("Add-On Connecting....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning, "", "", "", 0);
                //oCompany = new SAPbobsCOM.Company();
                //oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                Application.SBO_Application.StatusBar.SetSystemMessage("Connecting Add-on successfully.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
                #region Connect_Company
                //oCompany.Server = "HDB@192.168.1.11:30013";
                //oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                //oCompany.CompanyDB = "TLTELA_DEVELOPER_TECH1";
                //oCompany.UserName = "manager";
                //oCompany.Password = "1234";
                //oCompany.DbUserName = "SYSTEM";
                //oCompany.DbPassword = "Biz@2022";

                //int iCon = oCompany.Connect();

                //try
                //{
                //    if (iCon == 0)
                //    {
                //        oRS = (SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                //        string QuerySelect = "SELECT * FROM \"OUSR\"";

                //        oRS.DoQuery(QuerySelect);

                //        Application.SBO_Application.MessageBox(oCompany.UserName.ToString());
                //    }
                //}
                //catch
                //{
                //    if (iCon != 0)
                //    {
                //        oCompany = (SAPbobsCOM.Company)(Application.SBO_Application.Company.GetDICompany());

                //        Application.SBO_Application.MessageBox(oCompany.UserName.ToString());
                //    }
                //}
                #endregion
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        #region Declaration
        private static SAPbobsCOM.Company oCompany = null;
        public static SAPbobsCOM.Recordset oRS_Ref, oRS = null;
        private static SAPbobsCOM.StockTransfer TransferRequest = null;
        private static SAPbouiCOM.Form oForm = null;
        #endregion
        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.FormTypeEx == "1250000940"
                && BusinessObjectInfo.Type == "1250000001" && BusinessObjectInfo.ActionSuccess)
            {
                Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                #region Connect_Company
                try
                {
                    oCompany = new SAPbobsCOM.Company();
                    oCompany.Server = "HDB@192.168.1.11:30013";
                    oCompany.SLDServer = "192.168.1.11:40001";
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                    oCompany.CompanyDB = "TLTELA_DEVELOPER_TECH1";
                    oCompany.UserName = "manager";
                    oCompany.Password = "1234";
                    oCompany.DbUserName = "SYSTEM";
                    oCompany.DbPassword = "Biz@2022";
                    oCompany.Connect();
                    int iCon = oCompany.Connect();
                    if (iCon!=0)
                    {
                        oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                        return;
                    }
                #endregion
                if (oCompany.UserName == "manager")
                {
                    #region Inventory Transfer Request Copy to Inventory transfer Request 
                    Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                    oRS_Ref = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                    if (!string.IsNullOrEmpty(xmlStockTransfer) && TransferRequest.UserFields.Fields.Item("U_tl_itr_type").Value.ToString()== "Store Request")
                    {
                        //Intialize a new stock transfer through your xml
                        Application.SBO_Application.StatusBar.SetSystemMessage("Processing...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        TransferRequest = (SAPbobsCOM.StockTransfer)oCompany.GetBusinessObjectFromXML(xmlStockTransfer, 0);
                        //Change the fields that you want.
                        oRS_Ref.DoQuery($"SELECT \"DocEntry\",\"LineNum\",\"RefDocEntr\",\"RefDocNum\",\"RefObjType\" From \"WTQ21\" WHERE \"DocEntry\"={DocEntry}");
                        for (int i = 0; i < oRS_Ref.RecordCount; i++)
                        {
                            TransferRequest.DocumentReferences.Delete();
                            oRS_Ref.MoveNext();
                        }
                        TransferRequest.DueDate = DateTime.Today;
                        TransferRequest.DocDate = DateTime.Today;
                        TransferRequest.TaxDate = DateTime.Today;
                        TransferRequest.FromWarehouse = TransferRequest.UserFields.Fields.Item("U_tl_attn_ter").Value.ToString();
                        TransferRequest.ToWarehouse = TransferRequest.UserFields.Fields.Item("U_tl_attn_ter").Value.ToString();
                        #region Transfer.Line
                        for (int i = 0; i < TransferRequest.Lines.Count; i++)
                        {
                            TransferRequest.Lines.SetCurrentLine(i);
                            TransferRequest.Lines.FromWarehouseCode = TransferRequest.FromWarehouse;
                            TransferRequest.Lines.WarehouseCode = TransferRequest.ToWarehouse;
                        }
                        //Add the new transfer.
                        #endregion
                        string Inventory = null;
                        if (TransferRequest.Add() != 0)
                        {
                            Application.SBO_Application.SetStatusBarMessage($"{oCompany.GetLastErrorDescription()}");
                            return;
                        }
                        else
                        {
                            oCompany.GetNewObjectCode(out Inventory);
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
                                int update = TransferRequest.Update();
                            }
                            try
                            {
                                Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_StockTransfersRequest, "", Inventory);
                            }
                            catch (Exception er)
                            {
                                Application.SBO_Application.SetStatusBarMessage(er.Message);
                            }
                            Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "");
                        }
                    }
                    else
                    {
                        return;
                    }
                    #endregion
                }
                else
                {
                    Application.SBO_Application.SetStatusBarMessage($"The user [{oCompany.UserName}] cannot create new document.", SAPbouiCOM.BoMessageTime.bmt_Short);
                    return;
                }
                    oCompany.Disconnect();
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.MessageBox(ex.Message);
                    return;
                }
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
