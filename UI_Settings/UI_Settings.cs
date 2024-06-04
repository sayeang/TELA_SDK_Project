using SAPbouiCOM.Framework;
using System;

namespace CFL
{
    class UI_Settings
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
                Application.SBO_Application.MenuEvent += SBO_Application_MenuEvent;
                Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
                Application.SBO_Application.StatusBar.SetSystemMessage("Add-On Connecting....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning, "", "", "", 0);
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                Application.SBO_Application.StatusBar.SetSystemMessage("Add-On Connected Successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return;
            }
        }

        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            if ((BusinessObjectInfo.FormTypeEx == "720" || BusinessObjectInfo.FormTypeEx == "721") && BusinessObjectInfo.ActionSuccess && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
            {
                txtCardCode_1 = (SAPbouiCOM.EditText)oForm.Items.Item("U_tl_cardcode").Specific;
                txtCardCode_1.Item.Enabled = false;
            }
        }

        private static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if (pVal.MenuUID == "1282" && !pVal.BeforeAction && !pVal.InnerEvent)
                {
                    if (oForm.Type == 720 || oForm.Type == 721 || oForm.Type == 1250000940 || oForm.Type == 940)
                    {
                        //oRS_User = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRS_User.DoQuery($"SELECT \"U_tl_user_pro\" FROM OUSR WHERE \"USERID\" = {oCompany.UserSignature}");
                        TR_Auto_Creation = oRS_User.Fields.Item("U_tl_user_pro").Value.ToString();
                        if (TR_Auto_Creation == "Store")
                        {
                            oForm.Freeze(true);
                            #region 'ITR && ITF'
                            if ((oForm.Type == 1250000940 || oForm.Type == 940))
                            {
                                try
                                {
                                    txtCardCode = (SAPbouiCOM.EditText)oForm.Items.Item("3").Specific;
                                    combo_Series = (SAPbouiCOM.ComboBox)oForm.Items.Item("40").Specific;
                                    //oRS_Series = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    oRS_Series.DoQuery($"SELECT A.\"ObjectCode\",A.\"Series\",A.\"SeriesName\",A.\"BPLId\",B.\"DflCust\" FROM NNM1 A " +
                                                       $"LEFT OUTER JOIN \"OBPL\" B ON A.\"BPLId\"=B.\"BPLId\" WHERE A.\"Series\"='{combo_Series.Value}'");
                                    oForm.Freeze(true);
                                    txtCardCode.Item.Enabled = true;
                                    txtCardCode.Value = null;
                                    txtCardCode.Value = oRS_Series.Fields.Item("DflCust").Value.ToString();
                                    //oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                    //oForm.Items.Item("U_tl_attn_ter").Click();
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("23").Specific;
                                    col_ItemCode = (SAPbouiCOM.Column)oMatrix.Columns.Item("1");
                                    col_ItemCode.Cells.Item(oMatrix.VisualRowCount).Click();
                                    txtCardCode.Item.Enabled = false;
                                    oForm.Items.Item("7").Enabled = false;
                                    oForm.Freeze(false);
                                }
                                catch (Exception er)
                                {
                                    Application.SBO_Application.MessageBox(er.Message);
                                    return;
                                }
                            }
                            else
                            {
                                //txtCardCode.Item.Enabled = false;
                                //oForm.Items.Item("7").Enabled = false;
                            }
                            #endregion
                            #region 'GI & GR'
                            if ((oForm.Type == 720 || oForm.Type == 721))
                            {
                                try
                                {
                                    txtCardCode_1 = (SAPbouiCOM.EditText)oForm.Items.Item("U_tl_cardcode").Specific;
                                    combo_Series_1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("30").Specific;
                                    oRS_Series = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    oRS_Series.DoQuery($"SELECT A.\"ObjectCode\",A.\"Series\",A.\"SeriesName\",A.\"BPLId\",B.\"DflCust\" FROM NNM1 A " +
                                                       $"LEFT OUTER JOIN \"OBPL\" B ON A.\"BPLId\"=B.\"BPLId\" WHERE A.\"Series\"='{combo_Series_1.Value}'");
                                    oForm.Freeze(true);
                                    txtCardCode_1.Item.Enabled = true;
                                    txtCardCode_1.Value = null;
                                    txtCardCode_1.Value = oRS_Series.Fields.Item("DflCust").Value.ToString();
                                    //oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                    //oForm.Items.Item("U_tl_gitstore").Click();
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;
                                    col_ItemCode_1 = (SAPbouiCOM.Column)oMatrix.Columns.Item("1");
                                    col_ItemCode_1.Cells.Item(oMatrix.VisualRowCount).Click();
                                    txtCardCode_1.Item.Enabled = false;
                                    oForm.Freeze(false);
                                }
                                catch (Exception er)
                                {
                                    Application.SBO_Application.MessageBox(er.Message);
                                    return;
                                }
                            }
                            else
                            {
                                //txtCardCode_1.Item.Enabled = false;
                            }
                            #endregion
                            oForm.Freeze(false);
                        }
                    }
                    else
                    {
                        return;
                    }
                }
            }
            catch (Exception er)
            {
                System.Windows.Forms.MessageBox.Show($"{er.Message}");
                return;
            }


        }
        #region 'Declare object'
        private static SAPbouiCOM.Matrix oMatrix = null;
        private static SAPbouiCOM.Column col_ItemCode, col_ItemName, col_ItemCode_1, col_ItemName_1 = null;
        private static SAPbouiCOM.EditText txtCardCode, txtCardCode_1, txtItemCode, txtItemName, txtItemCode_1, txtItemName_1, Atten_teminal = null;
        private static SAPbouiCOM.ComboBox combo_Series, combo_Series_1 = null;
        private static SAPbouiCOM.ChooseFromList oCFL = null;
        private static SAPbouiCOM.Form oForm = null;
        private static SAPbouiCOM.Item oItem = null;
        public static SAPbouiCOM.Conditions oCons = null;
        public static SAPbouiCOM.Condition oCon = null;
        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbobsCOM.Recordset oRS_Cfl, oRS_User, oRS_Series = null;
        private static SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
        private static int FormID = 0;
        private static string TR_Auto_Creation, FormUI = null;
        public static bool menuID = false;

        #endregion

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {

                if ((pVal.FormType == 720 || pVal.FormType == 721 || pVal.FormType == 1250000940 || pVal.FormType == 940) && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    FormID = pVal.FormType;
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD || pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT))
                    {
                        oRS_User = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRS_User.DoQuery($"SELECT \"U_tl_user_pro\" FROM OUSR WHERE \"USERID\" = {oCompany.UserSignature}");
                        TR_Auto_Creation = oRS_User.Fields.Item("U_tl_user_pro").Value.ToString();
                        if (pVal.FormType == 1250000940 || pVal.FormType == 940)
                        {
                            if (pVal.Action_Success && TR_Auto_Creation == "Store")
                            {
                                txtCardCode = (SAPbouiCOM.EditText)oForm.Items.Item("3").Specific;
                                if (txtCardCode.Value == "" || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT &&
                                    ((pVal.FormType == 1250000940 && pVal.ItemUID == "1250000068") || (pVal.FormType == 940 && pVal.ItemUID == "40"))))
                                {
                                    try
                                    {
                                        oRS_Series = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        if (pVal.FormType == 940)
                                        {
                                          combo_Series = (SAPbouiCOM.ComboBox)oForm.Items.Item("40").Specific;
                                        }
                                        if (pVal.FormType == 1250000940)
                                        {
                                            combo_Series = (SAPbouiCOM.ComboBox)oForm.Items.Item("1250000068").Specific;
                                        }
                                        oRS_Series.DoQuery($"SELECT A.\"ObjectCode\",A.\"Series\",A.\"SeriesName\",A.\"BPLId\",B.\"DflCust\" FROM NNM1 A " +
                                                           $"LEFT OUTER JOIN \"OBPL\" B ON A.\"BPLId\"=B.\"BPLId\" WHERE A.\"Series\"='{combo_Series.Value}'");
                                        oForm.Freeze(true);
                                        txtCardCode.Item.Enabled = true;
                                        txtCardCode.Value = null;
                                        txtCardCode.Value = oRS_Series.Fields.Item("DflCust").Value.ToString();
                                        oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                        //oForm.Items.Item("U_tl_attn_ter").Click();
                                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("23").Specific;
                                        col_ItemCode = (SAPbouiCOM.Column)oMatrix.Columns.Item("1");
                                        col_ItemCode.Cells.Item(oMatrix.VisualRowCount).Click();
                                        txtCardCode.Item.Enabled = false;
                                        oForm.Items.Item("7").Enabled = false;
                                        oForm.Freeze(false);
                                    }
                                    catch (Exception er)
                                    {
                                        Application.SBO_Application.MessageBox(er.Message);
                                        return;
                                    }
                                }
                                else
                                {
                                    txtCardCode.Item.Enabled = false;
                                    //oForm.Items.Item("7").Enabled = false;
                                }

                            }
                           
                            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action)
                            {
                                Add_Attenterminal();
                                oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("23").Specific;
                                col_ItemCode = (SAPbouiCOM.Column)oMatrix.Columns.Item("1");
                                col_ItemName = (SAPbouiCOM.Column)oMatrix.Columns.Item("2");
                                col_ItemCode.ChooseFromListBefore += OColumn1_ChooseFromListBefore;
                                col_ItemName.ChooseFromListBefore += OColumn2_ChooseFromListBefore;
                            }
                        }
                        if (pVal.FormType == 720 || pVal.FormType == 721)
                        {
                            if (pVal.Action_Success && TR_Auto_Creation == "Store")
                            {
                                txtCardCode_1 = (SAPbouiCOM.EditText)oForm.Items.Item("U_tl_cardcode").Specific;
                                if (txtCardCode_1.Value == "" || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && ((pVal.FormType == 720 || pVal.FormType == 721) && pVal.ItemUID == "30")))
                                {
                                    try
                                    {
                                        combo_Series_1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("30").Specific;
                                        oRS_Series = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        oRS_Series.DoQuery($"SELECT A.\"ObjectCode\",A.\"Series\",A.\"SeriesName\",A.\"BPLId\",B.\"DflCust\" FROM NNM1 A " +
                                                           $"LEFT OUTER JOIN \"OBPL\" B ON A.\"BPLId\"=B.\"BPLId\" WHERE A.\"Series\"='{combo_Series_1.Value}'");
                                        oForm.Freeze(true);
                                        txtCardCode_1.Item.Enabled = true;
                                        txtCardCode_1.Value = null;
                                        txtCardCode_1.Value = oRS_Series.Fields.Item("DflCust").Value.ToString();
                                        oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                        //oForm.Items.Item("U_tl_gitstore").Click();
                                        oForm.Items.Item("21").Click();
                                        //oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;
                                        //col_ItemCode_1 = (SAPbouiCOM.Column)oMatrix.Columns.Item("1");
                                        //col_ItemCode_1.Cells.Item(oMatrix.VisualRowCount).Click();
                                        txtCardCode_1.Item.Enabled = false;
                                        oForm.Freeze(false);
                                    }
                                    catch (Exception er)
                                    {
                                        Application.SBO_Application.MessageBox(er.Message);
                                        return;
                                    }
                                }
                                else
                                {
                                    txtCardCode_1.Item.Enabled = false;
                                }
                            }
                            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action)
                            {
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;
                                col_ItemCode_1 = (SAPbouiCOM.Column)oMatrix.Columns.Item("1");
                                col_ItemName_1 = (SAPbouiCOM.Column)oMatrix.Columns.Item("2");
                                col_ItemCode_1.ChooseFromListBefore += Col_ItemCode_1_ChooseFromListBefore;
                                col_ItemName_1.ChooseFromListBefore += Col_ItemName_1_ChooseFromListBefore;
                            }

                        }
                    }
                }

                if (pVal.FormTypeEx == "0" && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD && !pVal.Before_Action && !pVal.InnerEvent)
                {
                    try
                    {
                        oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    }
                    catch (Exception)
                    {
                        //Application.SBO_Application.SetStatusBarMessage(er.Message);
                        return;
                    }
                    oForm.Freeze(false);
                    if (oForm.Visible == true)
                    {
                        oForm.Freeze(false);
                    }
                    oForm.Refresh();
                    oForm.Update();
                }
                if ((pVal.FormType == 10063) && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD && pVal.Before_Action && !pVal.InnerEvent)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    if (pVal.ItemUID == "2")
                    {
                        oForm.Close();
                    }
                    //if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.InnerEvent == true)
                    //{
                    //    SAPbouiCOM.Button btnNew = (SAPbouiCOM.Button)oForm.Items.Item("5").Specific;
                    //    btnNew.Item.Visible = false;
                    //}

                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage(er.Message);
            }
        }

        private static void Col_ItemName_1_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                #region ItemName

                if (TR_Auto_Creation == "Store")
                {
                    //txtItemName_1 = (SAPbouiCOM.EditText)col_ItemName_1.Cells.Item(pVal.Row).Specific;
                    oRS_Cfl = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //txtCardCode_1 = (SAPbouiCOM.EditText)oForm.Items.Item("U_tl_cardcode").Specific;
                    oRS_Cfl.DoQuery($"SELECT A.\"ItemCode\", A.\"CardCode\" FROM \"OSCN\" A LEFT OUTER JOIN \"OITM\" B ON A.\"ItemCode\" = B.\"ItemCode\" WHERE A.\"CardCode\"='{txtCardCode_1.Value}'");
                    oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(pVal.FormUID);
                    oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item(col_ItemName_1.ChooseFromListUID);
                    SAPbouiCOM.ConditionsClass emptyCon1;
                    emptyCon1 = new SAPbouiCOM.ConditionsClass();
                    oCFL.SetConditions(emptyCon1);
                    oCons = oCFL.GetConditions();
                    for (int i = 0; i < oRS_Cfl.RecordCount; i++)
                    {
                        oCon = oCons.Add();
                        oCon.Alias = "ItemCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oRS_Cfl.Fields.Item("ItemCode").Value.ToString();
                        if (i + 1 != oRS_Cfl.RecordCount)
                        {
                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                        }
                        oRS_Cfl.MoveNext();
                    }
                    oCFL.SetConditions(oCons);
                }
                #endregion
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage($"{er.Message}");
            }
        }

        private static void Col_ItemCode_1_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                #region ItemCode

                if (TR_Auto_Creation == "Store")
                {
                    //txtItemCode_1 = (SAPbouiCOM.EditText)col_ItemCode_1.Cells.Item(pVal.Row).Specific;
                    oRS_Cfl = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //txtCardCode_1 = (SAPbouiCOM.EditText)oForm.Items.Item("U_tl_cardcode").Specific;
                    oRS_Cfl.DoQuery($"SELECT A.\"ItemCode\", A.\"CardCode\" FROM \"OSCN\" A LEFT OUTER JOIN \"OITM\" B ON A.\"ItemCode\" = B.\"ItemCode\" WHERE A.\"CardCode\"='{txtCardCode_1.Value}'");
                    oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(pVal.FormUID);
                    oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item(col_ItemCode_1.ChooseFromListUID);
                    SAPbouiCOM.ConditionsClass emptyCon1;
                    emptyCon1 = new SAPbouiCOM.ConditionsClass();
                    oCFL.SetConditions(emptyCon1);
                    oCons = oCFL.GetConditions();
                    for (int i = 0; i < oRS_Cfl.RecordCount; i++)
                    {
                        oCon = oCons.Add();
                        oCon.Alias = "ItemCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oRS_Cfl.Fields.Item("ItemCode").Value.ToString();
                        if (i + 1 != oRS_Cfl.RecordCount)
                        {
                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                        }
                        oRS_Cfl.MoveNext();
                    }
                    oCFL.SetConditions(oCons);
                }
                #endregion
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage($"{er.Message}");
            }
        }

        private static void OColumn2_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                #region ItemName

                if (TR_Auto_Creation == "Store")
                {
                    //txtItemName = (SAPbouiCOM.EditText)col_ItemCode.Cells.Item(pVal.Row).Specific;
                    oRS_Cfl = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //txtCardCode = (SAPbouiCOM.EditText)oForm.Items.Item("3").Specific;
                    oRS_Cfl.DoQuery($"SELECT A.\"ItemCode\", A.\"CardCode\" FROM \"OSCN\" A LEFT OUTER JOIN \"OITM\" B ON A.\"ItemCode\" = B.\"ItemCode\" WHERE A.\"CardCode\"='{txtCardCode.Value}'");
                    oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(pVal.FormUID);
                    oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item(col_ItemName.ChooseFromListUID);
                    SAPbouiCOM.ConditionsClass emptyCon1;
                    emptyCon1 = new SAPbouiCOM.ConditionsClass();
                    oCFL.SetConditions(emptyCon1);
                    oCons = oCFL.GetConditions();
                    for (int i = 0; i < oRS_Cfl.RecordCount; i++)
                    {
                        oCon = oCons.Add();
                        oCon.Alias = "ItemCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oRS_Cfl.Fields.Item("ItemCode").Value.ToString();
                        if (i + 1 != oRS_Cfl.RecordCount)
                        {
                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                        }
                        oRS_Cfl.MoveNext();
                    }
                    oCFL.SetConditions(oCons);
                }
                #endregion
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage($"{er.Message}");
            }
        }

        private static void OColumn1_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                #region ItemCode
                oRS_User = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS_User.DoQuery($"SELECT \"U_tl_user_pro\" FROM OUSR WHERE \"USERID\" = {oCompany.UserSignature}");
                var TR_Auto_Creation = oRS_User.Fields.Item("U_tl_user_pro").Value.ToString();
                if (TR_Auto_Creation == "Store")
                {
                    //txtItemCode = (SAPbouiCOM.EditText)col_ItemCode.Cells.Item(pVal.Row).Specific;
                    oRS_Cfl = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //txtCardCode = (SAPbouiCOM.EditText)oForm.Items.Item("3").Specific;
                    oRS_Cfl.DoQuery($"SELECT A.\"ItemCode\", A.\"CardCode\" FROM \"OSCN\" A LEFT OUTER JOIN \"OITM\" B ON A.\"ItemCode\" = B.\"ItemCode\" WHERE A.\"CardCode\"='{txtCardCode.Value}'");
                    oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(pVal.FormUID);
                    oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item(col_ItemCode.ChooseFromListUID);
                    SAPbouiCOM.ConditionsClass emptyCon1;
                    emptyCon1 = new SAPbouiCOM.ConditionsClass();
                    oCFL.SetConditions(emptyCon1);
                    oCons = oCFL.GetConditions();
                    for (int i = 0; i < oRS_Cfl.RecordCount; i++)
                    {
                        oCon = oCons.Add();
                        oCon.Alias = "ItemCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oRS_Cfl.Fields.Item("ItemCode").Value.ToString();
                        if (i + 1 != oRS_Cfl.RecordCount)
                        {
                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                        }
                        oRS_Cfl.MoveNext();
                    }
                    oCFL.SetConditions(oCons);
                }
                #endregion
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage($"{er.Message}");
            }
        }

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
                if (oForm.TypeEx == "1250000940")
                {
                    Atten_teminal.DataBind.SetBound(true, "OWTQ", "U_tl_attn_ter");
                }
                if (oForm.TypeEx == "940")
                {
                    Atten_teminal.DataBind.SetBound(true, "OWTR", "U_tl_attn_ter");
                }

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

        #region Atten_teminal_ChooseFromListAfter
        private static void Atten_teminal_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new NotImplementedException();
            try
            {
                Atten_teminal = (SAPbouiCOM.EditText)oForm.Items.Item("Attn").Specific;
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
                oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetForm(FormID.ToString(), pVal.FormTypeCount);
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

        public static void CFL_ItemCode()
        {
            oRS_User = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRS_User.DoQuery($"SELECT \"U_tl_user_pro\" FROM OUSR WHERE \"USERID\" = {oCompany.UserSignature}");
            var TR_Auto_Creation = oRS_User.Fields.Item("U_tl_user_pro").Value.ToString();
            if (TR_Auto_Creation == "Store")
            {
                oRS_Cfl = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                txtCardCode = (SAPbouiCOM.EditText)oForm.Items.Item("3").Specific;
                oRS_Cfl.DoQuery($"SELECT A.\"ItemCode\", A.\"CardCode\" FROM \"OSCN\" A LEFT OUTER JOIN \"OITM\" B ON A.\"ItemCode\" = B.\"ItemCode\" WHERE A.\"CardCode\"='{txtCardCode.Value}'");
                if (txtItemCode.Active)
                {
                    oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item(col_ItemCode.ChooseFromListUID);
                }
                if (txtItemName.Active)
                {
                    oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item(col_ItemName.ChooseFromListUID);
                }
                SAPbouiCOM.ConditionsClass emptyCon1;
                emptyCon1 = new SAPbouiCOM.ConditionsClass();
                oCFL.SetConditions(emptyCon1);
                oCons = oCFL.GetConditions();
                for (int i = 0; i < oRS_Cfl.RecordCount; i++)
                {
                    oCon = oCons.Add();
                    oCon.Alias = "ItemCode";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = oRS_Cfl.Fields.Item("ItemCode").Value.ToString();
                    if (i + 1 != oRS_Cfl.RecordCount)
                    {
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                    }
                    oRS_Cfl.MoveNext();
                }
                oCFL.SetConditions(oCons);
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
