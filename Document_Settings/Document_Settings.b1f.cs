using SAPbouiCOM.Framework;
using System;

namespace SBOAddonProject_Setting
{
    [FormAttribute("SBOAddonProject_Setting.Document_Settings_b1f", "Document_Settings.b1f")]
    class Document_Settings_b1f : UserFormBase
    {
        private SAPbouiCOM.Button btnAdd, BtnCancel;
        private SAPbouiCOM.CheckBox CheckBox1, CheckBox2;
        private SAPbouiCOM.DataTable oDataTable, oDataTable0, oDataTable_0, oDataTable1, oDataTable_1, oDataTable2, oDataTable_2, oDataTable3, oDataTable_3, oDatatableIT, oDataTableIT_0;
        public SAPbouiCOM.Matrix Matrix0, Matrix1, Matrix2;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Folder Folder0, Folder1;
        private SAPbouiCOM.ChooseFromList oCFL, oCFLIT;
        private SAPbouiCOM.Conditions oCFLConditions, IoCFLConditions;
        private SAPbouiCOM.Condition oCFLCondition, IoCFLCondition;
        private SAPbouiCOM.ISBOChooseFromListEventArg CFLValue;
        private SAPbouiCOM.Menus oMenus = null;
        private SAPbouiCOM.MenuItem oMenuItem = null;
        private SAPbouiCOM.MenuCreationParams oCreationPackage1 = null;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.btnAdd = ((SAPbouiCOM.Button)(this.GetItem("btnAdd").Specific));
            this.btnAdd.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnAdd_ClickBefore);
            this.BtnCancel = ((SAPbouiCOM.Button)(this.GetItem("btnCancel").Specific));
            this.BtnCancel.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.CheckBox1 = ((SAPbouiCOM.CheckBox)(this.GetItem("check1").Specific));
            this.CheckBox1.ClickAfter += new SAPbouiCOM._ICheckBoxEvents_ClickAfterEventHandler(this.CheckBox1_ClickAfter);
            this.CheckBox2 = ((SAPbouiCOM.CheckBox)(this.GetItem("check2").Specific));
            this.CheckBox2.ClickAfter += new SAPbouiCOM._ICheckBoxEvents_ClickAfterEventHandler(this.CheckBox2_ClickAfter);
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_0").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("oMatrix").Specific));
            this.Matrix0.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix0_ClickAfter);
            this.Matrix0.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix0_KeyDownAfter);
            this.Matrix0.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix0_ChooseFromListBefore);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.Matrix2 = ((SAPbouiCOM.Matrix)(this.GetItem("oMatrix2").Specific));
            this.Matrix2.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix2_ClickAfter);
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_1").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_3").Specific));
            this.Matrix1.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix1_ClickAfter);
            this.Matrix1.KeyDownAfter += new SAPbouiCOM._IMatrixEvents_KeyDownAfterEventHandler(this.Matrix1_KeyDownAfter);
            this.Matrix1.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix1_ChooseFromListBefore);
            this.Matrix1.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix1_ChooseFromListAfter);
            this.OnCustomInitialize();
        }
        private void OnCustomInitialize()
        {
        }
        public Document_Settings_b1f()
        {
            try
            {
                //********************************************************************
                try
                {
                    oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                }
                catch (Exception)
                {
                    oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetForm("SBOAddonProject_Setting.Document_Settings_b1f", 1);
                }
                oDataTable = oForm.DataSources.DataTables.Add("TL_BP");
                oDataTable0 = oForm.DataSources.DataTables.Add("TL_BP3");
                oDataTable_0 = oForm.DataSources.DataTables.Add("TL_BP1");
                oDatatableIT = oForm.DataSources.DataTables.Add("TL_IT");
                oDataTableIT_0 = oForm.DataSources.DataTables.Add("TL_IT1");
                oDataTable1 = oForm.DataSources.DataTables.Add("TL_ProLine");
                oDataTable_1 = oForm.DataSources.DataTables.Add("TL_ProLine1");
                oDataTable2 = oForm.DataSources.DataTables.Add("TL_PriceList");
                oDataTable_2 = oForm.DataSources.DataTables.Add("TL_PriceList1");
                oDataTable3 = oForm.DataSources.DataTables.Add("TL_DocSet");
                oDataTable_3 = oForm.DataSources.DataTables.Add("TL_DocSet1");
                //********************************************************************
                oForm.Freeze(true);
                Add_ClearMenu();
                LoadData_DocSet();
                LoadDataIT();
                LoadDataBP();
                LoadData_PriceList();
                CheckBox();
                Matrix1.AutoResizeColumns();
                Matrix0.AutoResizeColumns();
                Matrix2.AutoResizeColumns();
                Folder0.Select();
                oForm.Freeze(false);
            }
            catch (Exception er)
            {
                //Application.SBO_Application.SetStatusBarMessage();
                oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetForm("SBOAddonProject_Setting.Document_Settings_b1f", 1);
                oForm.Select();
                //Application.SBO_Application.Forms.GetForm("SBOAddonProject_Setting.Document_Settings_b1f", oForm.TypeCount).Close();
            }
        }
        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.CloseAfter += new SAPbouiCOM.Framework.FormBase.CloseAfterHandler(this.Form_CloseAfter);
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            Application.SBO_Application.RightClickEvent += SBO_Application_RightClickEvent;
            Application.SBO_Application.MenuEvent += SBO_Application_MenuEvent;
        }
        public void Add_ClearMenu()
        {
            try
            {
                oForm = Application.SBO_Application.Forms.ActiveForm;
                oCreationPackage1 = (SAPbouiCOM.MenuCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage1.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage1.UniqueID = "Clear";
                oCreationPackage1.String = "Clear";
                oCreationPackage1.Enabled = true;
                oCreationPackage1.Position = -1;
                oForm.Menu.AddEx(oCreationPackage1);
                oCreationPackage1 = null;
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage("Form Already Exists.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                Application.SBO_Application.Forms.GetForm("SBOAddonProject_Setting.Document_Settings_b1f", oForm.TypeCount).Close();
            }
        }
        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                if (oForm.TypeEx == "SBOAddonProject_Setting.Document_Settings_b1f")
                {
                    if ((eventInfo.ItemUID == "oMatrix" || eventInfo.ItemUID == "Item_3"))
                    {
                        oForm.EnableMenu("1293", true);
                        oForm.EnableMenu("774", false);
                        oForm.EnableMenu("771", false);
                        //Application.SBO_Application.Menus.Item("Delete");
                        oForm.EnableMenu("Clear", true);
                    }
                    else
                    {
                        oForm.EnableMenu("Clear", false);
                        oForm.EnableMenu("1293", false);
                        oForm.EnableMenu("1292", false);
                    }

                }

            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }
        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction && !pVal.InnerEvent)
                {
                    if (pVal.MenuUID == "1293")
                    {
                        oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && oForm.TypeEx == "SBOAddonProject_Setting.Document_Settings_b1f")
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            btnAdd = (SAPbouiCOM.Button)oForm.Items.Item("btnAdd").Specific;
                            btnAdd.Caption = "Update";
                        }
                    }

                    if (pVal.MenuUID == "Clear")
                    {
                        oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                        if (oForm.PaneLevel == 1)
                        {
                            //int delete = Application.SBO_Application.MessageBox($"Do you want to clear all Customer?", 1, "Yes", "No");
                            //if (delete == 1)
                            //{
                            oForm.Freeze(true);
                            Matrix0 = (SAPbouiCOM.Matrix)oForm.Items.Item("oMatrix").Specific;
                            Matrix0.Clear();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            btnAdd = (SAPbouiCOM.Button)oForm.Items.Item("btnAdd").Specific;
                            btnAdd.Caption = "Update";
                            oForm.Freeze(false);
                            //}
                            //if (delete == 2)
                            //{
                            //    LoadDataBP();
                            //}
                        }
                        if (oForm.PaneLevel == 2)
                        {
                            //int delete = Application.SBO_Application.MessageBox($"Do you want to clear all Item?", 1, "Yes", "No");
                            //if (delete == 1)
                            //{
                            oForm.Freeze(true);
                            this.Matrix1 = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;
                            Matrix1.Clear();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            btnAdd = (SAPbouiCOM.Button)oForm.Items.Item("btnAdd").Specific;
                            btnAdd.Caption = "Update";
                            oForm.Freeze(false);

                            //}
                            //if (delete == 2)
                            //{
                            //    LoadDataIT();
                            //}
                        }
                    }
                    if (pVal.MenuUID == "1292")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oForm.PaneLevel == 1)
                            {
                                oForm.Freeze(true);
                                oDataTable.ExecuteQuery("DELETE FROM \"TL_Customer\" WHERE \"CardCode\"=''");
                                oDataTable.ExecuteQuery("INSERT INTO \"TL_Customer\" VALUES('','',CURRENT_TIMESTAMP)");
                                //oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                //btnAdd.Caption = "Update";
                                LoadDataBP();
                                oForm.Freeze(false);
                            }
                            if (oForm.PaneLevel == 2)
                            {
                                oForm.Freeze(true);
                                oDatatableIT.ExecuteQuery("DELETE FROM \"TL_ITEM\" WHERE \"ItemCode\"=''");
                                oDatatableIT.ExecuteQuery("INSERT INTO \"TL_ITEM\" VALUES('','',CURRENT_TIMESTAMP)");
                                //oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                //btnAdd.Caption = "Update";
                                LoadDataIT();
                                oForm.Freeze(false);
                            }
                        }
                    }
                }
            }
            catch (Exception er)
            {

                Application.SBO_Application.MessageBox(er.Message);
            }

        }
        private void Matrix0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new System.NotImplementedException();
            BubbleEvent = true;
            if (pVal.ColUID == "CardCode" && pVal.Row > 0)
            {
                ChooseFromList();
            }
        }

        public void Event_CFL(SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetForm("SBOAddonProject_Setting.Document_Settings_b1f", 1);
            for (int j = 0; j < CFLValue.SelectedObjects.Rows.Count; j++)
            {
                if (pVal.ItemUID == "oMatrix")
                {
                    oDataTable.SetValue("A", pVal.Row - 1 + j, j + pVal.Row);
                    oDataTable.SetValue("CardCode", pVal.Row - 1 + j, CFLValue.SelectedObjects.GetValue("CardCode", j).ToString());
                    oDataTable.SetValue("CardName", pVal.Row - 1 + j, CFLValue.SelectedObjects.GetValue("CardName", j).ToString());
                    if (Matrix0.RowCount > pVal.Row && j == 0)
                    {
                        Matrix0.AddRow();
                    }
                    else
                    {
                        oDataTable.Rows.Add();
                        Matrix0.AddRow();
                    }
                    Matrix0.AutoResizeColumns();
                    Matrix0.LoadFromDataSource();
                }
                if (pVal.ItemUID == "Item_3")
                {
                    oDatatableIT.SetValue("A", pVal.Row - 1 + j, j + pVal.Row);
                    oDatatableIT.SetValue("ItemCode", pVal.Row - 1 + j, CFLValue.SelectedObjects.GetValue("ItemCode", j).ToString());
                    oDatatableIT.SetValue("ItemName", pVal.Row - 1 + j, CFLValue.SelectedObjects.GetValue("ItemName", j).ToString());
                    if (Matrix1.RowCount > pVal.Row && j == 0)
                    {
                        Matrix1.AddRow();
                    }
                    else
                    {
                        oDatatableIT.Rows.Add();
                        Matrix1.AddRow();
                    }
                    Matrix1.AutoResizeColumns();
                    Matrix1.LoadFromDataSource();
                }
                Application.SBO_Application.StatusBar.SetSystemMessage($"Operating...[{j + 1}/{CFLValue.SelectedObjects.Rows.Count}]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
            }
            btnAdd.Caption = "Update";
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        }
        private void Matrix0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                CFLValue = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                if (pVal.ColUID == "CardCode")// & Matrix0.RowCount == pVal.Row)
                {
                    if (CFLValue.SelectedObjects != null & pVal.InnerEvent == true)
                    {
                        if (CFLValue.ActionSuccess == true)
                        {
                            bool ValueEx = false;
                            for (int i = 1; i <= Matrix0.RowCount; i++)
                            {
                                SAPbouiCOM.EditText text = (SAPbouiCOM.EditText)Matrix0.Columns.Item("CardCode").Cells.Item(i).Specific;
                                string C = CFLValue.SelectedObjects.GetValue("CardCode", 0).ToString();
                                if (text.Value == C && pVal.Row != i)
                                {
                                    Application.SBO_Application.MessageBox($"Customer Code ({text.Value.ToString()}) already existing.");
                                    ValueEx = true;
                                    break;
                                }
                            }
                            if (!ValueEx)
                            {
                                Event_CFL(pVal);
                            }

                        }

                    }

                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
            }
        }
        //Button Update*************************************************************************************************
        private void btnAdd_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.ActionSuccess != true)
            {
                if (btnAdd.Caption == "Add")
                {
                    if (oForm.PaneLevel == 1)
                    {
                        for (int i = 1; i <= Matrix0.GetNextSelectedRow(); i++)
                        {
                            SAPbouiCOM.EditText text = (SAPbouiCOM.EditText)Matrix0.Columns.Item("CardCode").Cells.Item(i).Specific;
                            if (text.Value != "")
                            {
                                ChooseFromList();
                            }
                            else
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                btnAdd.Caption = "Ok";
                            }
                        }
                    }
                    if (oForm.PaneLevel == 2)
                    {
                        for (int i = 1; i <= Matrix1.GetNextSelectedRow(); i++)
                        {
                            SAPbouiCOM.EditText text_1 = (SAPbouiCOM.EditText)Matrix1.Columns.Item("ItemCode").Cells.Item(i).Specific;
                            if (text_1.Value != "")
                            {
                                ChooseFromListIT();
                            }
                            else
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                btnAdd.Caption = "Ok";
                            }
                        }

                    }
                }
                if (pVal.FormMode == 2 && btnAdd.Caption == "Update")
                {
                    //if (oForm.PaneLevel == 1)
                    //{
                    Delete_BP();
                    Add_BP();
                    oDataTable.ExecuteQuery("DELETE FROM \"TL_Customer\" WHERE \"CardCode\"=''");
                    LoadDataBP();
                    //}
                    //if (oForm.PaneLevel == 2)
                    //{
                    Delete_IT();
                    Add_IT();
                    oDatatableIT.ExecuteQuery("DELETE FROM \"TL_ITEM\" WHERE \"ItemCode\"=''");
                    LoadDataIT();
                    //}
                    //if (oForm.PaneLevel == 3)
                    //{
                    Delete_PriceList();
                    Add_PriceList();
                    //}
                    Delete_DocSet();
                    Add_DocSet();
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    btnAdd.Caption = "Ok";
                    Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
                }
                if (pVal.FormMode == 1)
                {
                    if (oForm.PaneLevel == 1)
                    {
                        oDataTable.ExecuteQuery("DELETE FROM \"TL_Customer\" WHERE \"CardCode\"=''");
                    }
                    if (oForm.PaneLevel == 2)
                    {
                        oDatatableIT.ExecuteQuery("DELETE FROM \"TL_ITEM\" WHERE \"ItemCode\"=''");
                    }
                    oForm.Close();
                }
            }

        }
        //***************************************************************************************************************
        private void Button1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

                if (pVal.FormMode == 2)
                {
                    int message = Application.SBO_Application.MessageBox("Unsaved data will be lost. Do you want to continue without saving?", 1, "Yes", "No");
                    if (message == 1)
                    {
                        oForm.Close();
                    }
                    else
                    {
                        oForm.Visible = true;
                    }
                }
                else
                {
                    oForm.Close();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage(ex.Message);
            }
        }

        private void chFule_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            btnAdd.Caption = "Update";
        }

        private void chLPG_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            btnAdd.Caption = "Update";
        }

        private void chLube_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            btnAdd.Caption = "Update";
        }

        private void AddChooseFromListToEditTextBox(string ObjectType, string CFLUID, SAPbobsCOM.BoYesNoEnum Condtion, string ConAlias = "", string ConVal = "")
        {
            try
            {
                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = ObjectType;
                oCFLCreationParams.UniqueID = CFLUID;
                oCFL = oCFLs.Add(oCFLCreationParams);
                if (Condtion == SAPbobsCOM.BoYesNoEnum.tYES)
                {
                    oCons = oCFL.GetConditions();
                    oCon = oCons.Add();
                    oCon.Alias = ConAlias;
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = ConVal;
                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        private void CheckBox1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            btnAdd.Caption = "Update";
        }

        private void CheckBox2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            btnAdd.Caption = "Update";
        }

        private void Matrix2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            if (pVal.ColUID == "Active" & pVal.Row > 0)
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                btnAdd.Caption = "Update";
            }
            try
            {
                if (pVal.ItemUID == "oMatrix2" & pVal.ColUID == "Active" && pVal.Row == 0)
                {

                    string checkboxValue = oDataTable2.GetValue("Active", Matrix2.RowCount - 1).ToString();
                    // Toggle the checkbox value for all rows
                    oForm.Freeze(true);
                    for (int i = 0; i < Matrix2.RowCount; i++)
                    {
                        oDataTable2.SetValue("Active", i, (checkboxValue == "N") ? "Y" : "N");
                    }
                    // Refresh the matrix to update the checkbox display
                    //System.Threading.Thread.Sleep(1000);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    btnAdd.Caption = "Update";
                    Matrix2.LoadFromDataSource();
                    oForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                if (pVal.ItemUID == "oMatrix2" & pVal.ColUID == "Active" && pVal.Row == 0)
                {
                    oForm.Freeze(true);
                    string Query = "SELECT ROW_NUMBER() OVER(Order By A.\"ListNum\") As \"A\", A.\"ListNum\",A.\"ListName\",CASE WHEN B.\"ListNum\" IS NOT NULL AND B.\"ListName\" IS NOT NULL THEN 'Y' ELSE 'N' END \"Active\" FROM \"OPLN\" A LEFT OUTER JOIN \"TL_PriceList\" B ON A.\"ListNum\" = B.\"ListNum\" and A.\"ListName\" = B.\"ListName\"";
                    oDataTable2.ExecuteQuery(Query);
                    string checkboxValue = oDataTable2.GetValue("Active", Matrix2.RowCount - 1).ToString();
                    // Toggle the checkbox value for all rows
                    for (int i = 0; i < Matrix2.RowCount; i++)
                    {
                        oDataTable2.SetValue("Active", i, (checkboxValue == "N") ? "Y" : "N");
                    }
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    btnAdd.Caption = "Update";
                    Matrix2.LoadFromDataSource();
                    oForm.Freeze(false);
                }
            }
        }

        public void ChooseFromList()
        {
            oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.GetForm("SBOAddonProject_Setting.Document_Settings_b1f", 1);
            oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item("BP");
            SAPbouiCOM.ConditionsClass emptyCon;
            emptyCon = new SAPbouiCOM.ConditionsClass();
            oCFL.SetConditions(emptyCon);
            oCFLConditions = oCFL.GetConditions();
            oCFLCondition = oCFLConditions.Add();
            oCFLCondition.Alias = "CardType";
            oCFLCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCFLCondition.CondVal = "C";
            oCFL.SetConditions(oCFLConditions);
        }

        private void Matrix1_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            try
            {
                if (pVal.ItemUID == "Item_3" & pVal.ColUID == "ItemCode")
                {
                    btnAdd.Caption = "Add";
                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
            }

        }

        private void Matrix0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            try
            {
                if (pVal.Row > 0)
                {
                    if (pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_SHIFT)
                    {

                        Matrix0.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                        Matrix0.SelectRow(pVal.Row, true, true);
                    }
                    else
                    {
                        Matrix0.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                        Matrix0.SelectRow(pVal.Row, true, false);
                    }

                }

            }
            catch (Exception er)
            {

                Application.SBO_Application.MessageBox(er.Message);
            }
        }

        private void Matrix1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            if (pVal.Row > 0)
            {
                Matrix1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                Matrix1.SelectRow(pVal.Row, true, false);
            }
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (pVal.ActionSuccess == true && pVal.InnerEvent != true && oForm.Visible == true)
            {
                Matrix0 = (SAPbouiCOM.Matrix)oForm.Items.Item("oMatrix").Specific;
                Matrix1 = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_3").Specific;
                Matrix2 = (SAPbouiCOM.Matrix)oForm.Items.Item("oMatrix2").Specific;
                if (Matrix0.Item.Visible || Matrix1.Item.Visible || Matrix2.Item.Visible)
                {
                    Matrix0.AutoResizeColumns();
                    Matrix1.AutoResizeColumns();
                    Matrix2.AutoResizeColumns();
                }

            }
        }
        public void LoadDataBP()
        {
            //try
            //{
            oForm.Freeze(true);
            if (oDataTable.IsEmpty)
            {
                oDataTable.ExecuteQuery("DELETE FROM \"TL_Customer\" WHERE \"CardCode\"=''");
                oDataTable.ExecuteQuery("INSERT INTO \"TL_Customer\" VALUES('','',CURRENT_TIMESTAMP)");
            }
            oDataTable = oForm.DataSources.DataTables.Item("TL_BP");
            string Query = "SELECT ROW_NUMBER() OVER() AS \"A\",A.\"CardCode\",A.\"CardName\",A.\"UpdateDate\" FROM \"TL_Customer\" A";
            oDataTable.ExecuteQuery(Query);
            Matrix0.Columns.Item("#").DataBind.Bind("TL_BP", "A");
            Matrix0.Columns.Item("CardCode").DataBind.Bind("TL_BP", "CardCode");
            Matrix0.Columns.Item("CardName").DataBind.Bind("TL_BP", "CardName");
            Matrix0.LoadFromDataSource();
            Matrix0.AutoResizeColumns();
            oForm.Freeze(false);
            //}
            //catch (Exception er)
            //{
            //    Application.SBO_Application.MessageBox(er.Message);
            //}


        }

        private void Form_CloseAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            oDataTable.ExecuteQuery("DELETE FROM \"TL_Customer\" WHERE \"CardCode\"=''");
            oDatatableIT.ExecuteQuery("DELETE FROM \"TL_ITEM\" WHERE \"ItemCode\"=''");
        }

        private void Matrix0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            if (pVal.ItemUID == "oMatrix" && pVal.ColUID == "CardCode")
            {
                btnAdd.Caption = "Add";
            }
        }

        private void Matrix1_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            try
            {
                CFLValue = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                if (pVal.ColUID == "ItemCode")// & Matrix0.RowCount == pVal.Row)
                {
                    //LoadDataBP();
                    if (CFLValue.SelectedObjects != null & pVal.InnerEvent == true)
                    {
                        if (CFLValue.ActionSuccess == true)
                        {
                            bool ValueEx = false;
                            for (int i = 1; i <= Matrix1.RowCount; i++)
                            {
                                SAPbouiCOM.EditText text = (SAPbouiCOM.EditText)Matrix1.Columns.Item("ItemCode").Cells.Item(i).Specific;
                                string C = CFLValue.SelectedObjects.GetValue("ItemCode", 0).ToString();
                                if (text.Value == C && pVal.Row != i)
                                {
                                    Application.SBO_Application.MessageBox($"Item Code ({text.Value.ToString()}) already existing.");
                                    ValueEx = true;
                                    break;
                                }
                            }
                            if (!ValueEx)
                            {
                                Event_CFL(pVal);
                            }
                        }
                    }

                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.MessageBox(er.Message);
            }
        }

        private void Matrix1_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException();
            if (pVal.ColUID == "ItemCode" && pVal.Row > 0)
            {
                ChooseFromListIT();
            }
        }

        public void Delete_BP()
        {
            oDataTable_0.ExecuteQuery("DELETE  FROM \"TL_Customer\"");
        }
        public void Add_BP()
        {
            for (int i = 1; i <= Matrix0.VisualRowCount; i++)
            {
                object A = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("CardCode").Cells.Item(i).Specific).Value.ToString();
                object B = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("CardName").Cells.Item(i).Specific).Value.ToString();
                oDataTable.ExecuteQuery("INSERT INTO \"TL_Customer\" VALUES('" + A + "','" + B + "',CURRENT_TIMESTAMP)");
            }
        }
        public void ChooseFromListIT()
        {
            oCFLIT = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item("Item");
            SAPbouiCOM.ConditionsClass emptyCon;
            emptyCon = new SAPbouiCOM.ConditionsClass();
            oCFLIT.SetConditions(emptyCon);
            IoCFLConditions = oCFLIT.GetConditions();
            IoCFLCondition = IoCFLConditions.Add();
            IoCFLCondition.Alias = "SellItem";
            IoCFLCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            IoCFLCondition.CondVal = "Y";

            IoCFLCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            IoCFLCondition.BracketOpenNum = 1;
            IoCFLCondition = IoCFLConditions.Add();
            IoCFLCondition.Alias = "ItmsGrpCod";
            IoCFLCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            IoCFLCondition.CondVal = "100";

            IoCFLCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
            IoCFLCondition = IoCFLConditions.Add();
            IoCFLCondition.Alias = "ItmsGrpCod";
            IoCFLCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            IoCFLCondition.CondVal = "101";

            IoCFLCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
            IoCFLCondition = IoCFLConditions.Add();
            IoCFLCondition.Alias = "ItmsGrpCod";
            IoCFLCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            IoCFLCondition.CondVal = "102";

            IoCFLCondition.BracketCloseNum = 1;

            oCFLIT.SetConditions(IoCFLConditions);
        }
        public void LoadDataIT()
        {
            oForm.Freeze(true);
            if (oDatatableIT.IsEmpty)
            {
                oDatatableIT.ExecuteQuery("DELETE FROM \"TL_ITEM\" WHERE \"ItemCode\"=''");
                oDatatableIT.ExecuteQuery("INSERT INTO \"TL_ITEM\" VALUES('','',CURRENT_TIMESTAMP)");
            }
            string QueryIT = "SELECT ROW_NUMBER() OVER() AS \"A\",A.\"ItemCode\",A.\"ItemName\",A.\"UpdateDate\" FROM \"TL_ITEM\" A";
            oDatatableIT.ExecuteQuery(QueryIT);
            Matrix1.Columns.Item("#").DataBind.Bind("TL_IT", "A");
            Matrix1.Columns.Item("ItemCode").DataBind.Bind("TL_IT", "ItemCode");
            Matrix1.Columns.Item("ItemName").DataBind.Bind("TL_IT", "ItemName");
            Matrix1.LoadFromDataSource();
            Matrix1.AutoResizeColumns();
            oForm.Freeze(false);
        }
        public void Delete_IT()
        {
            oDataTableIT_0.ExecuteQuery("DELETE  FROM \"TL_ITEM\"");
        }
        public void Add_IT()
        {
            for (int i = 1; i <= Matrix1.VisualRowCount; i++)
            {
                object A = ((SAPbouiCOM.EditText)Matrix1.Columns.Item("ItemCode").Cells.Item(i).Specific).Value.ToString();
                object B = ((SAPbouiCOM.EditText)Matrix1.Columns.Item("ItemName").Cells.Item(i).Specific).Value.ToString();
                oDatatableIT.ExecuteQuery("INSERT INTO \"TL_ITEM\" VALUES('" + A + "','" + B + "',CURRENT_TIMESTAMP)");
            }
            //Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
        }

        public void LoadData_PriceList()
        {
            //try
            //{
            Matrix2.FlushToDataSource();
            string Query = "SELECT ROW_NUMBER() OVER(Order By A.\"ListNum\") As \"A\", A.\"ListNum\",A.\"ListName\",CASE WHEN B.\"ListNum\" IS NOT NULL AND B.\"ListName\" IS NOT NULL THEN 'Y' ELSE 'N' END \"Active\" FROM \"OPLN\" A LEFT OUTER JOIN \"TL_PriceList\" B ON A.\"ListNum\" = B.\"ListNum\" and A.\"ListName\" = B.\"ListName\"";
            oDataTable2.ExecuteQuery(Query);
            // Satrt Bind Data*********************************************
            Matrix2.Columns.Item("#").DataBind.Bind("TL_PriceList", "A");
            Matrix2.Columns.Item("ListNum").DataBind.Bind("TL_PriceList", "ListNum");
            Matrix2.Columns.Item("ListName").DataBind.Bind("TL_PriceList", "ListName");
            Matrix2.Columns.Item("Active").DataBind.Bind("TL_PriceList", "Active");
            //End Bind Data*********************************************
            Matrix2.AutoResizeColumns();
            Matrix2.LoadFromDataSource();
            //}
            //catch (Exception ex)
            //{
            //   Application.SBO_Application.SetStatusBarMessage(ex.Message);
            //}

        }

        public void Delete_PriceList()
        {
            string Query1 = "DELETE FROM \"TL_PriceList\" ";
            oDataTable2.ExecuteQuery(Query1);
            //Application.SBO_Application.SetStatusBarMessage("Updating...", SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }
        public void Add_PriceList()
        {
            for (int i = 1; i <= Matrix2.RowCount; i++)
            {
                SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)Matrix2.Columns.Item("Active").Cells.Item(i).Specific;

                if (oCheckBox.Checked == true)
                {
                    object A = ((SAPbouiCOM.EditText)Matrix2.Columns.Item("ListNum").Cells.Item(i).Specific).Value.ToString();
                    object B = ((SAPbouiCOM.EditText)Matrix2.Columns.Item("ListName").Cells.Item(i).Specific).Value.ToString();
                    string Query = "INSERT INTO \"TL_PriceList\"(\"ListNum\",\"ListName\",\"Active\",\"UpdateDate\") " +
                                    "VALUES('" + A + "','" + B + "','Y',CURRENT_TIMESTAMP)";
                    oDataTable2.ExecuteQuery(Query);
                }
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                btnAdd.Caption = "Update";
            }
            // Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "",0);
        }
        public void LoadData_DocSet()
        {
            //try
            //{
            CheckBox1 = (SAPbouiCOM.CheckBox)oForm.Items.Item("check1").Specific;
            CheckBox2 = (SAPbouiCOM.CheckBox)oForm.Items.Item("check2").Specific;
            string Query = "SELECT A.\"Lock_Sale\",A.\"Credit_Limit\" FROM \"TL_DocSet\" A";
            oDataTable3.ExecuteQuery(Query);
            CheckBox1.DataBind.Bind("TL_DocSet", "Lock_Sale");
            CheckBox2.DataBind.Bind("TL_DocSet", "Credit_Limit");
            //}
            //catch (Exception ex)
            //{
            //    Application.SBO_Application.SetStatusBarMessage(ex.Message);
            //}
        }
        public void Delete_DocSet()
        {
            try
            {
                string Query2 = "DELETE FROM \"TL_DocSet\"";
                oDataTable_3.ExecuteQuery(Query2);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage(ex.Message);
            }
        }
        public void Add_DocSet()
        {
            try
            {
                string A1 = "N";
                string A2 = "N";
                if (CheckBox1.Checked)
                {
                    A1 = "Y";
                }
                if (CheckBox2.Checked)
                {
                    A2 = "Y";
                }

                string QueryInsert = "INSERT INTO \"TL_DocSet\" (\"Lock_Sale\",\"Credit_Limit\",\"UpdateDate\") VALUES('" + A1 + "','" + A2 + "',CURRENT_TIMESTAMP )";
                oDataTable_3.ExecuteQuery(QueryInsert);
                //Application.SBO_Application.StatusBar.SetSystemMessage("Operation completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", "", 0);
            }
            catch (Exception ex)
            {

                Application.SBO_Application.SetStatusBarMessage(ex.Message);
            }

        }
        public void CheckBox()
        {
            CheckBox1.Item.Width = 20;
            CheckBox1.Item.Height = 18;
            CheckBox2.Item.Width = 20;
            CheckBox2.Item.Height = 18;
            //SAPbouiCOM.Item oNewItem = oForm.Items.Add("1212", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            //  oNewItem.Left = 45;
            //  oNewItem.Width = 300;
            //  oNewItem.Top = 130;
            //  oNewItem.Height = 100;
            //  oNewItem.FromPane = 2;
            //  oNewItem.ToPane = 2;
        }
    }
}
