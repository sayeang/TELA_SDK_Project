
using Microsoft.VisualBasic;
using System;
namespace SysForm
{
    public class SystemForm
    {
        //*****************************************************************
        // At the begining of every UI API project we should first
        // establish connection with a running SBO application.
        // *******************************************************************   
        public SAPbouiCOM.Application SBO_Application;
        public SAPbouiCOM.Form oOrderForm;
        public SAPbouiCOM.Form sForm;
        public SAPbouiCOM.Item oNewItem;
        public SAPbobsCOM.Company company;
        public SAPbouiCOM.Item oItem;
        public SAPbouiCOM.Item mItem;
        public SAPbouiCOM.Item oItem1;
        public SAPbouiCOM.Column oColumn;
        public SAPbouiCOM.Columns oColumns;
        public SAPbobsCOM.UserTable oUserTable;
        public SAPbobsCOM.Users oUser;
        public SAPbouiCOM.DataTable dataTable, dataTable1, dataTable2, dataTable3;
        public SAPbouiCOM.UserDataSource oUserData;
        public SAPbouiCOM.Folder oFolderItem;
        public SAPbouiCOM.Matrix oMatrix;
        public SAPbouiCOM.CheckBox oCheckBox, oCheckBox1, oCheckBox2, oCheckBox3;
        public SAPbouiCOM.Button oButton, oButton1, oButton2, oButton3;
        public SAPbouiCOM.EditText txtFind;
        public SAPbouiCOM.StaticText text, text1;


        private void SetApplication()
        {

            // *******************************************************************
            // Use an SboGuiApi object to establish the connection
            // with the application and return an initialized appliction object
            // *******************************************************************
            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();

            // by following the steps specified above, the following
            // statment should be suficient for either development or run mode

            sConnectionString = Interaction.Command();

            // connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString);

            // get an initialized application object
            SBO_Application = SboGuiApi.GetApplication(-1);

        }


        public void AddItemsToOrderForm()
        {
            company = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            dataTable1 = oOrderForm.DataSources.DataTables.Add("DT_2");
            dataTable2 = oOrderForm.DataSources.DataTables.Add("DT_3");
            //*********Text Warehouse*************************************************************
            oItem = oOrderForm.Items.Item("234000064");
            oNewItem = oOrderForm.Items.Add("text", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Left = 15;
            oNewItem.Width = 90;
            oNewItem.Top = 140;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            text = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            text.Caption = "Warehouse";
            //*********Matrix Form*************************************************************
            oItem = oOrderForm.Items.Item("234000064");
            oNewItem = oOrderForm.Items.Add("ok", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Left = 130;
            oNewItem.Width = 25;
            oNewItem.Top = 140;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "...";
            oButton.ClickBefore += OButton_ClickBefore;
            //*********Text Line Of Business*************************************************************
            oNewItem = oOrderForm.Items.Add("text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Left = 15;
            oNewItem.Width = 100;
            oNewItem.Top = 165;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            text1 = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            text1.Caption = "Line of business";
            //*********Rectangle*************************************************************
            oNewItem = oOrderForm.Items.Add("1212", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oNewItem.Left = 15;
            oNewItem.Width = 300;
            oNewItem.Top = 185;
            oNewItem.Height = 100;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            //***********CheckBox1*****************************************************************************
            oNewItem = oOrderForm.Items.Add("CheckBox", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oNewItem.Left = 30;
            oNewItem.Width = 100;
            oNewItem.Top = 190; //+ (i - 1) * 19;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            oCheckBox1 = ((SAPbouiCOM.CheckBox)(oNewItem.Specific));
            oCheckBox1.Caption = "Fuel";
            oCheckBox1.Checked = false;
            //***********CheckBox2*****************************************************************************
            oNewItem = oOrderForm.Items.Add("CheckBox1", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oNewItem.Left = 30;
            oNewItem.Width = 100;
            oNewItem.Top = 210; //+ (i - 1) * 19;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            oCheckBox2 = ((SAPbouiCOM.CheckBox)(oNewItem.Specific));
            oCheckBox2.Caption = "LPG";
            oCheckBox2.Checked = false;
            //***********CheckBox3*****************************************************************************
            oNewItem = oOrderForm.Items.Add("CheckBox2", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oNewItem.Left = 30;
            oNewItem.Width = 100;
            oNewItem.Top = 230;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            oCheckBox3 = ((SAPbouiCOM.CheckBox)(oNewItem.Specific));
            oCheckBox3.Caption = "Lube";
            oCheckBox3.Checked = false;
            oOrderForm.Refresh();
            //LoadDataLOB();

        }

        public SystemForm()
        {
            SetApplication();
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
        }
        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            // throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                LoadDataLOB();
                oOrderForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Please select \"User Name\"");
            }

        }

        public void AddMatrix()
        {
            try
            {
            sForm = SBO_Application.Forms.Add("M08");
            oUserData = sForm.DataSources.UserDataSources.Add("SYS_100", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            dataTable = sForm.DataSources.DataTables.Add("DT_1");
            oUserTable = (SAPbobsCOM.UserTable)company.UserTables.Item("USRW");

            sForm.Title = "List of Warehouse";
            sForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            //*******************************************************************
            mItem = sForm.Items.Add("Matrix1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            mItem.Left = 5;
            mItem.Width = 600;
            mItem.Top = 40;
            mItem.Height = 300;
            oMatrix = ((SAPbouiCOM.Matrix)(mItem.Specific));
            oColumns = oMatrix.Columns;
            oMatrix.DoubleClickAfter += OMatrix_DoubleClickAfter;
                oMatrix.ClickAfter += OMatrix_ClickAfter;
            //*******************************************************************
            oColumn = oColumns.Add("Nu", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;
            //*******************************************************************
            oColumn = oColumns.Add("BPLName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Branch Name";
            oColumn.Width = 150;
            oColumn.Editable = false;
            //*******************************************************************
            oColumn = oColumns.Add("WhsCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Warehouse Code";
            oColumn.Width = 150;
            oColumn.Editable = false;
            //********************************************************************
            oColumn = oColumns.Add("WhsName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Warehouse Name";
            oColumn.Width = 150;
            oColumn.Editable = false;
            //***************************************************************************
            oColumn = oColumns.Add("Check", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Assigned";
            oColumn.Width = 30;
            oColumn.Editable = true;
            //oCheckBox = ((SAPbouiCOM.CheckBox)(mItem.Specific));
            //******************************************************************
            oItem = sForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = 350;
            oItem.Height = 21;
            oButton1 = ((SAPbouiCOM.Button)(oItem.Specific));
            oButton1.Caption = "OK";
            oButton1.ClickBefore += OButton_ClickBefore1;
            //*********************************************************************
            oItem = sForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 75;
            oItem.Width = 65;
            oItem.Top = 350;
            oItem.Height = 21;
            oButton2 = ((SAPbouiCOM.Button)(oItem.Specific));
            oButton2.Caption = "Cancel";
            //******************************************
            oItem = sForm.Items.Add("txtFind", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = 5;
            oItem.Width = 163;
            oItem.Top = 10;
            oItem.Height = 20;
            txtFind = ((SAPbouiCOM.EditText)(oItem.Specific));
            txtFind.TabOrder = 0;
            sForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            //******************************************
            oItem = sForm.Items.Add("BtnFind", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 170;
            oItem.Width = 65;
            oItem.Top = 10;
            oItem.Height = 20;
            SAPbouiCOM.Button FButton = ((SAPbouiCOM.Button)(oItem.Specific));
            FButton.Caption = "Find";
            FButton.ClickBefore += FButton_ClickBefore;

            LoadData();
            sForm.Visible = true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message);
                //throw;
            }
        }

        

        public void Count()
        {
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                SBO_Application.SetStatusBarMessage(i.ToString());
            }
        }
        public void LoadData()
        {
            try
            {
                
                //dataTable = sForm.DataSources.DataTables.Add("DT_1");
                char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
                string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
                string Query = "SELECT ROW_NUMBER() OVER(ORDER BY B.\"BPLId\",A.\"WhsCode\") AS \"A\",B.\"BPLName\",A.\"WhsCode\",A.\"WhsName\",CASE WHEN D.\"WhsCode\" IS NOT NULL and D.\"UserID\" IS NOT NULL THEN 'Y' ELSE 'N' END \"CheckBox\" FROM OWHS A LEFT OUTER JOIN OBPL B ON A.\"BPLid\" = B.\"BPLId\" LEFT OUTER JOIN USR6 C ON B.\"BPLId\"=C.\"BPLId\" LEFT OUTER JOIN \"TL_WAREH\" D ON A.\"WhsCode\"=D.\"WhsCode\" and C.\"UserID\"=D.\"UserID\" WHERE C.\"UserID\"='" + title1 + "' ORDER BY B.\"BPLId\",A.\"WhsCode\"";
                dataTable.ExecuteQuery(Query);
                // Satrt Bind Data*********************************************
                oColumns.Item("Nu").DataBind.Bind("DT_1", "A");
                oColumns.Item("BPLName").DataBind.Bind("DT_1", "BPLName");
                oColumns.Item("WhsCode").DataBind.Bind("DT_1", "WhsCode");
                oColumns.Item("WhsName").DataBind.Bind("DT_1", "WhsName");
                oColumns.Item("Check").DataBind.Bind("DT_1", "CheckBox");
                //End Bind Data*********************************************
                oMatrix.AutoResizeColumns();
                oMatrix.LoadFromDataSourceEx();
            }
            catch (Exception ex)
            {

                SBO_Application.SetStatusBarMessage("Please select \"User Name\"");
            }

        }
        public void Delete()
        {
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Check").Cells.Item(i).Specific;

                if (oCheckBox.Checked == true)
                {
                    SAPbouiCOM.DataTable dataTable = sForm.DataSources.DataTables.Item("DT_1");
                    char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
                    string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
                    string Query1 = "DELETE FROM \"TL_WAREH\" WHERE \"UserID\"='" + title1 + "' ";
                    dataTable.ExecuteQuery(Query1);
                    SBO_Application.SetStatusBarMessage("Updating...", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
            }
        }
        public void Add()
        {
            //SAPbobsCOM.Recordset oRecordSet1 = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
            string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Check").Cells.Item(i).Specific;

                if (oCheckBox.Checked == true)
                {
                    object A = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("WhsCode").Cells.Item(i).Specific).Value.ToString();
                    string Query = "INSERT INTO \"TL_WAREH\"(\"UserID\",\"WhsCode\",\"UpdateDate\") " +
                                    "VALUES('" + title1 + "','" + A + "',CURRENT_TIMESTAMP)";
                    dataTable.ExecuteQuery(Query);
                    SBO_Application.SetStatusBarMessage("Updating...", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
            }
        }
        public void LoadDataLOB()
        {
            try
            {
                char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
                string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
                string Query = "SELECT A.\"Lob_Fuel\",A.\"Lob_LPG\",A.\"Lob_Lube\" FROM \"TL_LOB\" A  WHERE A.\"UserID\"='" + title1 + "'";
                dataTable1.ExecuteQuery(Query);
                oCheckBox1.DataBind.Bind("DT_2", "Lob_Fuel");
                oCheckBox2.DataBind.Bind("DT_2", "Lob_LPG");
                oCheckBox3.DataBind.Bind("DT_2", "Lob_Lube");
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Please select \"User Name\"");
            }

        }
        public void DeleteLOB()
        {
            try
            {
                if (oOrderForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
                    string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
                    string Query2 = "DELETE FROM \"TL_LOB\" WHERE   \"UserID\"='" + title1 + "' ";
                    dataTable2.ExecuteQuery(Query2);
                    //SBO_Application.SetStatusBarMessage("Deleting...");
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Please select \"User Name\"");

            }


        }
        public void AddLOB()
        {
            try
            {
                if (oOrderForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
                    string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
                    SAPbobsCOM.Recordset oRecordSet1 = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string A1 = "N";
                    string A2 = "N";
                    string A3 = "N";
                    if (oCheckBox1.Checked)
                    {
                        A1 = "Y";
                    }
                    if (oCheckBox2.Checked)
                    {
                        A2 = "Y";
                    }
                    if (oCheckBox3.Checked)
                    {
                        A3 = "Y";
                    }

                    string QueryInsert = "INSERT INTO \"TL_LOB\" (\"UserID\",\"Lob_Fuel\",\"Lob_LPG\",\"Lob_Lube\",\"UpdateDate\") VALUES('" + title1 + "','" + A1 + "','" + A2 + "','" + A3 + "',CURRENT_TIMESTAMP )";
                    oRecordSet1.DoQuery(QueryInsert);
                    SBO_Application.SetStatusBarMessage("Updating...", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
            }
            catch (Exception ex)
            {

                SBO_Application.SetStatusBarMessage("Please select \"User Name\"" + ex.Message);
            }

        }

        private void OButton_ClickBefore1(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if (sForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    Delete();
                    Add();
                }

            }
            catch (Exception er)
            {
                SBO_Application.MessageBox("error :" + er);
                //throw;
            }
        }
        private void OMatrix_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new NotImplementedException();
            try
            {
                if (pVal.ColUID == "Check" && pVal.Row == 0)
                {
                    mItem.Visible = false;
                    for (int i = 0; i < oMatrix.RowCount; i++)
                    {
                        
                        dataTable.SetValue("CheckBox", i, "Y");
                       
                    }
                    sForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    oMatrix.LoadFromDataSource();
                    mItem.Visible = true;
                }
                mItem.Visible = true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Error");
            }
        }
        private void OMatrix_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new NotImplementedException();
            try
            {
                if (pVal.ColUID == "Check" && pVal.Row == 0)
                {
                    //mItem.Visible = false;
                    for (int i = 0; i < oMatrix.RowCount; i++)
                    {
                        dataTable.SetValue("CheckBox", i, "N");
                    }
                    sForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    oMatrix.LoadFromDataSource();
                   mItem.Visible = true;
                }
                //mItem.Visible = true;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Warehouse updated",SAPbouiCOM.BoMessageTime.bmt_Short,false);
                sForm.Close();
                //throw;
            }
        }
        private void FButton_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;
            try
            {
                char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
                string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
                string textbox = txtFind.Value.ToString();
                string Query1 = "SELECT ROW_NUMBER() OVER(ORDER BY B.\"BPLId\",A.\"WhsCode\") AS \"A\",B.\"BPLName\",A.\"WhsCode\",A.\"WhsName\",CASE WHEN D.\"WhsCode\" IS NOT NULL and D.\"UserID\" IS NOT NULL THEN 'Y' ELSE 'N' END \"CheckBox\" FROM OWHS A LEFT OUTER JOIN OBPL B ON A.\"BPLid\" = B.\"BPLId\" LEFT OUTER JOIN USR6 C ON B.\"BPLId\"=C.\"BPLId\" LEFT OUTER JOIN \"TL_WAREH\" D ON A.\"WhsCode\"=D.\"WhsCode\" and C.\"UserID\"=D.\"UserID\" WHERE C.\"UserID\"='" + title1 + "' and LOWER(A.\"WhsCode\") like LOWER('%" + textbox + "%')";
                dataTable.ExecuteQuery(Query1);
                if (dataTable.IsEmpty)
                {
                    SBO_Application.MessageBox("Value is empty", 1, "Okay");
                    txtFind.Value = "";
                }
                if (!string.IsNullOrEmpty(textbox))
                {
                    mItem.Visible = false;
                    oMatrix.LoadFromDataSourceEx();
                    mItem.Visible = true;
                }
                else
                {
                        mItem.Visible = false;
                        oMatrix.LoadFromDataSourceEx();
                        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None;
                        mItem.Visible = true;
                }
                sForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

        }

        private void OButton_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            AddMatrix();
            // AddDataeble();
        }
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            // sForm = SBO_Application.Forms.ActiveForm;
            BubbleEvent = true;
            if (((pVal.FormType == 20700 & pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) & (pVal.Before_Action == true)))
            {
                oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                //oOrderForm = SBO_Application.Forms.Item(pVal.FormUID);
                if (((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) & (pVal.Before_Action == true)))
                {
                    // add a new folder item to the form
                    oNewItem = oOrderForm.Items.Add("UserFolder", SAPbouiCOM.BoFormItemTypes.it_FOLDER);

                    // use an existing folder item for grouping and setting the
                    // items properties (such as location properties)
                    // use the 'Display Debug Information' option (under 'Tools')
                    // in the application to acquire the UID of the desired folder
                    oItem = oOrderForm.Items.Item("234000058");

                    oNewItem.Top = oItem.Top;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Left = oItem.Left + oItem.Width;

                    oFolderItem = ((SAPbouiCOM.Folder)(oNewItem.Specific));

                    oFolderItem.Caption = "User Folder";
                    // group the folder with the desired folder item
                    oFolderItem.GroupWith("234000058");
                    // add your own items to the form

                    AddItemsToOrderForm();

                    oOrderForm.PaneLevel = 1;

                }
                if (pVal.ItemUID == "UserFolder" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.Before_Action == true)
                {
                    oOrderForm.PaneLevel = 4;
                    if (oOrderForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        oOrderForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    }
                }
                if (pVal.FormType == 20700 & pVal.ItemUID == "1" & oOrderForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    DeleteLOB();
                    AddLOB();
                }
            }
        }
        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
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
