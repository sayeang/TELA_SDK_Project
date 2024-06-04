
using Microsoft.VisualBasic;
using System;
namespace SysForm {
    public class SystemForm  {
        //*****************************************************************
        // At the begining of every UI API project we should first
        // establish connection with a running SBO application.
        // *******************************************************************   
        public SAPbouiCOM.Application SBO_Application;
        public SAPbouiCOM.Application oApp;

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
        public  SAPbouiCOM.DataTable dataTable, dataTable1, dataTable2, dataTable3;
        public SAPbouiCOM.UserDataSource oUserData;
        public SAPbouiCOM.DBDataSource oDBDataSource;
        public SAPbouiCOM.Folder oFolderItem;
        public SAPbouiCOM.Matrix oMatrix;
        public SAPbouiCOM.CheckBox oCheckBox, oCheckBox1, oCheckBox2, oCheckBox3;
        public SAPbouiCOM.Button oButton, oButton1, oButton2, oButton3;
        public SAPbouiCOM.EditText txtFind;
        public SAPbouiCOM.StaticText text, text1;
        public int i; 


        private void SetApplication() {

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
            
            SboGuiApi.Connect( sConnectionString ); 
            
            // get an initialized application object
            
            SBO_Application = SboGuiApi.GetApplication( -1 ); 
            
        } 
        
        
        public void AddItemsToOrderForm() {

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

        oItem = oOrderForm.Items.Item("234000064");
            oNewItem = oOrderForm.Items.Add("ok",SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Left = 130;
            oNewItem.Width = 25;
            oNewItem.Top = 140;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "...";
            oButton.ClickBefore += OButton_ClickBefore;

            oNewItem = oOrderForm.Items.Add("1212", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oNewItem.Left = 15;
            oNewItem.Width = 300;
            oNewItem.Top = 185;
            oNewItem.Height = 100;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;

            oNewItem = oOrderForm.Items.Add("text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Left = 15;
            oNewItem.Width = 100;
            oNewItem.Top = 165;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            text1 = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            text1.Caption = "Line of business";

            //************************************************************************************************
            oNewItem = oOrderForm.Items.Add("CheckBox", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oNewItem.Left = 30;
            oNewItem.Width = 100;
            oNewItem.Top = 190; //+ (i - 1) * 19;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            oCheckBox1 = ((SAPbouiCOM.CheckBox)(oNewItem.Specific));
            oCheckBox1.Caption = "Fuel";
            //oOrderForm.DataSources.UserDataSources.Add("Check1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            //oCheckBox1.DataBind.SetBound(true, "", "Check1");

            oNewItem = oOrderForm.Items.Add("CheckBox1", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oNewItem.Left = 30;
            oNewItem.Width = 100;
            oNewItem.Top = 210; //+ (i - 1) * 19;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            oCheckBox2 = ((SAPbouiCOM.CheckBox)(oNewItem.Specific));
            oCheckBox2.Caption = "LPG";
            //oOrderForm.DataSources.UserDataSources.Add("Check2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            //oCheckBox2.DataBind.SetBound(true, "", "Check2");

            oNewItem = oOrderForm.Items.Add("CheckBox2", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oNewItem.Left = 30;
            oNewItem.Width = 100;
            oNewItem.Top = 230; //+ (i - 1) * 19;
            oNewItem.Height = 19;
            oNewItem.FromPane = 4;
            oNewItem.ToPane = 4;
            oCheckBox3 = ((SAPbouiCOM.CheckBox)(oNewItem.Specific));
            oCheckBox3.Caption = "Lube";
            //oOrderForm.DataSources.UserDataSources.Add("Check3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            //oCheckBox3.DataBind.SetBound(true, "", "Check3");
            //oCheckBox.Checked=true;
            
            LoadDataLOB();
        }

        public SystemForm() {
            SetApplication();
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler( SBO_Application_ItemEvent );
        }
        public void AddMatrix()
        {
        // SAPbouiCOM.Form sForm = SBO_Application.Forms.ActiveForm;
            sForm = SBO_Application.Forms.Add("M14");
            sForm.Title = "List of Warehouse";
            sForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
           // sForm.Visible = false;
 
            mItem = sForm.Items.Add("Matrix1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            mItem.Left = 5;
            mItem.Width = 600;
            mItem.Top = 40;
            mItem.Height = 320;
            oMatrix = ((SAPbouiCOM.Matrix)(mItem.Specific));
            oColumns = oMatrix.Columns;
            oMatrix.DoubleClickAfter += OMatrix_DoubleClickAfter;
          //  mItem.Visible = false;

            oColumn = oColumns.Add("Nu", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            // Add a column for BP Card Code
            oColumn = oColumns.Add("BPLName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Branch Name";
            oColumn.Width = 150;
            oColumn.Editable = false;

            // Add a column for BP Card Name
            oColumn = oColumns.Add("WhsCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Warehouse Code";
            oColumn.Width = 150;
            oColumn.Editable = false;

            // Add a column for BP Card Phone
            oColumn = oColumns.Add("WhsName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Warehouse Name";
            oColumn.Width = 150;
            oColumn.Editable = false;
            //******************************************
            oColumn = oColumns.Add("Check", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Assigned";
            oColumn.Width = 30;
            oColumn.Editable = true;
            //******************************************
            oItem = sForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = 360;
            oItem.Height = 21;
            oButton1 = ((SAPbouiCOM.Button)(oItem.Specific));
            oButton1.Caption = "OK";
            oButton1.ClickBefore += OButton_ClickBefore1;
            //******************************************
            oItem = sForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 75;
            oItem.Width = 65;
            oItem.Top = 360;
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
            txtFind.TabOrder =0;
            txtFind.KeyDownBefore += TxtFind_KeyDownBefore;
            //******************************************
            oItem = sForm.Items.Add("BtnFind", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 170;
            oItem.Width = 65;
            oItem.Top = 10;
            oItem.Height = 20;
            SAPbouiCOM.Button  FButton = ((SAPbouiCOM.Button)(oItem.Specific));
            FButton.Caption = "Find";
            FButton.ClickBefore += FButton_ClickBefore;
            
            LoadData();
            sForm.Visible = true;
            sForm.Refresh();
        }

        private void TxtFind_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            // throw new NotImplementedException();
            BubbleEvent = true;

            //if ( pVal.ItemUID == "txtFind" && sForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            //{
            //    sForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            //}
           
        }

        private void OButton_ClickBefore1(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            try
            {
                if (sForm.Mode==SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    Delete();
                    Add();
                    //oOrderForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
            if (pVal.ColUID == "Check" && pVal.Row == 0)
            {
                //oUserData = oForm.DataSources.UserDataSources.Item("SYS_100");
                if (oUserData.Value == "Y")
                {

                    oUserData.Value = "N";
                    sForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    //mItem.Visible = false;
                }
                else
                {
                    oUserData.Value = "Y";
                    sForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                }
                // mItem.Visible = false;
                ///oMatrix.AutoResizeColumns();
                mItem.Visible = false;
                oMatrix.LoadFromDataSource();
                mItem.Visible = true;
            }
        }

        private void FButton_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            //throw new NotImplementedException();
            BubbleEvent = true;
            company = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            SAPbobsCOM.Users oUser = (SAPbobsCOM.Users)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
            SAPbouiCOM.DataTable dataTable = sForm.DataSources.DataTables.Item("DT_1");
            char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
            string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
            oUser.UserCode = title1;
            string textbox = txtFind.Value.ToString();
            string Query1 = "SELECT ROW_NUMBER() OVER(ORDER BY B.\"BPLId\",A.\"WhsCode\") AS \"A\",B.\"BPLName\",A.\"WhsCode\",A.\"WhsName\" FROM OWHS A LEFT OUTER JOIN OBPL B ON A.\"BPLid\" = B.\"BPLId\" LEFT OUTER JOIN USR6 C ON B.\"BPLId\"=C.\"BPLId\" WHERE C.\"UserID\"='" + title1 + "' and LOWER(A.\"WhsCode\") like LOWER('%" + textbox + "%')";
            dataTable.ExecuteQuery(Query1);
            if (dataTable.IsEmpty)
            {
                SBO_Application.MessageBox("Value is empty", 1, "Okay");
                txtFind.Value = "";
            }
            oMatrix.LoadFromDataSourceEx();
            //oMatrix.AutoResizeColumns();
        }
        public void LoadData()
        {
            //sForm = SBO_Application.Forms.ActiveForm;
            sForm = SBO_Application.Forms.Item("M14");
            oUserData = sForm.DataSources.UserDataSources.Add("SYS_100", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oApp = (SAPbouiCOM.Application)SBO_Application;
            company = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            oUser = (SAPbobsCOM.Users)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
            dataTable = sForm.DataSources.DataTables.Add("DT_1");
            char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
            string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
            oUser.UserCode = title1;
            string Query = "SELECT ROW_NUMBER() OVER(ORDER BY B.\"BPLId\",A.\"WhsCode\") AS \"A\",B.\"BPLName\",A.\"WhsCode\",A.\"WhsName\",CASE WHEN D.\"U_Warehouse\" IS NOT NULL and D.\"U_User\" IS NOT NULL THEN 'Y' ELSE 'N' END \"CheckBox\" FROM OWHS A LEFT OUTER JOIN OBPL B ON A.\"BPLid\" = B.\"BPLId\" LEFT OUTER JOIN USR6 C ON B.\"BPLId\"=C.\"BPLId\" LEFT OUTER JOIN \"@USRW\" D ON A.\"WhsCode\"=D.\"U_Warehouse\" and C.\"UserID\"=D.\"U_User\" WHERE C.\"UserID\"='" + title1 + "' ORDER BY B.\"BPLId\",A.\"WhsCode\"";
            dataTable.ExecuteQuery(Query);
            // Satrt Bind Data*********************************************
            oColumns.Item("Nu").DataBind.Bind("DT_1", "A");
            oColumns.Item("BPLName").DataBind.Bind("DT_1", "BPLName");
            oColumns.Item("WhsCode").DataBind.Bind("DT_1", "WhsCode");
            oColumns.Item("WhsName").DataBind.Bind("DT_1", "WhsName");
            oColumns.Item("Check").DataBind.Bind("DT_1", "CheckBox");
            //oColumns.Item("Check").DataBind.SetBound(true, "", "SYS_100");
            //End Bind Data*********************************************
            oUserData = sForm.DataSources.UserDataSources.Item("SYS_100");
            oUserData.Value = "Y";
            oMatrix.LoadFromDataSourceEx();
            oMatrix.AutoResizeColumns();
        }
        
        public void Delete()
        {
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Check").Cells.Item(i).Specific;

                if (oCheckBox.Checked == true)
                {
                    company = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                    SAPbobsCOM.Users oUser = (SAPbobsCOM.Users)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
                    SAPbouiCOM.DataTable dataTable = sForm.DataSources.DataTables.Item("DT_1");
                    char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
                    string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
                    oUser.UserCode = title1;
                    string Query1 = "DELETE FROM \"@USRW\" WHERE \"U_User\"='" + title1 + "' ";
                    dataTable.ExecuteQuery(Query1);
                    oApp.SetStatusBarMessage("Delete :"+i.ToString());
             }

            }
            oApp.SetStatusBarMessage("Scessfully");
            // sForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        }
        public void Add()
        {
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Check").Cells.Item(i).Specific;

                if (oCheckBox.Checked == true)
                {
                    company = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                    SAPbobsCOM.Users oUser = (SAPbobsCOM.Users)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
                    SAPbouiCOM.DataTable dataTable = sForm.DataSources.DataTables.Item("DT_1");
                    char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
            string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
                    oUser.UserCode = title1;
                    oUserTable = (SAPbobsCOM.UserTable)company.UserTables.Item("USRW");
                
                    string A = oUserTable.Code = DateTime.Now.ToString();
                    string B =  oUserTable.Name = i.ToString();
                    object C = oUserTable.UserFields.Fields.Item("U_User").Value = title1;
                  object D =  oUserTable.UserFields.Fields.Item("U_Warehouse").Value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("WhsCode").Cells.Item(i).Specific).Value.ToString();
                    //int ret = oUserTable.Add();
                    string Query = "INSERT INTO \"@USRW\"(\"Code\",\"LineId\",\"U_User\",\"U_Warehouse\") " +
                                    "VALUES('" + A + "','" + B + "','"+ C + "','" + D + "')";
                   // string Query1 = "DELETE FROM \"@USRW\"";
                    //dataTable.ExecuteQuery(Query1);
                    dataTable.ExecuteQuery(Query);
                    oApp.SetStatusBarMessage("Add :"+i.ToString());
              }
                
            }
            oApp.SetStatusBarMessage("Scessfully");
           // sForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        }

        public void AddUDT()
        {
            company = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Check").Cells.Item(i).Specific;

                if (oCheckBox.Checked == true)
                {
                    oUserTable = (SAPbobsCOM.UserTable)company.UserTables.Item("USRW");
                    oUserTable.Code = DateTime.Now.ToString();
                    oUserTable.Name = "50001";
                    oUserTable.UserFields.Fields.Item("U_User").Value = 12;
                    oUserTable.UserFields.Fields.Item("U_Warehouse").Value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("WhsCode").Cells.Item(i).Specific).Value.ToString();

                    int ret = oUserTable.Add();
                   if (ret==0)
                    {
                        oApp.MessageBox("" + i);
                    }
                    else
                    {
                        oMatrix.AddRow();
                        continue;
                    }
                }
                //continue;
            }
        }
       public void LoadDataLOB()
        {
            company = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            SAPbobsCOM.Users oUser = (SAPbobsCOM.Users)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

            char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
            string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
            oUser.UserCode = title1;
            dataTable1 = oOrderForm.DataSources.DataTables.Add("DT_2");
            dataTable2 = oOrderForm.DataSources.DataTables.Add("DT_3");
            dataTable3 = oOrderForm.DataSources.DataTables.Add("DT_4");
            string Query = "SELECT A.\"U_Fuel\",A.\"U_LPG\",A.\"U_Lube\" FROM \"@TL_LOB\" A  WHERE A.\"U_UserID\"='" + title1 + "'";
            dataTable1.ExecuteQuery(Query);
            dataTable2.ExecuteQuery(Query);
            dataTable3.ExecuteQuery(Query);
            oCheckBox1.DataBind.Bind("DT_2", "U_Fuel");
            oCheckBox2.DataBind.Bind("DT_2", "U_LPG");
            oCheckBox3.DataBind.Bind("DT_2", "U_Lube");
            
        }
        public void AddLOB()
        {
            if (oOrderForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                company = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                SAPbobsCOM.Users oUser = (SAPbobsCOM.Users)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
                //dataTable1 = oOrderForm.DataSources.DataTables.Add("DT_2");
                char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
                string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
                oUser.UserCode = title1;
                if (oCheckBox1.Checked)
                {
                    oUserTable = (SAPbobsCOM.UserTable)company.UserTables.Item("TL_LOB");
                    oUserTable.Name = title1;
                    oUserTable.UserFields.Fields.Item("U_UserID").Value = title1;
                    oUserTable.UserFields.Fields.Item("U_Fuel").Value.ToString();
                    oUserTable.Add();
                    //SBO_Application.SetStatusBarMessage("Fuel");
                }
                else if (oCheckBox1.Checked == true)
                {
                    oUserTable = (SAPbobsCOM.UserTable)company.UserTables.Item("TL_LOB");
                    oUserTable.Name = title1;
                    oUserTable.UserFields.Fields.Item("U_UserID").Value = title1;
                    oUserTable.UserFields.Fields.Item("U_LPG").Value.ToString();
                    oUserTable.Add();
                    //SBO_Application.SetStatusBarMessage("Fuel");
                }
                else if (oCheckBox1.Checked == true)
                {
                    oUserTable = (SAPbobsCOM.UserTable)company.UserTables.Item("TL_LOB");
                    oUserTable.Name = title1;
                    oUserTable.UserFields.Fields.Item("U_UserID").Value = title1;
                    oUserTable.UserFields.Fields.Item("U_Lube").Value.ToString();
                    oUserTable.Add();
                    //SBO_Application.SetStatusBarMessage("Fuel");
                }


            }
        }
        public void DeleteLOB()
        {
            company = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            SAPbobsCOM.Users oUser = (SAPbobsCOM.Users)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
            //dataTable1 = oOrderForm.DataSources.DataTables.Add("DT_2");
            char[] charsToTrim = { '<', '/', 'U', 'S', 'E', 'R', 'I', 'D', '>', '<', '/', 'U', 's', 'e', 'r', 'P', 'a', 'r', 'a', 'm', 's', '>' };
            string title1 = oOrderForm.BusinessObject.Key.Substring(60).Trim(charsToTrim);
            oUser.UserCode = title1;
            oUserTable = (SAPbobsCOM.UserTable)company.UserTables.Item("TL_WAREHOUSE");

            string Query1 = "DELETE FROM \"@TL_LOB\" WHERE \"U_UserID\"='" + title1 + "' ";
            dataTable1.ExecuteQuery(Query1);
            SBO_Application.SetStatusBarMessage("Delete :" + i.ToString());

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

            BubbleEvent = true;
            // sForm = SBO_Application.Forms.ActiveForm;

            if (((pVal.FormType == 20700 & pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) & (pVal.Before_Action == true)))
            {

                // get the event sending form
                oOrderForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

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
                //if (pVal.ItemUID == "CheckBox" || pVal.ItemUID == "CheckBox1" || pVal.ItemUID == "CheckBox2")
                //{
                //    if (oCheckBox1.Checked == true || oCheckBox2.Checked == true || oCheckBox3.Checked == true)
                //    {

                //        SBO_Application.SetStatusBarMessage("CheckBox : " + pVal.ItemUID);
                //    }
                //}

                if (pVal.ItemUID == "UserFolder" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.Before_Action == true)
                {
                    oOrderForm.PaneLevel = 4;
                    //  oOrderForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDAT//E_MODE;
                }
                if (pVal.FormType == 20700 & pVal.ItemUID == "1")
                {
                    DeleteLOB();
                    AddLOB();
                    //SBO_Application.SetStatusBarMessage("LOB");
                }
            }
            
        }

        private void SBO_Application_AppEvent( SAPbouiCOM.BoAppEventTypes EventType ) { 
            
            switch ( EventType ) {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    
                    // Take care of terminating your AddOn application
                    
                    SBO_Application.MessageBox( "A Shut Down Event has been caught" + Constants.vbNewLine + "Terminating 'Order Form Manipulation' Add On...", 1, "Ok", "", "" ); 
                    
                    System.Environment.Exit( 0 ); 
                    
                    break;
            }  
        }    
    }    
} 
