using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Threading;
using INI;

namespace MASTER_SET_CONTROL
{
    public partial class frmManage : Form
    {

        MySqlConnection conn;
        IniFile iniconfig;
        BindingSource bs = new BindingSource();
        BindingList<DataTable> tables = new BindingList<DataTable>();

        public string strCon;
        string strSQL;

        public string strAdminID = "";
        List<int> items = new List<int>() { 2, 4 };

        public string select_db;

        public frmManage()
        {
            InitializeComponent();


        }

        public bool Connection() // open connection db
        {
            conn = new MySqlConnection(strCon);
            try
            {
                conn.Open();
                return true;
            }
            catch (MySqlException ex)
            {
                switch (ex.Number)
                {
                    case 0: // Can't connect to client server.
                        break;
                    case 1042: // Unable to connect to any of the specified MySQL hosts.
                        break;
                    case 1045: // Access denied for user 'root'@'localhost' (using password: YES)
                        break;
                    case 1049: // Unknown database ''
                        break;
                }
            }
            return false;
        }


        public void frmManage_Load(object sender, EventArgs e)
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            if (select_db == "SAMPLE SET CONTROL")
            {
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                strCon = "Server=" + IP + ";";
                strCon += "Uid=root;";
                strCon += "Password=123456*;";
                strCon += "Database=" + DB + ";";

                conn = new MySqlConnection(strCon);

                conn.Open();
                //strCon = "Server=43.72.52.12;Database=eia_master_set_control;Uid=root;Password=123456*;Convert Zero Datetime=True;";
                //strCon = "host=localhost;Database=eia_master_set_control;Uid=root;Password=123456;Convert Zero Datetime=True;";
                this.Text = "SAMPLE SET CONTROL FORM MANAGE" + String.Format(" --- Version {0}", version) + " - Server : " + IP;
            }
            else if (select_db == "ENGINEERING TRAINING CENTER")
            {
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB = iniconfig.IniReadValue("MySQL_Server", "DB2");

                strCon = "Server=" + IP + ";";
                strCon += "Uid=root;";
                strCon += "Password=123456*;";
                strCon += "Database=" + DB + ";";

                conn = new MySqlConnection(strCon);
                conn.Open();

                //strCon = "Server=43.72.52.12;Database=eia_master_set_control_spacial;Uid=root;Password=123456*;Convert Zero Datetime=True;";
                //strCon = "host=localhost;Database=eia_master_set_control_spacial;Uid=root;Password=123456;Convert Zero Datetime=True;";
                this.Text = "ENGINEERING TRAINING CENTER FORM MANAGE" + String.Format(" --- Version {0}", version) + " - Server : " + IP;
            }
            else
            {
                MessageBox.Show("Not connect to server!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Load_Function();


        }

        public void Load_Function()
        {
            PleaseWait.Create();
            try
            {
                getMaster();
            }
            finally
            {
                PleaseWait.Destroy();
            }



            //MessageBox.Show(strAdminID);

            comboBoxEdit.Items.Add("SET_NO");
            comboBoxEdit.Items.Add("SECTION");
            comboBoxEdit.Items.Add("MODEL");
            comboBoxEdit.Items.Add("DETAIL");
            comboBoxEdit.Items.Add("PCN_NO");
            comboBoxEdit.Items.Add("SERIES");
            comboBoxEdit.Items.Add("EVENT");
            comboBoxEdit.Items.Add("SN_SET");
            comboBoxEdit.Items.Add("PURPOSE");
            comboBoxEdit.Items.Add("COL_WHERE");
            comboBoxEdit.SelectedIndex = comboBoxEdit.FindStringExact("");

            comboBoxAdd_New.Items.Add("SET_NO");
            comboBoxAdd_New.Items.Add("SECTION");
            comboBoxAdd_New.Items.Add("MODEL");
            comboBoxAdd_New.Items.Add("DETAIL");
            comboBoxAdd_New.Items.Add("PCN_NO");
            comboBoxAdd_New.Items.Add("SERIES");
            comboBoxAdd_New.Items.Add("EVENT");
            comboBoxAdd_New.Items.Add("SN_SET");
            comboBoxAdd_New.Items.Add("PURPOSE");
            comboBoxAdd_New.Items.Add("COL_WHERE");
            comboBoxAdd_New.SelectedIndex = comboBoxAdd_New.FindStringExact("");

            comboBoxTransfer.Items.Add("SET_NO");
            comboBoxTransfer.Items.Add("SECTION");
            comboBoxTransfer.Items.Add("MODEL");
            comboBoxTransfer.Items.Add("DETAIL");
            comboBoxTransfer.Items.Add("PCN_NO");
            comboBoxTransfer.Items.Add("SERIES");
            comboBoxTransfer.Items.Add("EVENT");
            comboBoxTransfer.Items.Add("SN_SET");
            comboBoxTransfer.Items.Add("PURPOSE");
            comboBoxTransfer.Items.Add("COL_WHERE");
            comboBoxTransfer.SelectedIndex = comboBoxTransfer.FindStringExact("");

            comboBoxDispose.Items.Add("SET_NO");
            comboBoxDispose.Items.Add("SECTION");
            comboBoxDispose.Items.Add("MODEL");
            comboBoxDispose.Items.Add("DETAIL");
            comboBoxDispose.Items.Add("PCN_NO");
            comboBoxDispose.Items.Add("SERIES");
            comboBoxDispose.Items.Add("EVENT");
            comboBoxDispose.Items.Add("SN_SET");
            comboBoxDispose.Items.Add("PURPOSE");
            comboBoxDispose.Items.Add("COL_WHERE");
            comboBoxDispose.SelectedIndex = comboBoxDispose.FindStringExact("");
        }


        public void closeCon()
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }

        }

        public void getMaster()
        {
            dtGridViwerHitory.DataSource = null;
            dtGridViwerHitory.Rows.Clear();
            dtGridViwerHitory.Refresh();
            bindingNavigator1.BindingSource = null;
            tables.Clear();



            strSQL = "SELECT * FROM `master_set` order by auto_id";

            // 1
            // Open connection
            Connection();

            {

                // 2
                // Create new DataAdapter
                using (MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn))
                {
                    // 3
                    // Use DataAdapter to fill DataTable
                    DataTable t = new DataTable();

                    da.Fill(t);
                    // 4
                    // Render data onto the screen
                    dtGridViwerHitory.DataSource = t;

                    if (t.Rows.Count != 0)
                    {
                        int count = 0;
                        DataTable dt = null;

                        foreach (DataRow dr in t.Rows)
                        {

                            if (count == 0)
                            {
                                dt = t.Clone();
                                tables.Add(dt);
                            }
                            dt.Rows.Add(dr.ItemArray);
                            count++;
                            if (count >= 50)
                            {
                                count = 0;
                            }
                            //dt.Rows.Clear();
                            dtGridViwerHitory.DataSource = null;
                        }

                        bindingNavigator1.BindingSource = bs;
                        bs.DataSource = tables;
                        bs.PositionChanged += bs_PositionChanged;
                        bs_PositionChanged(bs, EventArgs.Empty);

                    }

                }
            }
            closeCon();

        }

        void bs_PositionChanged(object sender, EventArgs e)
        {
            if (tables.Count != 0)
            {
                dtGridViwerHitory.DataSource = tables[bs.Position];
            }
        }


        private void dtGridViwerHitory_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dtGridViwerHitory.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            if (dtGridViwerHitory.SelectedCells.Count > 0 & e.ColumnIndex >= 0 & e.RowIndex >= 0)
            {

                dtGridViwerHitory.CurrentCell = dtGridViwerHitory.Rows[e.RowIndex].Cells[e.ColumnIndex];
                //Can leave these here - doesn't hurt
                dtGridViwerHitory.Rows[e.RowIndex].Selected = true;
                dtGridViwerHitory.Focus();

                int selectedrowindex = dtGridViwerHitory.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = dtGridViwerHitory.Rows[selectedrowindex];

                txtSetNoEdit.Text = Convert.ToString(selectedRow.Cells[1].Value.ToString());
                txtSectionEdit.Text = Convert.ToString(selectedRow.Cells[2].Value.ToString());
                txtModelEdit.Text = Convert.ToString(selectedRow.Cells[3].Value.ToString());
                txtDetailEdit.Text = Convert.ToString(selectedRow.Cells[4].Value.ToString());
                txtPCNEdit.Text = Convert.ToString(selectedRow.Cells[5].Value.ToString());
                txtSeriesEdit.Text = Convert.ToString(selectedRow.Cells[6].Value.ToString());
                txtEventEdit.Text = Convert.ToString(selectedRow.Cells[7].Value.ToString());
                txtSNSetEdit.Text = Convert.ToString(selectedRow.Cells[8].Value.ToString());
                txtPurposeEdit.Text = Convert.ToString(selectedRow.Cells[9].Value.ToString());
                txtWhereEdit.Text = Convert.ToString(selectedRow.Cells[10].Value.ToString());

                txtSetNoTransfer.Text = Convert.ToString(selectedRow.Cells[1].Value.ToString());
                txtSectionTransfer.Text = Convert.ToString(selectedRow.Cells[2].Value.ToString());
                txtModelTransfer.Text = Convert.ToString(selectedRow.Cells[3].Value.ToString());
                txtDetailTransfer.Text = Convert.ToString(selectedRow.Cells[4].Value.ToString());
                txtPCNTransfer.Text = Convert.ToString(selectedRow.Cells[5].Value.ToString());
                txtSeriesTransfer.Text = Convert.ToString(selectedRow.Cells[6].Value.ToString());
                txtEventTransfer.Text = Convert.ToString(selectedRow.Cells[7].Value.ToString());
                txtSNSetTransfer.Text = Convert.ToString(selectedRow.Cells[8].Value.ToString());
                txtPurposeTransfer.Text = Convert.ToString(selectedRow.Cells[9].Value.ToString());
                txtWhereTransfer.Text = Convert.ToString(selectedRow.Cells[10].Value.ToString());

                txtSetNoDispose.Text = Convert.ToString(selectedRow.Cells[1].Value.ToString());
                txtSectionDispose.Text = Convert.ToString(selectedRow.Cells[2].Value.ToString());
                txtModelDispose.Text = Convert.ToString(selectedRow.Cells[3].Value.ToString());
                txtDetailDispose.Text = Convert.ToString(selectedRow.Cells[4].Value.ToString());
                txtPCNDispose.Text = Convert.ToString(selectedRow.Cells[5].Value.ToString());
                txtSeriesDispose.Text = Convert.ToString(selectedRow.Cells[6].Value.ToString());
                txtEventDispose.Text = Convert.ToString(selectedRow.Cells[7].Value.ToString());
                txtSNSetDispose.Text = Convert.ToString(selectedRow.Cells[8].Value.ToString());
                txtPurposeDispose.Text = Convert.ToString(selectedRow.Cells[9].Value.ToString());
                txtWhereDispose.Text = Convert.ToString(selectedRow.Cells[10].Value.ToString());

            }

        }

        private void btAdd_Click(object sender, EventArgs e)
        {
            if (txtSetNo.Text.Trim() == "" | txtSection.Text.Trim() == "" | txtModel.Text.Trim() == "" | txtDetail.Text.Trim() == "" | txtPCN.Text.Trim() == "" | txtSeries.Text.Trim() == "" | txtEvent.Text.Trim() == "")
            {
                MessageBox.Show("Please fill all information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                if (checkDup(txtSetNo.Text.Trim()) == false)
                {
                    DialogResult dialogResult = MessageBox.Show("Register new item: " + txtSetNo.Text.Trim() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        PleaseWait.Create();
                        try
                        {
                            InsertMaster(txtSetNo.Text.Trim(), txtSection.Text.Trim(), txtModel.Text.Trim(), txtDetail.Text.Trim(), txtPCN.Text.Trim(), txtSeries.Text.Trim(), txtEvent.Text.Trim(), txtSNSet.Text.Trim(), txtPurpose.Text.Trim(), txtWhere.Text.Trim());
                            InsertHistory(txtSetNo.Text.Trim(), txtSection.Text.Trim(), txtModel.Text.Trim(), txtDetail.Text.Trim(), txtPCN.Text.Trim(), txtSeries.Text.Trim(), txtEvent.Text.Trim(), txtSNSet.Text.Trim(), txtPurpose.Text.Trim(), txtWhere.Text.Trim(), "", "", "", strAdminID, "Add New Master");
                            getMaster();
                        }
                        finally
                        {
                            PleaseWait.Destroy();
                            MessageBox.Show("Registeration Completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else if (dialogResult == DialogResult.No)
                    {

                    }

                }
                else
                {
                    MessageBox.Show("This set was registered!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private Boolean checkDup(string strSetNo)
        {
            string strSQL = "select * from master_set where SET_NO = '" + strSetNo + "'";

            Connection();
            MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
            DataTable dt = new DataTable();

            da.Fill(dt);
            foreach (DataRow drData in dt.Rows)
            {
                return true;

            }
            closeCon();

            return false;
        }

        public void InsertMaster(string strSetNo, string strSection, string strModel, string strDetail, string strPCN, string strSeries, string strEvent, string strSNSet, string strPurpose, string strWhere)
        {

            Connection();
            string command = "";

            command = "insert into master_set (`SET_NO`,`SECTION`,`MODEL`,`DETAIL`,`PCN_NO`,`SERIES`,`EVENT`,SN_SET,PURPOSE,COL_WHERE)values('" +
                strSetNo + "','" + strSection + "','" + strModel + "','" + strDetail + "','" + strPCN + "','" + strSeries + "','" + strEvent + "','" + strSNSet + "','" + strPurpose + "','" + strWhere + "')";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();


        }

        public void InsertHistory(string strSetNo, string strSection, string strModel, string strDetail, string strPCN, string strSeries, string strEvent, string strSNSet, string strPurpose, string strWhere, string strTransfer, string strTransferPic, string strDispose, string strRecName, string strHistoryDetails)
        {

            Connection();
            string command = "";

            command = "insert into history_log (`SET_NO`,`SECTION`,`MODEL`,`DETAIL`,`PCN_NO`,`SERIES`,`EVENT`,SN_SET,PURPOSE,COL_WHERE, "
                   + " TRANSFER,TRANSFER_PIC,DISPOSAL,REC_DATE,REC_NAME,HISTORY_DETAILS)values('" +
                strSetNo + "','" + strSection + "','" + strModel + "','" + strDetail + "','" + strPCN + "','" + strSeries + "','" + strEvent + "','" + strSNSet + "','" + strPurpose + "','" + strWhere
                + "','" + strTransfer + "','" + strTransferPic + "','" + strDispose + "',sysdate(),'" + strRecName + "','" + strHistoryDetails + "')";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }

        public void EditMaster(string strSetNo, string strSection, string strModel, string strDetail, string strPCN, string strSeries, string strEvent, string strSNSet, string strPurpose, string strWhere)
        {

            Connection();
            string command = "";

            command = "update master_set set SECTION = '" + strSection + "',MODEL = '" + strModel + "',DETAIL = '" + strDetail + "',PCN_NO = '" + strPCN + "',SERIES = '" + strSeries + "',EVENT = '" + strEvent + "',SN_SET='" + strSNSet + "',PURPOSE='" + strPurpose + "',COL_WHERE ='" + strWhere + "'"
                + " where SET_NO = '" + strSetNo + "'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }

        public void InsertTransfer(string strSetNo, string strSection, string strModel, string strDetail, string strPCN, string strSeries, string strEvent, string strTransfer, string strTransferPic, string strSNSet, string strPurpose, string strWhere)
        {

            Connection();
            string command = "";

            command = "insert into master_set_transfer (`SET_NO`,`SECTION`,`MODEL`,`DETAIL`,`PCN_NO`,`SERIES`,`EVENT`,`TRANSFER`,`TRANSFER_PIC`,`rec_date`,SN_SET,PURPOSE,COL_WHERE)values('" +
                strSetNo + "','" + strSection + "','" + strModel + "','" + strDetail + "','" + strPCN + "','" + strSeries + "','" + strEvent + "','" + strTransfer + "','" + strTransferPic + "',sysdate(),'" + strSNSet + "','" + strPurpose + "','" + strWhere + "')";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

            deleteMaster(strSetNo);



        }

        public void InsertDispose(string strSetNo, string strSection, string strModel, string strDetail, string strPCN, string strSeries, string strEvent, string strDispose, string strSNSet, string strPurpose, string strWhere)
        {

            Connection();
            string command = "";

            command = "insert into master_set_dispose (`SET_NO`,`SECTION`,`MODEL`,`DETAIL`,`PCN_NO`,`SERIES`,`EVENT`,`DISPOSAL`,`rec_date`,SN_SET,PURPOSE,COL_WHERE)values('" +
                strSetNo + "','" + strSection + "','" + strModel + "','" + strDetail + "','" + strPCN + "','" + strSeries + "','" + strEvent + "','" + strDispose + "',sysdate(),'" + strSNSet + "','" + strPurpose + "','" + strWhere + "')";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

            deleteMaster(strSetNo);


        }

        public void deleteMaster(string strSetNo)
        {

            Connection();
            string command = "";

            command = "delete from master_set where `SET_NO`='" + strSetNo + "'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }

        private void btEdit_Click(object sender, EventArgs e)
        {

            if (txtSetNoEdit.Text.Trim() == "" | txtSectionEdit.Text.Trim() == "" | txtModelEdit.Text.Trim() == "" | txtDetailEdit.Text.Trim() == "" | txtPCNEdit.Text.Trim() == "" | txtSeriesEdit.Text.Trim() == "" | txtEventEdit.Text.Trim() == "")
            {
                MessageBox.Show("Please fill all information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                if (checkDup(txtSetNoEdit.Text.Trim()) == true)
                {
                    DialogResult dialogResult = MessageBox.Show("Edit item: " + txtSetNoEdit.Text.Trim() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        PleaseWait.Create();
                        try
                        {
                            EditMaster(txtSetNoEdit.Text.Trim(), txtSectionEdit.Text.Trim(), txtModelEdit.Text.Trim(), txtDetailEdit.Text.Trim(), txtPCNEdit.Text.Trim(), txtSeriesEdit.Text.Trim(), txtEventEdit.Text.Trim(), txtSNSetEdit.Text.Trim(), txtPurposeEdit.Text.Trim(), txtWhereEdit.Text.Trim());
                            InsertHistory(txtSetNoEdit.Text.Trim(), txtSectionEdit.Text.Trim(), txtModelEdit.Text.Trim(), txtDetailEdit.Text.Trim(), txtPCNEdit.Text.Trim(), txtSeriesEdit.Text.Trim(), txtEventEdit.Text.Trim(), txtSNSetEdit.Text.Trim(), txtPurposeEdit.Text.Trim(), txtWhereEdit.Text.Trim(), "", "", "", strAdminID, "Edit Master");
                            getMaster();
                        }
                        finally
                        {
                            PleaseWait.Destroy();
                            MessageBox.Show("Edit Completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else if (dialogResult == DialogResult.No)
                    {

                    }

                }
                else
                {
                    MessageBox.Show("Have no set information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        private void btTransfer_Click(object sender, EventArgs e)
        {

            if (txtSetNoTransfer.Text.Trim() == "" | txtTransfer.Text.Trim() == "" | txtTransferPic.Text.Trim() == "")
            {
                MessageBox.Show("Please fill all information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {

                if (checkBorrow(txtSetNoTransfer.Text.Trim()) == true)
                {
                    MessageBox.Show("This set was borrowed, please return before transfer!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (checkDup(txtSetNoTransfer.Text.Trim()) == true)
                {
                    DialogResult dialogResult = MessageBox.Show("Transfer item: " + txtSetNoTransfer.Text.Trim() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        PleaseWait.Create();
                        try
                        {
                            InsertTransfer(txtSetNoTransfer.Text.Trim(), txtSectionTransfer.Text.Trim(), txtModelTransfer.Text.Trim(), txtDetailTransfer.Text.Trim(), txtPCNTransfer.Text.Trim(), txtSeriesTransfer.Text.Trim(), txtEventTransfer.Text.Trim(), txtTransfer.Text.Trim(), txtTransferPic.Text.Trim(), txtSNSetTransfer.Text.Trim(), txtPurposeTransfer.Text.Trim(), txtWhereTransfer.Text.Trim());
                            InsertHistory(txtSetNoTransfer.Text.Trim(), txtSectionTransfer.Text.Trim(), txtModelTransfer.Text.Trim(), txtDetailTransfer.Text.Trim(), txtPCNTransfer.Text.Trim(), txtSeriesTransfer.Text.Trim(), txtEventTransfer.Text.Trim(), txtSNSetTransfer.Text.Trim(), txtPurposeTransfer.Text.Trim(), txtWhereTransfer.Text.Trim(), txtTransfer.Text.Trim(), txtTransferPic.Text.Trim(), "", strAdminID, "Master -> Transfer");
                            getMaster();
                        }
                        finally
                        {
                            PleaseWait.Destroy();
                            MessageBox.Show("Transfer Completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                    }
                    else if (dialogResult == DialogResult.No)
                    {

                    }

                }
                else
                {
                    MessageBox.Show("Have no set information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btDispose_Click(object sender, EventArgs e)
        {
            if (txtSetNoDispose.Text.Trim() == "" | txtDispose.Text.Trim() == "")
            {
                MessageBox.Show("Please fill all information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                if (checkBorrow(txtSetNoDispose.Text.Trim()) == true)
                {
                    MessageBox.Show("This set was borrowed, please return before dispose!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (checkDup(txtSetNoDispose.Text.Trim()) == true)
                {
                    DialogResult dialogResult = MessageBox.Show("Dispose item: " + txtSetNoDispose.Text.Trim() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        PleaseWait.Create();
                        try
                        {
                            InsertDispose(txtSetNoDispose.Text.Trim(), txtSectionDispose.Text.Trim(), txtModelDispose.Text.Trim(), txtDetailDispose.Text.Trim(), txtPCNDispose.Text.Trim(), txtSeriesDispose.Text.Trim(), txtEventDispose.Text.Trim(), txtDispose.Text.Trim(), txtSNSetDispose.Text.Trim(), txtPurposeDispose.Text.Trim(), txtWhereDispose.Text.Trim());
                            InsertHistory(txtSetNoDispose.Text.Trim(), txtSectionDispose.Text.Trim(), txtModelDispose.Text.Trim(), txtDetailDispose.Text.Trim(), txtPCNDispose.Text.Trim(), txtSeriesDispose.Text.Trim(), txtEventDispose.Text.Trim(), txtSNSetDispose.Text.Trim(), txtPurposeDispose.Text.Trim(), txtWhereDispose.Text.Trim(), "", "", txtDispose.Text.Trim(), strAdminID, "Master -> Dispose");
                            getMaster();
                        }
                        finally
                        {
                            PleaseWait.Destroy();
                            MessageBox.Show("Dispose Completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                    }
                    else if (dialogResult == DialogResult.No)
                    {

                    }

                }
                else
                {
                    MessageBox.Show("Have no set information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private Boolean checkBorrow(string strSetNo)
        {
            string strSQL = "select * from set_data_update where flg_status=1 and SET_NO = '" + strSetNo + "'";

            Connection();
            MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
            DataTable dt = new DataTable();

            da.Fill(dt);
            foreach (DataRow drData in dt.Rows)
            {
                return true;

            }
            closeCon();

            return false;
        }



        public void searchEdit(string Editsearch)
        {
            Connection();
            if (comboBoxEdit.Text == "SET_NO")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SET_NO LIKE '%" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else if (comboBoxEdit.Text == "SECTION")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SECTION LIKE '%" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else if (comboBoxEdit.Text == "MODEL")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE MODEL LIKE '%" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else if (comboBoxEdit.Text == "DETAIL")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE DETAIL LIKE '%" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else if (comboBoxEdit.Text == "PCN_NO")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE PCN_NO LIKE '%" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else if (comboBoxEdit.Text == "SERIES")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SERIES LIKE '%" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else if (comboBoxEdit.Text == "EVENT")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE EVENT LIKE '" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else if (comboBoxEdit.Text == "SN_SET")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SN_SET LIKE '%" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else if (comboBoxEdit.Text == "PURPOSE")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE PURPOSE LIKE '%" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else if (comboBoxEdit.Text == "COL_WHERE")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE COL_WHERE LIKE '%" + txtSearchEdit.Text + "%'", conn);
                DataTable te = new DataTable();
                adap.Fill(te);
                dtGridViwerHitory.DataSource = te;
            }
            else
            {
                MessageBox.Show("Please Select Column name !");
            }
        }

        public void searchAddNew(string AddNewsearch)
        {
            Connection();
            if (comboBoxAdd_New.Text == "SET_NO")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SET_NO LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else if (comboBoxAdd_New.Text == "SECTION")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SECTION LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else if (comboBoxAdd_New.Text == "MODEL")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE MODEL LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else if (comboBoxAdd_New.Text == "DETAIL")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE DETAIL LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else if (comboBoxAdd_New.Text == "PCN_NO")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE PCN_NO LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else if (comboBoxAdd_New.Text == "SERIES")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SERIES LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else if (comboBoxAdd_New.Text == "EVENT")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE EVENT LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else if (comboBoxAdd_New.Text == "SN_SET")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SN_SET LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else if (comboBoxAdd_New.Text == "PURPOSE")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE PURPOSE LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else if (comboBoxAdd_New.Text == "COL_WHERE")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE COL_WHERE LIKE '%" + txtSearchAddNew.Text + "%'", conn);
                DataTable ta = new DataTable();
                adap.Fill(ta);
                dtGridViwerHitory.DataSource = ta;
            }
            else
            {
                MessageBox.Show("Please Select Column name !");
            }
        }

        public void searchTransfer(string Transfersearch)
        {
            if (comboBoxTransfer.Text == "SET_NO")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SET_NO LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else if (comboBoxTransfer.Text == "SECTION")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SECTION LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else if (comboBoxTransfer.Text == "MODEL")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE MODEL LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else if (comboBoxTransfer.Text == "DETAIL")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE DETAIL LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else if (comboBoxTransfer.Text == "PCN_NO")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE PCN_NO LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else if (comboBoxTransfer.Text == "SERIES")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SERIES LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else if (comboBoxTransfer.Text == "EVENT")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE EVENT LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else if (comboBoxTransfer.Text == "SN_SET")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SN_SET LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else if (comboBoxTransfer.Text == "PURPOSE")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE PURPOSE LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else if (comboBoxTransfer.Text == "COL_WHERE")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE COL_WHERE LIKE '%" + txtSearchTransfer.Text + "%'", conn);
                DataTable tt = new DataTable();
                adap.Fill(tt);
                dtGridViwerHitory.DataSource = tt;
            }
            else
            {
                MessageBox.Show("Please Select Column name !");
            }
        }

        public void searchDispose(string Disposesearch)
        {
            if (comboBoxDispose.Text == "SET_NO")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SET_NO LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else if (comboBoxDispose.Text == "SECTION")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SECTION LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else if (comboBoxDispose.Text == "MODEL")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE MODEL LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else if (comboBoxDispose.Text == "DETAIL")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE DETAIL LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else if (comboBoxDispose.Text == "PCN_NO")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE PCN_NO LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else if (comboBoxDispose.Text == "SERIES")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SERIES LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else if (comboBoxDispose.Text == "EVENT")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE EVENT LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else if (comboBoxDispose.Text == "SN_SET")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE SN_SET LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else if (comboBoxDispose.Text == "PURPOSE")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE PURPOSE LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else if (comboBoxDispose.Text == "COL_WHERE")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set WHERE COL_WHERE LIKE '%" + txtSearchDispose.Text + "%'", conn);
                DataTable td = new DataTable();
                adap.Fill(td);
                dtGridViwerHitory.DataSource = td;
            }
            else
            {
                MessageBox.Show("Please Select Column name !");
            }
        }

        private void btnSearchEdit_Click(object sender, EventArgs e)
        {
            if (txtSearchEdit.Text.Trim() == "")
            {
                MessageBox.Show("Please Fill Text Search!");
            }
            else
            {
                PleaseWait.Create();
                try
                {
                    string Editsearch = txtSearchEdit.Text.ToString();
                    searchEdit(Editsearch);
                }
                finally
                {
                    PleaseWait.Destroy();
                }
            }
        }

        private void btn_SearchAddNew_Click(object sender, EventArgs e)
        {
            if (txtSearchAddNew.Text.Trim() == "")
            {
                MessageBox.Show("Please Fill Text Search!");
            }
            else
            {
                PleaseWait.Create();
                try
                {
                    string AddNewsearch = txtSearchAddNew.Text.ToString();
                    searchAddNew(AddNewsearch);
                }
                finally
                {
                    PleaseWait.Destroy();
                }
            }
        }

        private void btn_SearchTransfer_Click(object sender, EventArgs e)
        {
            if (txtSearchTransfer.Text.Trim() == "")
            {
                MessageBox.Show("Please Fill Text Search!");
            }
            else
            {
                PleaseWait.Create();
                try
                {
                    string Transfersearch = txtSearchTransfer.Text.ToString();
                    searchTransfer(Transfersearch);
                }
                finally
                {
                    PleaseWait.Destroy();
                }
            }
        }

        private void btn_SearchDispose_Click(object sender, EventArgs e)
        {
            if (txtSearchDispose.Text.Trim() == "")
            {
                MessageBox.Show("Please Fill Text Search!");
            }
            else
            {
                PleaseWait.Create();
                try
                {
                    string Disposesearch = txtSearchDispose.Text.ToString();
                    searchDispose(Disposesearch);
                }
                finally
                {
                    PleaseWait.Destroy();
                }
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster();
                comboBoxAdd_New.SelectedIndex = comboBoxAdd_New.FindStringExact("");
                txtSearchAddNew.Clear();
                txtSetNo.Clear();
                txtPCN.Clear();
                txtPurpose.Clear();
                txtSection.Clear();
                txtSeries.Clear();
                txtWhere.Clear();
                txtModel.Clear();
                txtEvent.Clear();
                txtSNSet.Clear();
                txtDetail.Clear();
            }
            finally
            {
                PleaseWait.Destroy();
            }

        }

        private void btnRefreshEdit_Click(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster();
                comboBoxEdit.SelectedIndex = comboBoxEdit.FindStringExact("");
                txtSearchEdit.Clear();
                txtSetNoEdit.Clear();
                txtPCNEdit.Clear();
                txtPurposeEdit.Clear();
                txtSectionEdit.Clear();
                txtSeriesEdit.Clear();
                txtWhereEdit.Clear();
                txtModelEdit.Clear();
                txtEventEdit.Clear();
                txtSNSetEdit.Clear();
                txtDetailEdit.Clear();
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void btnRefreshTransfer_Click(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster();
                comboBoxTransfer.SelectedIndex = comboBoxTransfer.FindStringExact("");
                txtSearchTransfer.Clear();
                txtSetNoTransfer.Clear();
                txtPCNTransfer.Clear();
                txtPurposeTransfer.Clear();
                txtSectionTransfer.Clear();
                txtSeriesTransfer.Clear();
                txtWhereTransfer.Clear();
                txtModelTransfer.Clear();
                txtEventTransfer.Clear();
                txtSNSetTransfer.Clear();
                txtTransfer.Clear();
                txtTransferPic.Clear();
                txtDetailTransfer.Clear();
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }


        private void btnRefreshDispose_Click(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster();
                comboBoxDispose.SelectedIndex = comboBoxDispose.FindStringExact("");
                txtSearchTransfer.Clear();
                txtSetNoDispose.Clear();
                txtPCNDispose.Clear();
                txtPurposeDispose.Clear();
                txtSectionDispose.Clear();
                txtSeriesDispose.Clear();
                txtWhereDispose.Clear();
                txtModelDispose.Clear();
                txtEventDispose.Clear();
                txtSNSetDispose.Clear();
                txtDispose.Clear();
                txtDetailDispose.Clear();
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtTransfer.Text.Trim() == "" && txtTransferPic.Text.Trim() == "" && txtSetNoTransfer.Text.Trim() == "")
            {
                MessageBox.Show("Please fill all information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                var selectedRows = dtGridViwerHitory.SelectedRows
               .OfType<DataGridViewRow>()
               .Where(row => !row.IsNewRow)
               .ToArray();
                DataGridView dgv = sender as DataGridView;

                DialogResult result = MessageBox.Show("Do you want to Transfer Data?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    foreach (var rows in selectedRows)
                    {

                        if (rows != null)
                        {
                            string strSetNo = Convert.ToString(rows.Cells["SET_NO"].Value.ToString());
                            string strSection = Convert.ToString(rows.Cells["SECTION"].Value.ToString());
                            string strModel = Convert.ToString(rows.Cells["MODEL"].Value.ToString());
                            string strDetail = Convert.ToString(rows.Cells["DETAIL"].Value.ToString());
                            string strPCN = Convert.ToString(rows.Cells["PCN_NO"].Value.ToString());
                            string strSeries = Convert.ToString(rows.Cells["SERIES"].Value.ToString());
                            string strEvent = Convert.ToString(rows.Cells["EVENT"].Value.ToString());
                            string strTransfer = txtTransfer.Text.Trim();
                            string strTransferPic = txtTransferPic.Text.Trim();
                            string strSNSet = Convert.ToString(rows.Cells["SN_SET"].Value.ToString());
                            string strPurpose = Convert.ToString(rows.Cells["PURPOSE"].Value.ToString());
                            string strWhere = Convert.ToString(rows.Cells["COL_WHERE"].Value.ToString());

                            if (checkBorrow(strSetNo) == true)
                            {
                                MessageBox.Show("Set No : " + strSetNo + " has been borrowed, please return before transfer!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else
                            {
                                if (checkDup(strSetNo) == true)
                                {
                                    PleaseWait.Create();
                                    try
                                    {
                                        InsertTransfer(strSetNo, strSection, strModel, strDetail, strPCN, strSeries, strEvent, strTransfer, strTransferPic, strSNSet, strPurpose, strWhere);
                                        //InsertTransfer(txtSetNoTransfer.Text.Trim(), txtSectionTransfer.Text.Trim(), txtModelTransfer.Text.Trim(), txtDetailTransfer.Text.Trim(), txtPCNTransfer.Text.Trim(), txtSeriesTransfer.Text.Trim(), txtEventTransfer.Text.Trim(), txtTransfer.Text.Trim(), txtTransferPic.Text.Trim(), txtSNSetTransfer.Text.Trim(), txtPurposeTransfer.Text.Trim(), txtWhereTransfer.Text.Trim());
                                        InsertHistory(strSetNo, strSection, strModel, strDetail, strPCN, strSeries, strEvent, strSNSet, strPurpose, strWhere, strTransfer, strTransferPic, "", strAdminID, "Master -> Transfer");


                                        getMaster();
                                        txtSearchTransfer.Clear();
                                        txtSetNoTransfer.Clear();
                                        txtPCNTransfer.Clear();
                                        txtPurposeTransfer.Clear();
                                        txtSectionTransfer.Clear();
                                        txtSeriesTransfer.Clear();
                                        txtWhereTransfer.Clear();
                                        txtModelTransfer.Clear();
                                        txtEventTransfer.Clear();
                                        txtSNSetTransfer.Clear();
                                        txtTransfer.Clear();
                                        txtTransferPic.Clear();
                                        txtDetailTransfer.Clear();

                                    }
                                    finally
                                    {
                                        PleaseWait.Destroy();
                                        MessageBox.Show("Transfer Completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }


                                }
                                else
                                {
                                    MessageBox.Show("Have no set information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                            }

                        }

                    }


                }
                else
                {

                }

            }

        }

        private void btnDispose_Click(object sender, EventArgs e)
        {
            if (txtDispose.Text.Trim() == "" && txtSetNoDispose.Text.Trim() == "")
            {
                MessageBox.Show("Please fill Disposal!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                var selectedRows = dtGridViwerHitory.SelectedRows
               .OfType<DataGridViewRow>()
               .Where(row => !row.IsNewRow)
               .ToArray();
                DataGridView dgv = sender as DataGridView;

                DialogResult result = MessageBox.Show("Do you want to Dispose Data?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    foreach (var rows in selectedRows)
                    {
                        if (rows != null)
                        {
                            string strSetNo = Convert.ToString(rows.Cells["SET_NO"].Value.ToString());
                            string strSection = Convert.ToString(rows.Cells["SECTION"].Value.ToString());
                            string strModel = Convert.ToString(rows.Cells["MODEL"].Value.ToString());
                            string strDetail = Convert.ToString(rows.Cells["DETAIL"].Value.ToString());
                            string strPCN = Convert.ToString(rows.Cells["PCN_NO"].Value.ToString());
                            string strSeries = Convert.ToString(rows.Cells["SERIES"].Value.ToString());
                            string strEvent = Convert.ToString(rows.Cells["EVENT"].Value.ToString());
                            string strDispose = txtDispose.Text.Trim();
                            string strSNSet = Convert.ToString(rows.Cells["SN_SET"].Value.ToString());
                            string strPurpose = Convert.ToString(rows.Cells["PURPOSE"].Value.ToString());
                            string strWhere = Convert.ToString(rows.Cells["COL_WHERE"].Value.ToString());

                            if (checkBorrow(strSetNo) == true)
                            {
                                MessageBox.Show("Set No : " + strSetNo + " has been borrowed, please return before dispose!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            else
                            {
                                if (checkDup(strSetNo) == true)
                                {
                                    PleaseWait.Create();
                                    try
                                    {
                                        InsertDispose(strSetNo, strSection, strModel, strDetail, strPCN, strSeries, strEvent, strDispose, strSNSet, strPurpose, strWhere);
                                        InsertHistory(strSetNo, strSection, strModel, strDetail, strPCN, strSeries, strEvent, strSNSet, strPurpose, strWhere, "", "", txtDispose.Text.Trim(), strAdminID, "Master -> Dispose");
                                        //InsertHistory(txtSetNoDispose.Text.Trim(), txtSectionDispose.Text.Trim(), txtModelDispose.Text.Trim(), txtDetailDispose.Text.Trim(), txtPCNDispose.Text.Trim(), txtSeriesDispose.Text.Trim(), txtEventDispose.Text.Trim(), txtSNSetDispose.Text.Trim(), txtPurposeDispose.Text.Trim(), txtWhereDispose.Text.Trim(), "", "", txtDispose.Text.Trim(), strAdminID, "Master -> Dispose");

                                        getMaster();
                                        txtSearchTransfer.Clear();
                                        txtSetNoDispose.Clear();
                                        txtPCNDispose.Clear();
                                        txtPurposeDispose.Clear();
                                        txtSectionDispose.Clear();
                                        txtSeriesDispose.Clear();
                                        txtWhereDispose.Clear();
                                        txtModelDispose.Clear();
                                        txtEventDispose.Clear();
                                        txtSNSetDispose.Clear();
                                        txtDispose.Clear();
                                        txtDetailDispose.Clear();
                                    }
                                    finally
                                    {
                                        PleaseWait.Destroy();
                                        MessageBox.Show("Dispose Completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }
   
                                }
                                else
                                {
                                    MessageBox.Show("Have no set information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                            }


                        }

                    }
           
                }
                else
                {
                }
            }


        }

        private void dtGridViwerHitory_SelectionChanged(object sender, EventArgs e)
        {
            txtTotalTransfer.Text = dtGridViwerHitory.SelectedRows.Count.ToString();
            txtTotalDispose.Text = dtGridViwerHitory.SelectedRows.Count.ToString();
        }


    }

}


