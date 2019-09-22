using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using INI;


namespace MASTER_SET_CONTROL
{
    public partial class frmMember : Form
    {
        MySqlConnection conn;
        IniFile iniconfig;
        public string select_db;
        public string strCon;

        BindingSource bs = new BindingSource();
        BindingList<DataTable> tables = new BindingList<DataTable>();

        public frmMember()
        {
            InitializeComponent();
        }

        private void frmMember_Load(object sender, EventArgs e)
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            if (select_db == "SAMPLE SET CONTROL")
            {
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB2 = iniconfig.IniReadValue("MySQL_Server", "DB");

                strCon = "Server=" + IP + ";";
                strCon += "Uid=root;";
                strCon += "Password=123456*;";
                strCon += "Database=" + DB2 + ";";

                conn = new MySqlConnection(strCon);

                conn.Open();

                this.Text = "SAMPLE SET CONTROL FORM MEMBER" + String.Format(" --- Version {0}", version) + " - Server : " + IP;

                PleaseWait.Create();
                try
                {
                    getMember();
                }
                finally
                {
                    PleaseWait.Destroy();
                }
            }
            else if (select_db == "ENGINEERING TRAINING CENTER")
            {
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB2 = iniconfig.IniReadValue("MySQL_Server", "DB2");

                strCon = "Server=" + IP + ";";
                strCon += "Uid=root;";
                strCon += "Password=123456*;";
                strCon += "Database=" + DB2 + ";";

                conn = new MySqlConnection(strCon);

                conn.Open();
                this.Text = "ENGINEERING TRAINING CENTER FORM MEMBER" + String.Format(" --- Version {0}", version) + " - Server : "+IP;

                PleaseWait.Create();
                try
                {
                    getMember();
                }
                finally
                {
                    PleaseWait.Destroy();
                }
            }
            else
            {
                MessageBox.Show("Not Connect to server!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            comboSearch.Items.Add("personal_id");
            comboSearch.Items.Add("personal_name");
            comboSearch.Items.Add("personal_email");
            comboSearch.Items.Add("Department");
            comboGroupCode.Items.Add("B1");
            getDept();

        }

        public void getDept()
        {
            Connection();
            MySqlDataAdapter adap = new MySqlDataAdapter("select dept from depthead_master", conn);
            DataTable table = new DataTable();
            adap.Fill(table);

            foreach (DataRow row in table.Rows)
            {
                comboDept.Items.Add(row["dept"]);
            }

            closeCon();
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

        public void closeCon()
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }

        }


        public void getHistoryMember()
        {
            Connection();
            MySqlDataAdapter adapter = new MySqlDataAdapter("select * from v_dept_email_group where dept = '" + comboDept.Text + "'", conn);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboEiaAdmin.Items.Clear();
            txtEIA_admin.Text = "";
            comboDeptHead.Items.Clear();

            foreach (DataRow row in dt.Rows)
            {
                comboDeptHead.Items.Add(row["dept_head"]);

            }
            closeCon();

        }

        public void getDeptHead()
        {
            MySqlDataAdapter adapter = new MySqlDataAdapter("select * from v_dept_email_group where dept = '" + comboDept.SelectedItem.ToString() + "'", conn);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            txtEIA_admin.Text = "";

            foreach (DataRow row in dt.Rows)
            {
                string dh = Convert.ToString(row["dept_head"]);
                if (comboDeptHead.Text == dh)
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select * from v_dept_email_group where dept_head = '" + comboDeptHead.SelectedItem.ToString() + "' AND dept = '" + comboDept.SelectedItem.ToString() + "'", conn);
                    DataTable data = new DataTable();
                    adap.Fill(data);
                    comboEiaAdmin.Items.Clear();

                    foreach (DataRow rw in data.Rows)
                    {
                        comboEiaAdmin.Items.Add(rw["eia_admin"]);
                    }
                }

            }

        }

        public void getEIA()
        {
            MySqlDataAdapter adapter = new MySqlDataAdapter("select * from v_dept_email_group where dept_head = '" + comboDeptHead.SelectedItem.ToString() + "'", conn);
            DataTable dt = new DataTable();
            adapter.Fill(dt);

            foreach (DataRow row in dt.Rows)
            {
                string EA = Convert.ToString(row["eia_admin"]);
                if (comboEiaAdmin.Text == EA)
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select * from v_dept_email_group where eia_admin = '" + comboEiaAdmin.SelectedItem.ToString() + "'", conn);
                    DataTable data = new DataTable();
                    adapter.Fill(data);

                    foreach (DataRow rw in data.Rows)
                    {
                        string eiaAdmin = Convert.ToString(rw["eia_admin"]);
                        txtEIA_admin.Text = eiaAdmin;
                    }
                }
            }
        }

        DataTable t = new DataTable();
        private void getMember()
        {

            dtGridViwerHitory.DataSource = null;
            dtGridViwerHitory.Rows.Clear();
            dtGridViwerHitory.Refresh();
            bindingNavigator.BindingSource = null;
            tables.Clear();
            t.Clear();

            string strSQL = "";

            strSQL = "SELECT personal_master.*,personal_email.dept,personal_email.dept_head,personal_email.eia_admin FROM personal_master left join personal_email on personal_master.personal_id = personal_email.personal_id";

            Connection();
            {
                using (MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn))
                {
                    da.Fill(t);
                    dtGridViwerHitory.DataSource = t;

                    if (t.Rows.Count != 0)
                    {
                        int count = 0;
                        DataTable dt = new DataTable();
                        dt = null;

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
                            dtGridViwerHitory.DataSource = null;

                        }

                        bindingNavigator.BindingSource = bs;
                        bs.DataSource = tables;
                        bs.PositionChanged += bs_PositionChanged;
                        bs_PositionChanged(bs, EventArgs.Empty);

                    }

                }
            }

        }

        void bs_PositionChanged(object sender, EventArgs e)
        {
            if (tables.Count != 0)
            {
                dtGridViwerHitory.DataSource = tables[bs.Position];
            }
        }

        private void getDeptHeadMaster()
        {
            Connection();
            MySqlDataAdapter adapter = new MySqlDataAdapter("select * from depthead_master", conn);
            DataTable tb = new DataTable();
            adapter.Fill(tb);
            dtGridViwerHitory.DataSource = tb;
            closeCon();
        }

        string strDept;
        private void dtGridViwerHitory_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                if (dtGridViwerHitory.SelectedCells.Count > 0 & e.ColumnIndex >= 0 & e.RowIndex >= 0)
                {

                    dtGridViwerHitory.CurrentCell = dtGridViwerHitory.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    //Can leave these here - doesn't hurt
                    dtGridViwerHitory.Rows[e.RowIndex].Selected = true;
                    dtGridViwerHitory.Focus();

                    int selectedrowindex = dtGridViwerHitory.SelectedCells[0].RowIndex;
                    DataGridViewRow selectedRow = dtGridViwerHitory.Rows[selectedrowindex];

                    txtID.Text = Convert.ToString(selectedRow.Cells[1].Value.ToString());
                    txtName.Text = Convert.ToString(selectedRow.Cells[2].Value.ToString());
                    txtEmail.Text = Convert.ToString(selectedRow.Cells[3].Value.ToString());
                    txtGroupCode.Text = Convert.ToString(selectedRow.Cells[4].Value.ToString());

                    comboDept.Text = Convert.ToString(selectedRow.Cells[5].Value.ToString());
                    comboDeptHead.Text = Convert.ToString(selectedRow.Cells[6].Value.ToString());
                    comboEiaAdmin.Text = Convert.ToString(selectedRow.Cells[7].Value.ToString());
                    txtEIA_admin.Text = Convert.ToString(selectedRow.Cells[7].Value.ToString());

                    if (Convert.ToString(selectedRow.Cells[4].Value.ToString()) == "A1")
                    {
                        txtID.Enabled = false;
                        txtName.Enabled = false;
                        txtEmail.Enabled = false;
                        txtGroupCode.Enabled = false;

                        txtID.BackColor = Color.Silver;
                        txtName.BackColor = Color.Silver;
                        txtEmail.BackColor = Color.Silver;
                        txtGroupCode.BackColor = Color.Silver;

                        comboDept.Enabled = false;
                        comboDeptHead.Enabled = false;
                        comboEiaAdmin.Enabled = false;
                        txtEIA_admin.Enabled = false;

                        btAdd.Enabled = false;
                        btDelete.Enabled = false;
                        btnEdit.Enabled = false;
                        btnRefresh.Enabled = false;
                    }
                    else
                    {
                        txtID.Enabled = true;
                        txtName.Enabled = true;
                        txtEmail.Enabled = true;
                        txtGroupCode.Enabled = false;

                        txtID.BackColor = Color.Cyan;
                        txtName.BackColor = Color.Cyan;
                        txtEmail.BackColor = Color.Cyan;
                        txtGroupCode.BackColor = Color.Silver;

                        comboDept.Enabled = true;
                        comboDeptHead.Enabled = true;
                        comboEiaAdmin.Enabled = true;
                        txtEIA_admin.Enabled = true;

                        btAdd.Enabled = true;
                        btDelete.Enabled = true;
                        btnEdit.Enabled = true;
                        btnRefresh.Enabled = true;
                    }

                }

            }
            else if (tabControl1.SelectedIndex == 1)
            {
                if (dtGridViwerHitory.SelectedCells.Count > 0 & e.ColumnIndex >= 0 & e.RowIndex >= 0)
                {

                    dtGridViwerHitory.CurrentCell = dtGridViwerHitory.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    //Can leave these here - doesn't hurt
                    dtGridViwerHitory.Rows[e.RowIndex].Selected = true;
                    dtGridViwerHitory.Focus();

                    int selectedrowindex = dtGridViwerHitory.SelectedCells[0].RowIndex;
                    DataGridViewRow selectedRow = dtGridViwerHitory.Rows[selectedrowindex];

                    txtDept.Text = Convert.ToString(selectedRow.Cells[1].Value.ToString());
                    strDept = Convert.ToString(selectedRow.Cells[1].Value.ToString());
                    txtDeptHead.Text = Convert.ToString(selectedRow.Cells[2].Value.ToString());



                }

            }

        }

        public void UpdatePersonalEmail()
        {
            Connection();
            string command = "";

            command = "update personal_email set dept = '" + txtDept.Text + "', dept_head = '" + txtDeptHead.Text + "' where dept = '" + strDept + "'  ";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }

        public void UpdateDeptHead()
        {
            Connection();
            string command = "";

            command = "update depthead_master set dept = '" + txtDept.Text + "', dept_head = '" + txtDeptHead.Text + "' where dept = '" + strDept + "'  ";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }



        public void InsertMember(string strID, string strName, string strEmail)
        {

            Connection();
            string command = "";

            command = "insert into personal_master (`personal_id`,`personal_name`,`personal_email`,`GroupCode`)values('" +
                strID + "','" + strName + "','" + strEmail + "','B1')";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }


        public void InsertPersonalEmail(string strID, string strName, string strEmail, string strDept, string strDeptHead, string strEIAadmin)
        {
            Connection();
            string command = "";

            command = "insert into personal_email (personal_id,personal_name,personal_email,GroupCode,dept,dept_head,eia_admin)values('" +
                strID + "','" + strName + "','" + strEmail + "','B1', '" + strDept + "', '" + strDeptHead + "', '" + strEIAadmin + "')";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }

        public void deleteMember(string strID)
        {

            Connection();
            string command = "";

            command = "delete from personal_master where `personal_id`='" + strID + "'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }

        public void deletePersonalEmail(string strID)
        {
            Connection();
            string command = "";

            command = "delete from personal_email where `personal_id`='" + strID + "'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }


        public void EditMember(string strName, string strEmail)
        {
            Connection();
            string command = "";

            command = "update personal_master set personal_name = '" + strName + "',personal_email='" + strEmail + "', GroupCode='B1' where personal_id = '" + txtID.Text.Trim() + "'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }

        public void EditPersonalEmail(string strName, string strEmail, string strDept, string strDeptHead, string strEia)
        {
            Connection();
            string command = "";

            command = "update personal_email set personal_name = '" + strName + "',personal_email='" + strEmail + "',dept = '" + strDept + "',dept_head = '" + strDeptHead + "',eia_admin = '" + strEia + "', GroupCode='B1' where personal_id = '" + txtID.Text.Trim() + "'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }

        public void InsertDeptheadMaster(string strDept, string strDeptHead)
        {
            Connection();
            string command = "";

            command = "insert into depthead_master(dept,dept_head) values('" + strDept + "','" + strDeptHead + "')";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }

        private Boolean checkMember(string strID)
        {
            string strSQL = "select * from personal_master where personal_id = '" + strID + "'";

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

        private Boolean checkAdmin(string strID)
        {
            string strSQL = "select * from personal_master where personal_id = '" + strID + "' and GroupCode = 'A1'";

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

        private Boolean checkDept(string strDept, string strDeptHead)
        {
            string strSQL = "select dept from depthead_master where dept = '" + strDept + "' and dept_head = '" + strDeptHead + "' ";

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


        private void btAdd_Click(object sender, EventArgs e)
        {
            if (txtID.Text.Trim() == "" | txtName.Text.Trim() == "" | txtEmail.Text.Trim() == "")
            {
                MessageBox.Show("Please fill all information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (checkMember(txtID.Text.Trim()) == true)
            {
                MessageBox.Show("This member was registeraton already!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            else
            {
                DialogResult dialogResult = MessageBox.Show("Register new member: " + txtID.Text.Trim() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    if (checkDept(comboDept.Text, comboDeptHead.Text) == true)
                    {
                        PleaseWait.Create();
                        try
                        {
                            InsertMember(txtID.Text.Trim(), txtName.Text.Trim(), txtEmail.Text.Trim());
                            InsertPersonalEmail(txtID.Text.Trim(), txtName.Text.Trim(), txtEmail.Text.Trim(), comboDept.Text, comboDeptHead.Text, txtEIA_admin.Text.Trim());
                            getMember();

                        }
                        finally
                        {
                            PleaseWait.Destroy();
                            MessageBox.Show("Registeration Completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Department and Department Head Not Match.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
        }

        private void btDelete_Click(object sender, EventArgs e)
        {
            if (checkAdmin(txtID.Text.Trim()) == true)
            {
                MessageBox.Show("You can not delete admin!", "Access Denide", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (checkMember(txtID.Text.Trim()) == false)
            {
                MessageBox.Show("Have no member information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {

                DialogResult dialogResult = MessageBox.Show("Delete member: " + txtID.Text.Trim() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    PleaseWait.Create();
                    try
                    {
                        deleteMember(txtID.Text.Trim());
                        deletePersonalEmail(txtID.Text.Trim());
                        getMember();
                    }
                    finally
                    {
                        PleaseWait.Destroy();
                        MessageBox.Show("Delete Completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void comboDept_SelectedIndexChanged(object sender, EventArgs e)
        {
            getHistoryMember();
        }

        private void comboDeptHead_SelectedIndexChanged(object sender, EventArgs e)
        {
            getDeptHead();
        }

        private void comboEiaAdmin_SelectedIndexChanged(object sender, EventArgs e)
        {
            getEIA();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (checkAdmin(txtID.Text.Trim()) == true)
            {
                MessageBox.Show("You can not Edit admin!", "Access Denide", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (checkMember(txtID.Text.Trim()) == true)
            {
                DialogResult dialogResult = MessageBox.Show("Edit member: " + txtID.Text.Trim() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    if (checkDept(comboDept.Text, comboDeptHead.Text) == true)
                    {
                        PleaseWait.Create();
                        try
                        {
                            EditMember(txtName.Text.Trim(), txtEmail.Text.Trim());
                            EditPersonalEmail(txtName.Text.Trim(), txtEmail.Text.Trim(), comboDept.Text, comboDeptHead.Text, txtEIA_admin.Text.Trim());
                            getMember();
                        }
                        finally
                        {
                            PleaseWait.Destroy();
                            MessageBox.Show("Edit Completed!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Department and Department Head Not Match.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }

            }
            else
            {
                MessageBox.Show("Have no member information!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            txtID.Text = "";
            txtName.Text = "";
            txtEmail.Text = "";
            txtGroupCode.Text = "";

            comboDeptHead.Items.Clear();
            comboDept.Text = "";
            comboDeptHead.Text = "";
            comboEiaAdmin.Items.Clear();
            txtEIA_admin.Text = "";
            txtSearch.Text = "";
            comboSearch.SelectedItem = "";
            getMember();

        }

        private void comboGroupCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectGroupCode();
        }

        public void selectGroupCode()
        {
            if ((string)comboGroupCode.SelectedItem == "B1")
            {
                txtGroupCode.Text = comboGroupCode.SelectedItem.ToString();
            }
        }

        private void tabControl1_Click(object sender, EventArgs e)
        {
            comboDept.Text = "";
            comboDept.Items.Clear();
            comboDeptHead.Text = "";
            comboEiaAdmin.Items.Clear();
            txtEIA_admin.Text = "";
            txtDept.Text = "";
            txtDeptHead.Text = "";
            txtID.Text = "";
            txtName.Text = "";
            txtEmail.Text = "";
            txtGroupCode.Text = "";

            checktab();

        }

        public void checktab()
        {
            if (tabControl1.SelectedIndex == 0)
            {
                PleaseWait.Create();
                try
                {
                    getMember();
                    getDept();
                    txtID.Enabled = true;
                    txtName.Enabled = true;
                    txtEmail.Enabled = true;
                    txtGroupCode.Enabled = false;

                    txtID.BackColor = Color.Cyan;
                    txtName.BackColor = Color.Cyan;
                    txtEmail.BackColor = Color.Cyan;
                    txtGroupCode.BackColor = Color.Silver;

                }
                finally
                {
                    PleaseWait.Destroy();
                }

            }
            else if (tabControl1.SelectedIndex == 1)
            {

                PleaseWait.Create();
                try
                {
                    getDeptHeadMaster();
                    txtID.Enabled = false;
                    txtName.Enabled = false;
                    txtEmail.Enabled = false;
                    txtGroupCode.Enabled = false;

                    txtID.BackColor = Color.Silver;
                    txtName.BackColor = Color.Silver;
                    txtEmail.BackColor = Color.Silver;
                    txtGroupCode.BackColor = Color.Silver;
                }
                finally
                {
                    PleaseWait.Destroy();
                }

            }
        }


        private void btn_EditDeptHead_Click(object sender, EventArgs e)
        {
            if (txtDept.Text != "" && txtDeptHead.Text != "")
            {
                PleaseWait.Create();
                try
                {
                    UpdatePersonalEmail();
                    UpdateDeptHead();
                    getDeptHeadMaster();
                }
                finally
                {
                    PleaseWait.Destroy();
                }
                MessageBox.Show("Edit Complete.", "Confirmation!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else
            {
                MessageBox.Show("Please Select Cell Edit", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void btnDeptHead_Click(object sender, EventArgs e)
        {
            if (checkDept(txtDept.Text.Trim(), txtDeptHead.Text.Trim()) == true)
            {
                MessageBox.Show("Department and Department Head Duplicate.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (txtDept.Text.Trim() != "" && txtDeptHead.Text.Trim() != "")
                {
                    PleaseWait.Create();
                    try
                    {
                        InsertDeptheadMaster(txtDept.Text.Trim(), txtDeptHead.Text.Trim());
                        getDeptHeadMaster();

                    }
                    finally
                    {
                        PleaseWait.Destroy();
                    }

                }
                else
                {
                    MessageBox.Show("Please Fill Department and Department Head.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        public void searchMember(string type)
        {
            Connection();
            if ((string)comboSearch.SelectedItem == "personal_id")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM personal_email WHERE personal_id LIKE '%" + type + "%'", conn);
                DataTable tb = new DataTable();
                adap.Fill(tb);
                dtGridViwerHitory.DataSource = tb;
            }
            else if ((string)comboSearch.SelectedItem == "personal_name")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM personal_email WHERE personal_name LIKE '%" + type + "%'", conn);
                DataTable tb = new DataTable();
                adap.Fill(tb);
                dtGridViwerHitory.DataSource = tb;
            }
            else if ((string)comboSearch.SelectedItem == "personal_email")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM personal_email WHERE personal_email LIKE '%" + type + "%'", conn);
                DataTable tb = new DataTable();
                adap.Fill(tb);
                dtGridViwerHitory.DataSource = tb;
            }
            else if ((string)comboSearch.SelectedItem == "Department")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM personal_email WHERE dept LIKE '%" + type + "%'", conn);
                DataTable tb = new DataTable();
                adap.Fill(tb);
                dtGridViwerHitory.DataSource = tb;
            }
            closeCon();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtSearch.Text.Trim() == "")
            {
                MessageBox.Show("Please Fill Text Search!");
            }
            else
            {
                PleaseWait.Create();
                try
                {
                    searchMember(txtSearch.Text);
                }
                finally
                {
                    PleaseWait.Destroy();
                }

            }
        }

    }

}
