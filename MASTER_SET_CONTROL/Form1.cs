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
using Microsoft.VisualBasic;
using System.Data.SqlClient;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using System.IO.Ports;

using INI;
using STTADUser;

namespace MASTER_SET_CONTROL
{
    public partial class Form1 : Form
    {

        MySqlConnection conn;
        //EIA
        public string strCon;

        //IniFile Gen;
        BindingSource bs = new BindingSource();
        BindingList<DataTable> tables = new BindingList<DataTable>();

        DataTable dt2 = new DataTable();
        DataTable dtExportpdf = new DataTable();
        DataTable newdata = new DataTable();
        DataTable tt = new DataTable();
        DataTable TableNew = new DataTable();
        DataTable TableNew2 = new DataTable();
        //*************************************
        string strSetNo = "";
        string strSection = "";
        string strModel = "";
        string strDetail = "";
        string strPCN = "";
        string strSeries = "";
        string strEvent = "";
        string strSNSet = "";
        string strPurpose = "";
        string strWhere= "";
        public string Setvaluefortext;
        string cmd_multiborrow = "";
        string cmd_multiReturn = "";
        string cmd_multiborrowNotIssue = "";
        public int Counter;
        string number;
        public string com;
        
        //*************************************
        public string b;
        string Data;
        string get_Data;
        string textuser;

        IniFile iniconfig;
        SqlConnection SQLConnection;
        //*************************************
        //public string DocumentNo, BorrowDate, ENNo, NameSurname, Department;

        //*************************************
      
        public Form1()
        {
            InitializeComponent();
        }

        public void ab(string a)
        {
            b = a.ToString();
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            if (b == "SAMPLE SET CONTROL")
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
                this.Text = "SAMPLE SET CONTROL" + String.Format(" --- Version {0}", version) + " - Server : " + IP;
                if (b == "SAMPLE SET CONTROL")
                {
                    Data = b;
                    checkselect();
                }
            }
            else if (b == "ENGINEERING TRAINING CENTER")
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
                this.Text = "ENGINEERING TRAINING CENTER" + String.Format(" --- Version {0}", version) + " - Server : " + IP;
            }
            else
            {
                MessageBox.Show("Not connect to server");
            }
        }


        public void checkselect()
        {
            if (Data == "SAMPLE SET CONTROL")
            {
                get_Data = Data;
            }
            else if (Data == "ENGINEERING TRAINING CENTER")
            {
                get_Data = Data;
            }
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            radioNotUseRFID.ForeColor = Color.Yellow;

            TableNew.Columns.Add("SET_NO");
            TableNew.Columns.Add("MODEL");

            TableNew2.Columns.Add("SET_NO");
            TableNew2.Columns.Add("MODEL");

            dt2.Columns.Add("SET_NO");
            dt2.Columns.Add("SECTION");
            dt2.Columns.Add("MODEL");
            dt2.Columns.Add("DETAIL");
            dt2.Columns.Add("PCN_NO");
            dt2.Columns.Add("SERIES");
            dt2.Columns.Add("EVENT");
            dt2.Columns.Add("SN_SET");
            dt2.Columns.Add("PURPOSE");
            dt2.Columns.Add("COL_WHERE");
            dt2.Columns.Add("TRANSFER");
            dt2.Columns.Add("TRANSFER_PIC");
            dt2.Columns.Add("DISPOSAL");


            newdata.Columns.Add("unit");
            newdata.Columns.Add("DocumentNo");
            newdata.Columns.Add("ScanID");
            newdata.Columns.Add("Name");
            newdata.Columns.Add("Department");
            newdata.Columns.Add("Model");
            newdata.Columns.Add("EIA");
            newdata.Columns.Add("Count");
            newdata.Columns.Add("checkBox");
            newdata.Columns.Add("DateBorrow");
            newdata.Columns.Add("AUTO_ID");

            tab.Columns.Add("auto_id");
            tab.Columns.Add("unit");
            tab.Columns.Add("DocumentNoSampleMulti");
            tab.Columns.Add("strDuDate");
            tab.Columns.Add("ScanID");
            tab.Columns.Add("Name");
            tab.Columns.Add("Department");
            tab.Columns.Add("Model");
            tab.Columns.Add("EIA");
            tab.Columns.Add("Count");
            tab.Columns.Add("ID");
            tab.Columns.Add("checkBox");

            

            datePickerFrom.Value = DateTime.Today.AddDays(-7);
            timePickerFrom.Value = Convert.ToDateTime("08:00:00");
            timePickerTo.Value = Convert.ToDateTime("20:00:00");

            comboSearchType.SelectedIndex = 0;

            datePickerReturn.Value = DateTime.Today.AddDays(7);
            checkSpacial.Checked = false;
            txtSpacial.Text = "";

            radioBorrow.Checked = true;
            radioBorrow.ForeColor = Color.Yellow;
            radioReturn.Checked = false;
            radioReturn.ForeColor = Color.White;
            radioDefaultBorrow.Checked = true;
            radioDefaultBorrow.ForeColor = Color.Yellow;
            radioMultipleBorrow.Checked = false;
            radioMultipleBorrow.ForeColor = Color.White;

            lblRequestName.Text = "";

            txtScanID.Text = "";
            txtScanID.Enabled = true;
            txtScanID.Select();
            txtScanID.BackColor = Color.Yellow;

            txtScanSetNo.Text = "";
            txtScanSetNo.Enabled = false;
            txtScanSetNo.BackColor = Color.Gray;

            btnCancleBorrow.Visible = false;
            btnMultipleBorrow.Visible = false;

            PleaseWait.Create();
            try
            {
                getHistory("NORMAL BORROW");
                getCount();
            }
            finally
            {
                PleaseWait.Destroy();
            }

        }


        private void radioBorrow_Click(object sender, EventArgs e)
        {
            l = 0;
            n = 0;
            k = 0;
            getModel = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
            A = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
            B = new string[500];
            C = new int[501];
            D = new string[501];

            lblRequestName.Text = "";
            radioNotUseRFID.Checked = true;
            radioNotUseRFID.ForeColor = Color.Yellow;
            radioRFID.Checked = false;
            radioRFID.ForeColor = Color.White;

            radioIssueDocument.Checked = false;
            radioIssueDocument.ForeColor = Color.White;
            radioNotIssue.Checked = false;
            radioNotIssue.ForeColor = Color.White;
            groupBox8.Visible = true;
            groupBox6.Visible = true;
            checkBoxInternal.Checked = false;
            checkBoxExternal.Checked = false;

            radioBorrow.Checked = true;
            radioBorrow.ForeColor = Color.Yellow;
            radioReturn.Checked = false;
            radioReturn.ForeColor = Color.White;
            radioDefaultBorrow.Checked = true;
            radioDefaultBorrow.ForeColor = Color.Yellow;
            radioMultipleBorrow.Checked = false;
            radioMultipleBorrow.ForeColor = Color.White;

            btnCancleBorrow.Visible = false;
            btnMultipleBorrow.Visible = false;

            txtScanID.Text = "";
            txtScanID.Enabled = true;
            txtScanID.Focus();
            txtScanID.BackColor = Color.Yellow;

            txtScanSetNo.Text = "";
            txtScanSetNo.Enabled = false;
            txtScanSetNo.BackColor = Color.Gray;

            if (radioBorrow.Checked == true)
            {
                PleaseWait.Create();
                try
                {
                    getHistory(comboSearchType.Text);
                    getCount();
                }
                finally
                {
                    PleaseWait.Destroy();
                }
            }
        }

        private void radioReturn_Click(object sender, EventArgs e)
        {
            l = 0;
            n = 0;
            k = 0;
            getModel = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
            A = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
            B = new string[500];
            C = new int[501];
            D = new string[501];

            lblRequestName.Text = "";
            radioNotUseRFID.Checked = true;
            radioNotUseRFID.ForeColor = Color.Yellow;
            radioRFID.Checked = false;
            radioRFID.ForeColor = Color.White;

            radioBorrow.Checked = false;
            radioBorrow.ForeColor = Color.White;
            radioReturn.Checked = true;
            radioReturn.ForeColor = Color.Yellow;
            radioDefaultBorrow.Checked = true;
            radioDefaultBorrow.ForeColor = Color.Yellow;
            radioMultipleBorrow.Checked = false;
            radioMultipleBorrow.ForeColor = Color.White;
            btnCancleBorrow.Visible = false;
            btnMultipleBorrow.Visible = false;
            datePickerReturn.Value = DateTime.Today.AddDays(7);
            checkSpacial.Checked = false;
            txtSpacial.Text = "";

            txtScanID.Text = "";
            txtScanID.Focus();
            txtScanID.Enabled = true;
            txtScanID.BackColor = Color.Yellow;

            txtScanSetNo.Text = "";
            txtScanSetNo.Enabled = false;
            txtScanSetNo.BackColor = Color.Gray;

            if (radioReturn.Checked == true)
            {
                PleaseWait.Create();
                try
                {
                    //radioIssueDocument.Checked = false;
                    //radioIssueDocument.ForeColor = Color.White;
                    //radioNotIssue.Checked = false;
                    //radioNotIssue.ForeColor = Color.White;
                    groupBox8.Visible = false;

                    groupBox6.Visible = false;
                    //checkBoxInternal.Checked = false;
                    //checkBoxExternal.Checked = false;

                    getHistory(comboSearchType.Text);
                    getCount();
                }
                finally
                {
                    PleaseWait.Destroy();
                }
            }
        }

        private void txtScanID_Click(object sender, EventArgs e)
        {

            txtScanID.SelectAll();
            txtScanID.Focus();

            txtScanSetNo.Text = "";
            txtScanSetNo.Enabled = false;
            txtScanSetNo.BackColor = Color.Gray;

            txtSpacial.BackColor = Color.White;

        }

        private void txtScanID_TextChanged(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
            statusStrip1.Refresh();

            lblRequestName.Text = "";

            if (txtScanID.Text != "")
            {
                if (radioBorrow.Checked == true)
                {
                    if (radioIssueDocument.Checked == true)
                    {
                        if (checkBoxInternal.Checked == true | checkBoxExternal.Checked == true)
                        {
                            if (txtScanID.Text.Length == 8)
                            {
                                lblRequestName.Text = getAuthen(txtScanID.Text);
                                if (lblRequestName.Text != "")
                                {
                                    txtScanSetNo.Text = "";
                                    txtScanSetNo.Enabled = true;
                                    txtScanSetNo.BackColor = Color.Yellow;

                                    txtScanSetNo.SelectAll();
                                    txtScanSetNo.Focus();
                                }
                                else
                                {
                                    toolStripStatusLabel1.Text = "Status: Access denide, please try again.";
                                    statusStrip1.Refresh();

                                    lblRequestName.Text = "";
                                    txtScanID.SelectAll();
                                    txtScanID.Focus();
                                }


                            }

                        }
                        else
                        {
                            DialogResult result = MessageBox.Show("Please Select Purpose!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            if (result == DialogResult.OK)
                            {
                                txtScanID.Text = "";
                                txtScanID.Focus();
                                txtScanID.SelectAll();

                            }

                        }

                    }
                    else if (radioNotIssue.Checked == true)
                    {
                        if (txtScanID.Text.Trim() != "")
                        {
                            if (txtScanID.Text.Length == 8)
                            {
                                lblRequestName.Text = getAuthen(txtScanID.Text);
                                if (lblRequestName.Text != "")
                                {
                                    txtScanSetNo.Focus();
                                    txtScanSetNo.SelectAll();

                                    txtScanSetNo.Text = "";
                                    txtScanSetNo.Enabled = true;
                                    txtScanSetNo.BackColor = Color.Yellow;

                                    txtScanSetNo.SelectAll();
                                    txtScanSetNo.Focus();
                                }
                                else
                                {
                                    toolStripStatusLabel1.Text = "Status: Access denide, please try again.";
                                    statusStrip1.Refresh();

                                    lblRequestName.Text = "";
                                    txtScanID.SelectAll();
                                    txtScanID.Focus();
                                }


                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select issue document or not issue document", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtScanID.Text = "";
                    }

                }
                else if (radioReturn.Checked == true)
                {
                    if (radioNotIssue.Checked == false | radioIssueDocument.Checked == false)
                    {
                        if (txtScanID.Text.Length == 8)
                        {
                            lblRequestName.Text = getAuthen(txtScanID.Text);
                            if (lblRequestName.Text != "")
                            {
                                txtScanSetNo.Text = "";
                                txtScanSetNo.Enabled = true;
                                txtScanSetNo.BackColor = Color.Yellow;

                                txtScanSetNo.SelectAll();
                                txtScanSetNo.Focus();
                            }
                            else
                            {
                                toolStripStatusLabel1.Text = "Status: Access denide, please try again.";
                                statusStrip1.Refresh();

                                lblRequestName.Text = "";
                                txtScanID.SelectAll();
                                txtScanID.Focus();
                            }

                        }
                    }
                    else
                    {
                        MessageBox.Show("Not select issue document", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtScanID.Text = "";
                    }

                }

            }

        }


        private void txtScanSetNo_TextChanged(object sender, EventArgs e)
        {
            string strDueDateMultiBorrow = datePickerReturn.Value.Date.ToString("yyyy-MM-dd");
            //int flg = 0;
            toolStripStatusLabel1.Text = "";
            statusStrip1.Refresh();
            //string  strDueDate = "DATE_FORMAT(STR_TO_DATE('" +  datePickerReturn.Value.ToShortDateString() + "', '%c/%e/%Y %H:%i'), '%Y-%m-%d %H:%m:%s')";
            //string strDueDate = "DATE_FORMAT('" + datePickerReturn.Value.ToShortDateString() + "', '%Y-%m-%d %H:%m:%s')";
            string strDueDate1 = datePickerReturn.Value.Date.ToString("yyyy-MM-dd");
            string strDueDate = "DATE_FORMAT('" + datePickerReturn.Value.Date.ToString("yyyy-MM-dd HH:mm") + "', '%Y-%m-%d %H:%m:%s')";
            DateTime d = DateTime.Now;
            string dateDoc = d.ToString("yyyyMM");

                if (radioBorrow.Checked == true)
                {
                    if (radioNotIssue.Checked == true)
                    {
                        if (radioDefaultBorrow.Checked == true)
                        {
                            if (txtScanSetNo.Text.Trim() != "")
                            {
                                if (txtScanSetNo.Text.Length == 9)
                                {
                                    if (checkDup(txtScanSetNo.Text) == false)
                                    {
                                        if (getMasterData(txtScanSetNo.Text) == true)
                                        {
                                            DialogResult dialogResult = MessageBox.Show("Borrow item: " + txtScanSetNo.Text + "- Return Date: " + datePickerReturn.Value.ToShortDateString() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                            if (dialogResult == DialogResult.Yes)
                                            {
                                                PleaseWait.Create();
                                                try
                                                {
                                                    InsertBorrowNotIssue(txtScanSetNo.Text, lblRequestName.Text, txtSpacial.Text, strDueDate);

                                                    getHistory("NORMAL BORROW");
                                                    getCount();
                                                    txtScanID.Text = "";
                                                    txtScanSetNo.Text = "";
                                                    lblRequestName.Text = "";
                                                    txtScanID.Focus();
                                                    txtScanSetNo.Enabled = false;
                                                    txtScanSetNo.BackColor = Color.Gray;

                                                }
                                                finally
                                                {
                                                    PleaseWait.Destroy();
                                                }

                                                txtScanSetNo.SelectAll();
                                                txtScanSetNo.Focus();
                                            }
                                            else
                                            {
                                                PleaseWait.Create();
                                                try
                                                {
                                                    getHistory("NORMAL BORROW");
                                                    getCount();
                                                }
                                                finally
                                                {
                                                    PleaseWait.Destroy();
                                                }

                                                txtScanSetNo.Text = "";
                                                txtScanSetNo.SelectAll();
                                                txtScanSetNo.Focus();
                                            }
                                        }
                                        else
                                        {
                                            toolStripStatusLabel1.Text = "Status: Have no information in Master Data, please try another set.";
                                            statusStrip1.Refresh();

                                            txtScanSetNo.SelectAll();
                                            txtScanSetNo.Focus();
                                        }
                                    }
                                    else
                                    {

                                        toolStripStatusLabel1.Text = "Status: This set was borrowed, please try another set.";
                                        statusStrip1.Refresh();

                                        txtScanSetNo.SelectAll();
                                        txtScanSetNo.Focus();
                                          
                                    }
                                }
                            }
                        }
                        else if (radioMultipleBorrow.Checked == true)
                        {
                            if (txtScanSetNo.Text.Trim() != "")
                            {
                                if (txtScanSetNo.Text.Length == 9)
                                {
                                    //if (checkDup(txtScanSetNo.Text) == false)
                                    //{
                                    //    if (getMasterData(txtScanSetNo.Text) == true)
                                    //    {
                                            getData_MultiBorrow();

                                            txtScanSetNo.SelectAll();
                                            txtScanSetNo.Focus();
                                    //    }
                                    //}
                                }
                            }
                        }
                    }
                    else if (radioIssueDocument.Checked == true)
                    {
                        if (radioDefaultBorrow.Checked == true)
                        {
                            if (txtScanSetNo.Text.Trim() != "")
                            {
                                if (txtScanSetNo.Text.Length == 9)
                                {
                                    if (checkDup(txtScanSetNo.Text) == false)
                                    {
                                        if (getMasterData(txtScanSetNo.Text) == true)
                                        {

                                            DialogResult dialogResult = MessageBox.Show("Borrow item: " + txtScanSetNo.Text + "- Return Date: " + datePickerReturn.Value.ToShortDateString() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                            if (dialogResult == DialogResult.Yes)
                                            {
                                                PleaseWait.Create();
                                                try
                                                {
                                                    InsertBorrowIssue(txtScanSetNo.Text, lblRequestName.Text, txtSpacial.Text, strDueDate);
                                                    getHistory("NORMAL BORROW");
                                                    getCount();

                                                }
                                                finally
                                                {
                                                    PleaseWait.Destroy();
                                                }

                                                if (b == "SAMPLE SET CONTROL")
                                                {

                                                    if (checkBoxInternal.Checked == true)
                                                    {

                                                        Connection();
                                                        string dt2model;
                                                        string name = lblRequestName.Text;
                                                        string EIA = txtScanSetNo.Text.Trim();
                                                        MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM employee", conn);
                                                        DataTable tab = new DataTable();
                                                        adap.Fill(tab);

                                                        MySqlDataAdapter adapter = new MySqlDataAdapter("select * from master_set", conn);
                                                        DataTable tbb = new DataTable();
                                                        adapter.Fill(tbb);

                                                        foreach (DataRow rw in tab.Rows)
                                                        {
                                                            string setIDText = txtScanID.Text.Trim();
                                                            string setIDEmail = Convert.ToString(rw["serial"]);

                                                            if (setIDText == setIDEmail) //เช็ค id
                                                            {
                                                                foreach (DataRow roww in tbb.Rows) //set_no อุปกรณ์ที่ยืม
                                                                {
                                                                    string set_noexp = Convert.ToString(roww["SET_NO"]);
                                                                    if (txtScanSetNo.Text.Trim() == set_noexp)
                                                                    {
                                                                        EIARange = txtScanSetNo.Text.Trim();
                                                                        DocumentNoSampleMulti = "SSC";
                                                                        DepartmentSampleMulti = Convert.ToString(rw["dept"]);
                                                                        dt2model = Convert.ToString(roww["MODEL"]);
                                                                        frmPdfExport pdf = new frmPdfExport(DocumentNoSampleMulti, strDueDateMultiBorrow, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, dt2model, EIARange, "1", "Internal", "SAMPLE SET CONTROL",txtScanSetNo.Text.Trim());

                                                                    }

                                                                }

                                                            }

                                                        }
                                                        closeCon();

                                                    }
                                                    else if (checkBoxExternal.Checked == true)
                                                    {

                                                        Connection();
                                                        string dt2model;
                                                        string name = lblRequestName.Text;
                                                        string EIA = txtScanSetNo.Text.Trim();
                                                        MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM employee", conn);
                                                        DataTable tab = new DataTable();
                                                        adap.Fill(tab);

                                                        MySqlDataAdapter adapter = new MySqlDataAdapter("select * from master_set", conn);
                                                        DataTable tbb = new DataTable();
                                                        adapter.Fill(tbb);

                                                        foreach (DataRow rw in tab.Rows)
                                                        {
                                                            string setIDText = txtScanID.Text.Trim();
                                                            string setIDEmail = Convert.ToString(rw["serial"]);

                                                            if (setIDText == setIDEmail) //เช็ค id
                                                            {
                                                                foreach (DataRow roww in tbb.Rows) //set_no อุปกรณ์ที่ยืม
                                                                {
                                                                    string set_noexp = Convert.ToString(roww["SET_NO"]);
                                                                    if (txtScanSetNo.Text.Trim() == set_noexp)
                                                                    {
                                                                        EIARange = txtScanSetNo.Text.Trim();
                                                                        DocumentNoSampleMulti = "SSC";
                                                                        DepartmentSampleMulti = Convert.ToString(rw["dept"]);
                                                                        dt2model = Convert.ToString(roww["MODEL"]);
                                                                        frmPdfExport pdf = new frmPdfExport(DocumentNoSampleMulti, strDueDateMultiBorrow, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, dt2model, EIARange, "1", "External", "SAMPLE SET CONTROL", txtScanSetNo.Text.Trim());
                                                                    }

                                                                }

                                                            }

                                                        }
                                                        closeCon();
                                                    }
                                                }
                                                else if (b == "ENGINEERING TRAINING CENTER")
                                                {
                                                    if (checkBoxInternal.Checked == true)
                                                    {

                                                        Connection();
                                                        string dt2model;
                                                        string name = lblRequestName.Text;
                                                        string EIA = txtScanSetNo.Text.Trim();
                                                        MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM employee", conn);
                                                        DataTable tab = new DataTable();
                                                        adap.Fill(tab);

                                                        MySqlDataAdapter adapter = new MySqlDataAdapter("select * from master_set", conn);
                                                        DataTable tbb = new DataTable();
                                                        adapter.Fill(tbb);

                                                        foreach (DataRow rw in tab.Rows)
                                                        {
                                                            string setIDText = txtScanID.Text.Trim();
                                                            string setIDEmail = Convert.ToString(rw["serial"]);

                                                            if (setIDText == setIDEmail) //เช็ค id
                                                            {
                                                                foreach (DataRow roww in tbb.Rows) //set_no อุปกรณ์ที่ยืม
                                                                {
                                                                    string set_noexp = Convert.ToString(roww["SET_NO"]);
                                                                    if (txtScanSetNo.Text.Trim() == set_noexp)
                                                                    {
                                                                        EIARange = txtScanSetNo.Text.Trim();
                                                                        DocumentNoSampleMulti = "ETC";
                                                                        DepartmentSampleMulti = Convert.ToString(rw["dept"]);
                                                                        dt2model = Convert.ToString(roww["MODEL"]);
                                                                        frmPdfExport pdf = new frmPdfExport(DocumentNoSampleMulti, strDueDateMultiBorrow, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, dt2model, EIARange, "1", "Internal", "ENGINEERING TRAINING CENTER", txtScanSetNo.Text.Trim());
                                                                    }

                                                                }

                                                            }

                                                        }
                                                        closeCon();
                                                    }
                                                    else if (checkBoxExternal.Checked == true)
                                                    {

                                                        Connection();
                                                        string dt2model;
                                                        string name = lblRequestName.Text;
                                                        string EIA = txtScanSetNo.Text.Trim();
                                                        MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM employee", conn);
                                                        DataTable tab = new DataTable();
                                                        adap.Fill(tab);

                                                        MySqlDataAdapter adapter = new MySqlDataAdapter("select * from master_set", conn);
                                                        DataTable tbb = new DataTable();
                                                        adapter.Fill(tbb);

                                                        foreach (DataRow rw in tab.Rows)
                                                        {
                                                            string setIDText = txtScanID.Text.Trim();
                                                            string setIDEmail = Convert.ToString(rw["serial"]);

                                                            if (setIDText == setIDEmail) //เช็ค id
                                                            {
                                                                foreach (DataRow roww in tbb.Rows) //set_no อุปกรณ์ที่ยืม
                                                                {
                                                                    string set_noexp = Convert.ToString(roww["SET_NO"]);
                                                                    if (txtScanSetNo.Text.Trim() == set_noexp)
                                                                    {
                                                                        EIARange = txtScanSetNo.Text.Trim();
                                                                        DocumentNoSampleMulti = "ETC";
                                                                        DepartmentSampleMulti = Convert.ToString(rw["dept"]);
                                                                        dt2model = Convert.ToString(roww["MODEL"]);
                                                                        frmPdfExport pdf = new frmPdfExport(DocumentNoSampleMulti, strDueDateMultiBorrow, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, dt2model, EIARange, "1", "External", "ENGINEERING TRAINING CENTER",txtScanSetNo.Text.Trim());
                                                                    }

                                                                }

                                                            }

                                                        }
                                                        closeCon();
                                                    }
                                                }

                                                txtScanID.Text = "";
                                                txtScanSetNo.Text = "";
                                                txtScanID.Focus();
                                                txtScanID.SelectAll();
                                                txtScanSetNo.Enabled = false;
                                                txtScanSetNo.BackColor = Color.Gray;
                                            }

                                            else
                                            {
                                                PleaseWait.Create();
                                                try
                                                {
                                                    getHistory("NORMAL BORROW");
                                                    getCount();

                                                }
                                                finally
                                                {
                                                    PleaseWait.Destroy();
                                                }

                                                txtScanSetNo.Text = "";
                                                txtScanSetNo.SelectAll();
                                                txtScanSetNo.Focus();

                                            }
                                            txtScanID.Text = "";
                                            txtScanSetNo.Text = "";
                                            txtScanID.Focus();
                                            txtScanSetNo.Enabled = false;
                                            txtScanSetNo.BackColor = Color.Gray;

                                        }
                                        else
                                        {
                                            toolStripStatusLabel1.Text = "Status: Have no information in Master Data, please try another set.";
                                            statusStrip1.Refresh();

                                            txtScanSetNo.SelectAll();
                                            txtScanSetNo.Focus();
                                        }
                                    }
                                    else
                                    {
                                        toolStripStatusLabel1.Text = "Status: This set was borrowed, please try another set.";
                                        statusStrip1.Refresh();

                                        txtScanSetNo.SelectAll();
                                        txtScanSetNo.Focus();
                                    }
                                }
                            }
                        }
                        else if (radioMultipleBorrow.Checked == true)
                        {
                            if (txtScanSetNo.Text.Trim() != "")
                            {
                                if (txtScanSetNo.Text.Length == 9)
                                {
                                    //if (checkDup(txtScanSetNo.Text) == false)
                                    //{
                                    //    if (getMasterData(txtScanSetNo.Text) == true)
                                    //    {
                                            getData_MultiBorrow();

                                            txtScanSetNo.SelectAll();
                                            txtScanSetNo.Focus();
                                    //    }
                                    //}

                                }
                            }
                        }
                    }

                }
                else if (radioReturn.Checked == true)
                {
                    if (radioDefaultBorrow.Checked == true)
                    {
                        if (txtScanSetNo.Text.Trim() != "")
                        {
                            if (txtScanSetNo.Text.Length == 9)
                            {
                                if (checkDup(txtScanSetNo.Text) == true)
                                {
                                    DialogResult dialogResult = MessageBox.Show("Return item: " + txtScanSetNo.Text + "- By: " + lblRequestName.Text + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    if (dialogResult == DialogResult.Yes)
                                    {
                                        PleaseWait.Create();
                                        try
                                        {
                                            UpdateReturn(txtScanSetNo.Text, lblRequestName.Text);

                                            getHistory("NORMAL BORROW");
                                            getCount();

                                            txtScanID.Text = "";
                                            txtScanSetNo.Text = "";
                                            txtScanID.Focus();
                                            txtScanSetNo.Enabled = false;
                                            txtScanSetNo.BackColor = Color.Gray;

                                        }
                                        finally
                                        {
                                            PleaseWait.Destroy();
                                        }

                                        txtScanSetNo.SelectAll();
                                        txtScanSetNo.Focus();
                                    }
                                    else if (dialogResult == DialogResult.No)
                                    {
                                        txtScanSetNo.SelectAll();
                                        txtScanSetNo.Focus();
                                    }
                                }
                                else
                                {
                                    toolStripStatusLabel1.Text = "Status: Have no borrow information, please try another set.";
                                    statusStrip1.Refresh();

                                    txtScanSetNo.SelectAll();
                                    txtScanSetNo.Focus();
                                }
                            }
                        }
                    }
                    else if (radioMultipleBorrow.Checked == true)
                    {
                        if (txtScanSetNo.Text.Trim() != "")
                        {
                            if (txtScanSetNo.Text.Length == 9)
                            {
                                ReturnMultiBorrow();

                                txtScanSetNo.SelectAll();
                                txtScanSetNo.Focus();
                            }
                        }
                    }
                }
            

   }
        

        string a = "";
        string sqlcmd;
        int k;
        int n;
        int l;
        string[] getModel = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] A = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] B = new string[500];
        string eia;
        string check_cut;
        string checkcutA;

        public void getData_MultiBorrow()
        {
            
            int flg = 0;
            string _SetNo;
            string strDueDate = "DATE_FORMAT('" + datePickerReturn.Value.Date.ToString("yyyy-MM-dd HH:mm") + "', '%Y-%m-%d %H:%m:%s')";

            Connection();
            MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM master_set ", conn);
            DataTable results = new DataTable();
            adap.Fill(results);

            foreach (DataRow rows in results.Rows)
            {
                string flg1 = "";
                string flg3 = "";
                string flg2 = "";
                _SetNo = Convert.ToString(rows["SET_NO"]);
                string temp = Convert.ToString(rows["MODEL"]);
                if (txtScanSetNo.Text.Trim() == _SetNo) //เปรียบเทียบ set_no ที่พิมกับ set_no ในระบบ
                {
                    sqlcmd = "SELECT * FROM `set_data_update` where set_no = '" + txtScanSetNo.Text.Trim() + "' and flg_status = '1'";
                    MySqlCommand cmd = new MySqlCommand(sqlcmd, conn);
                    MySqlDataReader reader = cmd.ExecuteReader();
                    if (!reader.Read())
                    {
                        if (flg == 0)
                        {
                            for (int i = 0; i < dt2.Rows.Count; i++)
                            {
                                if (txtScanSetNo.Text.Trim() != dt2.Rows[i][0].ToString()) //set ที่จะยืมต้องไม่เท่ากับ set ที่มีใน dt2
                                { 
                                }
                                else
                                {
                                    toolStripStatusLabel1.Text = "Status : Serial Number duplicate, please try another set.";
                                    statusStrip1.Refresh();

                                    txtScanSetNo.SelectAll();
                                    txtScanSetNo.Focus();
                   
                                    a = "Dup";
                                }
                            }
                            if (a != "Dup")
                            {
                                eia = txtScanSetNo.Text.Trim();
                                check_cut = eia.Substring(3, 6); //003310 
                                checkcutA = check_cut.Substring(0,4); //0033 เริ่มที่0ถึงตัวที่4

                                    for (int i = 0; i <= 17; i++) //เช็ค Array Model 6 ตัว
                                    {
                                        if (getModel[i] != "")
                                        {
                                            k++;
                                        }
                                    }
                                    for (int j = 0; j <= 499; j++)
                                    {
                                        if (B[j] != null)
                                        {
                                            n++;
                                        }
                                    }
                                    for (int t = 0; t <= 17; t++)
                                    {
                                        if (A[t] != "")
                                        {
                                            l++;
                                        }
                                    }

                                    if (k == 0)
                                    {
                                        getModel[0] = Convert.ToString(rows["MODEL"]); // ถ้า k==0 เพิ่ม model ตัวแรก
                                    }
                                    if (n == 0)
                                    {           
                                        B[0] = check_cut;
                                    }
                                    if (l == 0)
                                    {
                                        A[0] = checkcutA;
                                    }

                                    if (k != 0)
                                    {
                                        for (int i = 0; i < k; i++)
                                        {
                                            if (getModel[i] == Convert.ToString(rows["MODEL"])) //เชคว่าโมเดลซ้ำไหม ถ้าซ้ำห้ามเพิ่ม model
                                            {
                                                flg1 = "noadd";
                                            }
                                        }
                                            if (flg1 != "noadd" && k <= 18) //ถ้า model ไม่ซ้ำและ model น้อยกว่าหรือเท่ากับ 6 ให้เพิ่ม model เข้าไป
                                            {
                                                if (k < 18)
                                                {
                                                    getModel[k] = Convert.ToString(rows["MODEL"]);
                      
                                                }
                                                else
                                                {
                                                    MessageBox.Show("model EIA failed ");
                                                    flg2 = "no show";
                                                }
                                            }
                                        k = 0;

                                    }

                                    if (n != 0)
                                    {
                                        for (int j = 0; j < n; j++) //เชคตัวที่ยัดเขา array ซ้ำไหม
                                        {
                                            if (B[j] == check_cut) 
                                            {
                                                flg3 = "noadd";
                                            }
                                        }
                                        if (flg3 != "noadd" && n <= 500) //
                                        {
                                            if (n < 500)
                                            {                   
                                                B[n] = check_cut;
                                            }
                                            else
                                            {
                                                MessageBox.Show("model EIA failed ");
                                                flg2 = "no show";
                                            }
                                        }
                                        n = 0;
                                    }

                                    if (l != 0)
                                    {
                                        for (int t = 0; t < l; t++)
                                        {
                                            if (A[t] == checkcutA)
                                            {
                                                flg3 = "noadd";
                                            }
                                        }
                                        if (flg3 != "noadd" && l <= 18) //เชค Array A เกิน limit ไหม
                                        {
                                            if (l < 18)
                                            {
                                                A[l] = checkcutA;
                                            }
                                            else
                                            {
                                                MessageBox.Show("model EIA failed");
                                                flg2 = "no show";
                                            }
                                        }
                                        l = 0;

                                    }

                                    if (flg2 != "no show")
                                    {
                                        dt2.Rows.Add(
                                    Convert.ToString(rows["SET_NO"])
                                   , Convert.ToString(rows["SECTION"])
                                   , Convert.ToString(rows["MODEL"])
                                   , Convert.ToString(rows["DETAIL"])
                                   , Convert.ToString(rows["PCN_NO"])
                                   , Convert.ToString(rows["SERIES"])
                                   , Convert.ToString(rows["EVENT"])
                                   , Convert.ToString(rows["SN_SET"])
                                   , Convert.ToString(rows["PURPOSE"])
                                   , Convert.ToString(rows["COL_WHERE"])
                                   , Convert.ToString(rows["TRANSFER"])
                                   , Convert.ToString(rows["TRANSFER_PIC"])
                                   , Convert.ToString(rows["DISPOSAL"]));

                                        if (radioIssueDocument.Checked == true)
                                        {
                                            dtGridViwerHitory.DataSource = dt2;
                                            InsertMultiple(Convert.ToString(rows["SET_NO"]), lblRequestName.Text, txtSpacial.Text.Trim(), strDueDate);
                                            Counter = dtGridViwerHitory.Rows.Count;
                                            txtScanSetNo.SelectAll();
                                            txtScanSetNo.Focus();
                                        }
                                        else if (radioNotIssue.Checked == true)
                                        {
                                            dtGridViwerHitory.DataSource = dt2;
                                            InsertMultipleNotIssue(Convert.ToString(rows["SET_NO"]), lblRequestName.Text, txtSpacial.Text.Trim(), strDueDate);
                                            Counter = dtGridViwerHitory.Rows.Count;
                                        }
            
                                    }

                                }
           
                            }
                            else
                            {

                            }

                        }
                        else
                        {
                            toolStripStatusLabel1.Text = "Status : Serial Number " + txtScanSetNo.Text + " has been borrowed, please try another set.";
                            statusStrip1.Refresh();

                            txtScanSetNo.SelectAll();
                            txtScanSetNo.Focus();

                        }

                    }

                }

            a = "";
            sqlcmd = "";
            closeCon();

         }

        int flgs = 0;
        string cmdCheckRows;
        public void ReturnMultiBorrow()
        {
            
            Connection();
            string strDueDate = "DATE_FORMAT('" + datePickerReturn.Value.Date.ToString("yyyy-MM-dd HH:mm") + "', '%Y-%m-%d %H:%m:%s')";

            MySqlDataAdapter adapter = new MySqlDataAdapter("select * from master_set", conn);
            DataTable dtTable = new DataTable();
            adapter.Fill(dtTable);
        
            foreach (DataRow dr in dtTable.Rows)
            {
                string CheckRow = Convert.ToString(dr["SET_NO"]);
                if (txtScanSetNo.Text.Trim() == CheckRow)
                {
                    cmdCheckRows = "select * from set_data_update where set_no = '" + txtScanSetNo.Text.Trim() + "' and flg_status = '1'";
                    MySqlCommand command = new MySqlCommand(cmdCheckRows, conn);
                    MySqlDataReader Reader = command.ExecuteReader();
                    if (Reader.Read())
                    {

                        if (flgs != 0)
                        {
                            for (int i = 0; i < dt2.Rows.Count; i++)
                            {
                                if (txtScanSetNo.Text.Trim() != dt2.Rows[i][0].ToString())
                                {
                                }
                                else
                                {
                                    toolStripStatusLabel1.Text = "Status : This serial number duplicate , please try another set.";
                                    statusStrip1.Refresh();

                                    txtScanSetNo.SelectAll();
                                    txtScanSetNo.Focus();

                                    a = "Dup";
                                }
                            }
                            if (a != "Dup")
                            {
                                dt2.Rows.Add(
                                           Convert.ToString(dr["SET_NO"])
                                         , Convert.ToString(dr["SECTION"])
                                         , Convert.ToString(dr["MODEL"])
                                         , Convert.ToString(dr["DETAIL"])
                                         , Convert.ToString(dr["PCN_NO"])
                                         , Convert.ToString(dr["SERIES"])
                                         , Convert.ToString(dr["EVENT"])
                                         , Convert.ToString(dr["SN_SET"])
                                         , Convert.ToString(dr["PURPOSE"])
                                         , Convert.ToString(dr["COL_WHERE"])
                                         , Convert.ToString(dr["TRANSFER"])
                                         , Convert.ToString(dr["TRANSFER_PIC"])
                                         , Convert.ToString(dr["DISPOSAL"])

                                          );

     
                                dtGridViwerHitory.DataSource = dt2;
                                UpdateReturnMultiBorrow(Convert.ToString(dr["SET_NO"]), lblRequestName.Text);
                                Counter = dtGridViwerHitory.Rows.Count;
                            }
                        }
                        else
                        {
                            dt2.Rows.Add(
                             Convert.ToString(dr["SET_NO"])
                           , Convert.ToString(dr["SECTION"])
                           , Convert.ToString(dr["MODEL"])
                           , Convert.ToString(dr["DETAIL"])
                           , Convert.ToString(dr["PCN_NO"])
                           , Convert.ToString(dr["SERIES"])
                           , Convert.ToString(dr["EVENT"])
                           , Convert.ToString(dr["SN_SET"])
                           , Convert.ToString(dr["PURPOSE"])
                           , Convert.ToString(dr["COL_WHERE"])
                           , Convert.ToString(dr["TRANSFER"])
                           , Convert.ToString(dr["TRANSFER_PIC"])
                           , Convert.ToString(dr["DISPOSAL"])

                            );

                            dtGridViwerHitory.DataSource = dt2;
                            UpdateReturnMultiBorrow(Convert.ToString(dr["SET_NO"]), lblRequestName.Text);
                            Counter = dtGridViwerHitory.Rows.Count;

                            flgs = 1;
                         
                        }
                    }
                    else
                    {
                        toolStripStatusLabel1.Text = "Status : This serial number has not been borrowed , please try another set.";
                        statusStrip1.Refresh();

                        txtScanSetNo.SelectAll();
                        txtScanSetNo.Focus();
                    }
                }
              
            }
  
            a = "";
            cmdCheckRows = "";
            closeCon();
        }

        private void AddARow(DataTable table)
        {
            // Use the NewRow method to create a DataRow with 
            // the table's schema.
            DataRow newRow = table.NewRow();

            // Add the row to the rows collection.
            table.Rows.Add(newRow);
        }

        public void InsertBorrowNotIssue(string strSetNo, string strRequestName, string strSpacial,string strDueDate)
        {
            Connection();
            string command = "";

            if (strSpacial == "")
            {
                command = "insert into set_data_update (set_no,request_name,request_date,due_date,flg_status,DOC_NO)values('" +
                strSetNo + "','" + strRequestName + "',sysdate()," + strDueDate + ",'1','Not Issue')";
            }
            else
            {
                command = "insert into set_data_update (set_no,request_name,request_date,due_date,flg_status,flg_spacial,DOC_NO)values('" +
                strSetNo + "','" + strRequestName + "',sysdate()," + strDueDate + ",'1','" + strSpacial + "','Not Issue')";

            }

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }

        public void InsertBorrowIssue(string strSetNo, string strRequestName, string strSpacial, string strDueDate)
        {
            Connection();
            string command = "";

            if (strSpacial == "")
            {
                command = "insert into set_data_update (set_no,request_name,request_date,due_date,flg_status)values('" +
                strSetNo + "','" + strRequestName + "',sysdate()," + strDueDate + ",'1')";
            }
            else
            {
                command = "insert into set_data_update (set_no,request_name,request_date,due_date,flg_status,flg_spacial)values('" +
                strSetNo + "','" + strRequestName + "',sysdate()," + strDueDate + ",'1','" + strSpacial + "')";

            }

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }




        public void UpdateReturn(string strSetNo, string strRequestName)
        {

            Connection();
            string command = "";

            command = "update set_data_update set return_name = '" + strRequestName + "' ,return_date = sysdate(),flg_status= '0'" + " where set_no = '" + strSetNo + "' and flg_status = '1'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }

        
        private Boolean checkDup(string strSetNo)
        {
 
            string strSQL = "select * from set_data_update where flg_status=1 and SET_NO = '" + strSetNo + "'";

            Connection();
            MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
            DataTable dts = new DataTable();

            da.Fill(dts);
            foreach (DataRow drData in dts.Rows)
            {
                return true;

            }
            closeCon();

            return false;
        }

        
        private Boolean getMasterData(string strSetNo)
        {

            string strSQL = "select * from master_set where SET_NO = '" + strSetNo + "'";

            Connection();
            MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
            DataTable dtt = new DataTable();

            da.Fill(dtt);
            foreach (DataRow drData in dtt.Rows)
            {
                return true;
            }
            closeCon();

            return false;
        }

        private string getAuthen(string strUser)
        {
            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
            string IP = iniconfig.IniReadValue("Server", "IP");
            string DB = iniconfig.IniReadValue("Server", "DB");

            string ConnectionString = "Server='" + IP + "';";
            ConnectionString += "User ID=sa;";
            ConnectionString += "Password=s;";
            ConnectionString += "Database='" + DB + "';";

            SQLConnection = new SqlConnection(ConnectionString);

            SQLConnection.Open();
            string strSQL = "SELECT * FROM [STTC_HUMAN_RESOURCE].[dbo].[TBL_MANPOW_EMPID_RFID] where EMPID = '" + strUser + "' ";

            SqlDataAdapter da = new SqlDataAdapter(strSQL,SQLConnection);
            DataTable dt = new DataTable();

            da.Fill(dt);
            foreach (DataRow drData in dt.Rows)
            {
                string name = Convert.ToString(drData["Remark"]);
                return name;
            }
            closeCon();

            return "";
        }
        

        private Boolean getAuthenAdmin(string strUser)
        {
            if (b == "SAMPLE SET CONTROL")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM `personal_login` where right(`userID`,8) = '" + value + "'", conn);
                DataTable table = new DataTable();
                adap.Fill(table);

                foreach (DataRow row in table.Rows)
                {
                    string group = Convert.ToString(row["Group"]);
                    if (group == "SSC")
                    {
                        iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                        string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                        string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                        string ConnectionString = "Server=" + IP + ";";
                        ConnectionString += "Uid=root;";
                        ConnectionString += "Password=123456*;";
                        ConnectionString += "Database=" + DB + ";";

                        conn = new MySqlConnection(ConnectionString);

                        conn.Open();
                        string strSQL = "SELECT * FROM `personal_login` where right(`userID`,8) = '" + value + "'";

                        MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        foreach (DataRow drData in dt.Rows)
                        {
                            return true;

                        }
                        closeCon();
                    }
                }
            }
            else if (b == "ENGINEERING TRAINING CENTER")
            {
                MySqlDataAdapter adap = new MySqlDataAdapter("SELECT * FROM `personal_login` where right(`userID`,8) = '" + value + "'", conn);
                DataTable table = new DataTable();
                adap.Fill(table);

                foreach (DataRow row in table.Rows)
                {
                    string group2 = Convert.ToString(row["Group"]);
                    if (group2 == "ETC")
                    {
                        iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                        string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                        string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                        string ConnectionString = "Server=" + IP + ";";
                        ConnectionString += "Uid=root;";
                        ConnectionString += "Password=123456*;";
                        ConnectionString += "Database=" + DB + ";";

                        conn = new MySqlConnection(ConnectionString);

                        conn.Open();
                        string strSQL = "SELECT * FROM `personal_login` where right(`userID`,8) = '" + value + "'";

                        MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        foreach (DataRow drData in dt.Rows)
                        {
                            return true;

                        }
                        closeCon();
                    }

                }
            }

            return false;
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

        DataTable t = new DataTable();
        public void getHistory(string type)
        {

            dtGridViwerHitory.DataSource = null;
            dtGridViwerHitory.Rows.Clear();
            dtGridViwerHitory.Refresh();
            bindingNavigator.BindingSource = null;
            tables.Clear();
            t.Clear();

            String strFromDate = null;
            String strToDate = null;

            string strFromD = datePickerFrom.Value.ToShortDateString();
            string strFromT = timePickerFrom.Value.ToLongTimeString();

            string strToD = datePickerTo.Value.ToShortDateString();
            string strToT = timePickerTo.Value.ToLongTimeString();

            strFromDate = "DATE_FORMAT(STR_TO_DATE('" + strFromD + " " + strFromT + "', '%c/%e/%Y %H:%i'), '%Y-%m-%d %H:%m:%s')";
            strToDate = "DATE_FORMAT(STR_TO_DATE('" + strToD + " " + strToT + "', '%c/%e/%Y %H:%i'), '%Y-%m-%d %H:%m:%s')";

            //DATE_FORMAT(STR_TO_DATE('5/16/2011 20:14 PM', '%c/%e/%Y %H:%i'), '%Y-%m-%d %H:%m:%s')

            string strSQL = "";

            if (type == "NORMAL BORROW")
            {
                //strSQL = "SELECT * FROM `set_data_update` WHERE flg_status = 1 and flg_spacial is null order by request_date desc";
                
                strSQL = "SELECT ";
                strSQL = strSQL + "A.auto_id,A.set_no,B.SECTION,B.MODEL,B.DETAIL,B.PCN_NO,B.SERIES,B.EVENT,B.SN_SET,B.PURPOSE,A.request_name,A.request_date,A.due_date,A.return_name,A.return_date,A.flg_status,A.flg_spacial " ;
                strSQL = strSQL + "FROM `set_data_update` A LEFT JOIN master_set B on A.set_no = B.set_no " ;
                strSQL = strSQL + "WHERE A.flg_status = 1 and A.flg_spacial is null order by A.request_date desc" ;
                

            }
            else if (type == "SPACIAL BORROW")
            {
                //strSQL = "select * from set_data_update where flg_status=1 and flg_spacial is not null order by request_date desc";
                
                strSQL = "SELECT ";
                strSQL = strSQL + "A.auto_id,A.set_no,B.SECTION,B.MODEL,B.DETAIL,B.PCN_NO,B.SERIES,B.EVENT,B.SN_SET,B.PURPOSE,A.request_name,A.request_date,A.due_date,A.return_name,A.return_date,A.flg_status,A.flg_spacial ";
                strSQL = strSQL + "FROM `set_data_update` A LEFT JOIN master_set B on A.set_no = B.set_no ";
                strSQL = strSQL + "where A.flg_status=1 and A.flg_spacial is not null order by A.request_date desc";

            }
            else if (type == "TOTAL")
            {
                if (comboOption.Text.Trim() == "" | txtSearchOption.Text.Trim() == "")
                {
                    strSQL = "select * from master_set order by auto_id";
                }
                else
                {
                    strSQL = "select * from master_set where " + comboOption.Text.Trim() + " like '%" + txtSearchOption.Text.Trim() + "%' order by auto_id";
                }
            }
            else if (type == "REMAIN")
            {
                if (comboOption.Text.Trim() == "" | txtSearchOption.Text.Trim() == "")
                {
                    strSQL = "select * from master_set where SET_NO not in (select SET_NO from set_data_update where flg_status=1) order by auto_id";
                }
                else
                {
                    strSQL = "select * from master_set where SET_NO not in (select SET_NO from set_data_update where flg_status=1) and " + comboOption.Text.Trim() + " like '%" + txtSearchOption.Text.Trim() + "%' order by auto_id";
                }
            }
            else if (type == "OVER DUE DATE")
            {
                strSQL = "SELECT * FROM `v_set_over_duedate` order by due_date desc";
            }
            else if (type == "HISTORY")
            {
                strSQL = "select * from set_data_update where request_date between " + strFromDate + " and " + strToDate;
            }
            else if (type == "DISPOSE")
            {
                if (comboOption.Text.Trim() == "" | txtSearchOption.Text.Trim() == "")
                {
                    strSQL = "select * from master_set_dispose  order by rec_date desc";
                }
                else
                {
                    strSQL = "select * from master_set_dispose where " + comboOption.Text.Trim() + " like '%" + txtSearchOption.Text.Trim() + "%' order by rec_date desc";
                }
            }
            else if (type == "TRANSFER")
            {
                if (comboOption.Text.Trim() == "" | txtSearchOption.Text.Trim() == "")
                {
                    strSQL = "select * from master_set_transfer order by rec_date desc";
                }
                else
                {
                    strSQL = "select * from master_set_transfer where " + comboOption.Text.Trim() + " like '%" + txtSearchOption.Text.Trim() + "%' order by rec_date desc";
                }
            }
            else if (type == "OVER ALL SET")
            {

                strSQL = "select * from v_overall_set order by SET_NO";
 
            }
            else if (type == "SET LOG")
            {
                if (comboOption.Text.Trim() == "" | txtSearchOption.Text.Trim() == "")
                {
                    strSQL = "select * from history_log  order by rec_date";
                }
                else
                {
                    strSQL = "select * from history_log where " + comboOption.Text.Trim() + " like '%" + txtSearchOption.Text.Trim() + "%' order by rec_date";
                }
            }           
            else
            {
                return;
            }

            // 1
            // Open connection
            Connection();
            {

                // 2
                // Create new DataAdapter
                using (MySqlDataAdapter da = new MySqlDataAdapter(
                    strSQL, conn))
                {
                    // 3
                    // Use DataAdapter to fill DataTable
                    

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
                            if (count >= 20)
                            {
                               count = 0;
                            }
                            //dt.Rows.Clear();
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

        private void getCount()
        {
            string strSQL = "SELECT * FROM v_count_status";          

            Connection();
            MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
            DataTable dt = new DataTable();

            da.Fill(dt);
            foreach (DataRow drData in dt.Rows)
            {
                lblTotal.Text = drData[0].ToString();
                lblRemain.Text = drData[1].ToString();
                lblSpacial.Text = drData[2].ToString();
                lblNormal.Text = drData[3].ToString();
                lblOverDue.Text = drData[4].ToString();
                lblOverAllSet.Text = drData[5].ToString(); 

            }
            closeCon();

            if (Int32.Parse(lblOverDue.Text) > 0)
            {
                timerOverDue.Enabled = true;
            }
            else
            {
                timerOverDue.Enabled = false;
                lblOverDue.BackColor = Color.Orange;
            }
        }

        private void timerOverDue_Tick(object sender, EventArgs e)
        {
            if (lblOverDue.BackColor == Color.Orange)
            {
                lblOverDue.BackColor = Color.Red;
            }
            else
            {
                lblOverDue.BackColor = Color.Orange;
            }
            
        }

        private void btSearch_Click(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getHistory(comboSearchType.Text);
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void lblOverDue_DoubleClick(object sender, EventArgs e)
        {
   
            PleaseWait.Create();
            try
            {
                getHistory("OVER DUE DATE");
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }


        private void lblTotal_DoubleClick(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getHistory("TOTAL");
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void lblRemain_DoubleClick(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getHistory("REMAIN");
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void lblSpacial_DoubleClick(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getHistory("SPACIAL BORROW");
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void lblNormal_DoubleClick(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getHistory("NORMAL BORROW");
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void txtScanSetNo_Click(object sender, EventArgs e)
        {
            txtScanSetNo.SelectAll();
            txtScanSetNo.Focus();

            txtSpacial.BackColor = Color.White;
        }

        private void checkSpacial_CheckedChanged(object sender, EventArgs e)
        {
            txtSpacial.Text = "";
        }

        private void timerCheck_Tick(object sender, EventArgs e)
        {
            if (DateTime.Now.ToString("HH:mm") == "12:00")
            {
                //MessageBox.Show(DateTime.Now.ToString("HH:mm"));
                getCount();
                Thread.Sleep(61000);
            }

        }

        string value;
        public void aDDDisposeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            value = Interaction.InputBox("Please enter your id [Last 8 digit]", "PASSWORD", "");
          
            if (value == "")
            {
                return;
            }

            if (getAuthenAdmin(value) == true)
            {
                checkselect();
                frmManage frmManage = new frmManage();
                frmManage.strAdminID = value;
                frmManage.select_db = b;
                frmManage.Show();
            }
            else
            {
                MessageBox.Show("Access denied:" + value);
            }
            
        }

        private void addNewMemberToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            value = Interaction.InputBox("Please enter your id [Last 8 digit]", "PASSWORD", "");

            if (value == "")
            {
                return;
            }

            if (getAuthenAdmin(value) == true)
            {
                checkselect();
                frmMember frmMember = new frmMember();
                frmMember.select_db = b;
                frmMember.Show();

            }
            else
            {
                MessageBox.Show("Access denied:" + value);
            }
        }

        private void btRefresh_Click(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getHistory(comboSearchType.Text);
                getCount();
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void pictureBox2_DoubleClick(object sender, EventArgs e)
        {
 
            PleaseWait.Create();
            try
            {
                ExportToExcel exportGrid2excel = new ExportToExcel();
                exportGrid2excel.export2excel(t);
            }
            finally
            {
                PleaseWait.Destroy();
            }

        }

        private void dtGridViwerHitory_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (comboSearchType.Text == "DISPOSE" || comboSearchType.Text == "TRANSFER")
            {


                if (dtGridViwerHitory.SelectedCells.Count > 0 & e.ColumnIndex >= 0 & e.RowIndex >= 0)
                {

                    dtGridViwerHitory.CurrentCell = dtGridViwerHitory.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    //Can leave these here - doesn't hurt
                    dtGridViwerHitory.Rows[e.RowIndex].Selected = true;
                    dtGridViwerHitory.Focus();

                    int selectedrowindex = dtGridViwerHitory.SelectedCells[0].RowIndex;
                    DataGridViewRow selectedRow = dtGridViwerHitory.Rows[selectedrowindex];

                    strSetNo = Convert.ToString(selectedRow.Cells[1].Value.ToString());

                    if (strSetNo == "")
                    {
                        pictureBoxRestore.Visible = false;
                        return;
                    }

                    strSection = Convert.ToString(selectedRow.Cells[2].Value.ToString());
                    strModel = Convert.ToString(selectedRow.Cells[3].Value.ToString());
                    strDetail = Convert.ToString(selectedRow.Cells[4].Value.ToString());
                    strPCN = Convert.ToString(selectedRow.Cells[5].Value.ToString());
                    strSeries = Convert.ToString(selectedRow.Cells[6].Value.ToString());
                    strEvent = Convert.ToString(selectedRow.Cells[7].Value.ToString());
                    strSNSet = Convert.ToString(selectedRow.Cells[8].Value.ToString());
                    strPurpose = Convert.ToString(selectedRow.Cells[9].Value.ToString());
                    strWhere = Convert.ToString(selectedRow.Cells[10].Value.ToString());

                    pictureBoxRestore.Visible = true;
                }

            }else
            {
                pictureBoxRestore.Visible = false;
            }
        }

        private void comboSearchType_MouseClick(object sender, MouseEventArgs e)
        {
            pictureBoxRestore.Visible = false;

            strSetNo = "";
            strSection = "";
            strModel = "";
            strDetail = "";
            strPCN = "";
            strSeries = "";
            strEvent = "";
            strSNSet = "";
            strPurpose = "";
            strWhere = "";   

                             
        }

        private void pictureBoxRestore_DoubleClick(object sender, EventArgs e)
        {
            value = Interaction.InputBox("Do you want to restore item: " + strSetNo  + "?\n" + "Please enter your id." , "Confirm", "");

            if (value == "")
            {
                return;
            }

            if (getAuthenAdmin(value) == true)
            {

                string strDocNo = "";

                strDocNo = Interaction.InputBox("Please enter restore document number.", "Enter Document Number", "");

                PleaseWait.Create();
                try
                {
                    InsertMaster(strSetNo, strSection, strModel, strDetail, strPCN, strSeries, strEvent, strSNSet, strPurpose, strWhere);

                    if (comboSearchType.Text == "TRANSFER")
                    {
                        InsertHistory(strSetNo, strSection, strModel, strDetail, strPCN, strSeries, strEvent, strSNSet, strPurpose, strWhere, strDocNo, "", "", value, "Transfer -> Master");   
                        deleteTransfer(strSetNo);
                    }
                    else if (comboSearchType.Text == "DISPOSE")
                    {
                        InsertHistory(strSetNo, strSection, strModel, strDetail, strPCN, strSeries, strEvent, strSNSet, strPurpose, strWhere, "", "", strDocNo, value, "Dispose -> Master"); 
                        deleteDispose(strSetNo);
                    }

                    getHistory(comboSearchType.Text);
                    getCount();
                }
                finally
                {
                    PleaseWait.Destroy();
                }

                MessageBox.Show("Restore Completed:" + strSetNo);

            }
            else
            {
                MessageBox.Show("Access Denied:" + value);
            }
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

        public void deleteTransfer(string strSetNo)
        {

            Connection();
            string command = "";

            command = "delete from master_set_transfer where `SET_NO`='" + strSetNo + "'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }

        public void deleteDispose(string strSetNo)
        {

            Connection();
            string command = "";

            command = "delete from master_set_dispose where `SET_NO`='" + strSetNo + "'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();

        }

        private void comboSearchType_DropDownClosed(object sender, EventArgs e)
        {
            if (comboSearchType.Text == "TOTAL" | comboSearchType.Text == "REMAIN" | comboSearchType.Text == "DISPOSE" | comboSearchType.Text == "TRANSFER" | comboSearchType.Text == "SET LOG")
            {
                GroupOption.Enabled = true;

            }
            else
            {
                txtSearchOption.Text = "";
                GroupOption.Enabled = false;
            }
        }

        private void lblOverAllSet_DoubleClick(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getHistory("OVER ALL SET");
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void radioDefaultBorrow_Click(object sender, EventArgs e)
        {
            if (radioBorrow.Checked == true)
            {
                if (radioDefaultBorrow.Checked == true)
                {
                    radioIssueDocument.Checked = false;
                    radioNotIssue.Checked = false;
                    groupBox6.Visible = true;
                    radioIssueDocument.ForeColor = Color.White;
                    radioNotIssue.ForeColor = Color.White;

                    radioDefaultBorrow.ForeColor = Color.Yellow;
                    radioMultipleBorrow.ForeColor = Color.White;
                    btnCancleBorrow.Visible = false;
                    btnMultipleBorrow.Visible = false;
                    txtScanSetNo.Enabled = false;
                    txtScanSetNo.BackColor = Color.Gray;
                   
                    PleaseWait.Create();
                    try
                    {
                        txtScanID.Text = "";
                        txtScanSetNo.Text = "";
                        txtScanID.Enabled = true;
                        txtScanID.ForeColor = Color.Black;
                        txtScanID.BackColor = Color.Yellow;
                        txtScanID.Focus();
                        getHistory(comboSearchType.Text);
                        getCount();
                        tables.Clear();
                        dt2.Clear();

                    }
                    finally
                    {
                        PleaseWait.Destroy();
                    }
                }
                else if (radioMultipleBorrow.Checked == false)
                {
                    radioMultipleBorrow.ForeColor = Color.White;
                }
            }
            else if (radioReturn.Checked == true)
            {
                if (radioDefaultBorrow.Checked == true)
                {
                    groupBox6.Visible = false;
                    groupBox8.Visible = false;

                    radioDefaultBorrow.ForeColor = Color.Yellow;
                    radioMultipleBorrow.ForeColor = Color.White;
                    btnCancleBorrow.Visible = false;
                    btnMultipleBorrow.Visible = false;
                    txtScanSetNo.Text = "";
                    txtScanSetNo.Enabled = false;
                    txtScanSetNo.BackColor = Color.Gray;

                    PleaseWait.Create();
                    try
                    {
                        txtScanID.Text = "";
                        txtScanID.Enabled = true;
                        //txtScanID.BackColor = Color.Yellow;
                        txtScanID.Focus();
                        getHistory(comboSearchType.Text);
                        getCount();
                        tables.Clear();
                        dt2.Clear();

                    }
                    finally
                    {
                        PleaseWait.Destroy();
                    }
                }
                else if (radioMultipleBorrow.Checked == false)
                {
                    radioMultipleBorrow.ForeColor = Color.White;
                }
            }
            else
            {
                radioDefaultBorrow.Checked = false;
                MessageBox.Show("Please Select Borrow or Return", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
        }

        private void radioMultipleBorrow_Click(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                if (radioBorrow.Checked == true)
                {
                    if (radioMultipleBorrow.Checked == true)
                    {
                        getHistory("NORMAL BORROW");

                        radioIssueDocument.Checked = false;
                        radioNotIssue.Checked = false;
                        groupBox6.Visible = true;
                        radioIssueDocument.ForeColor = Color.White;
                        radioNotIssue.ForeColor = Color.White;

                        checkBoxInternal.Checked = false;
                        checkBoxExternal.Checked = false;
                        radioDefaultBorrow.ForeColor = Color.White;
                        radioMultipleBorrow.ForeColor = Color.Yellow;
                        btnMultipleBorrow.Visible = true;
                        btnCancleBorrow.Visible = true;
                     
                        groupBox8.Visible = true;

                        txtScanSetNo.Text = "";
                        txtScanSetNo.Enabled = false;
                        txtScanSetNo.BackColor = Color.Gray;
                        txtScanID.Enabled = true;
                        txtScanID.ForeColor = Color.Black;
                        txtScanID.BackColor = Color.Yellow;
                        dtGridViwerHitory.DataSource = null;
                        dt2.Clear();
                        tables.Clear();
                        
                        txtScanID.Text = "";
                        txtScanID.SelectAll();
                        txtScanID.Focus();


                    }
                }
                else if (radioReturn.Checked == true)
                {
                    if (radioMultipleBorrow.Checked == true)
                    {
                        getHistory("NORMAL BORROW");
                        groupBox6.Visible = false;
                        groupBox8.Visible = false;

                        radioDefaultBorrow.ForeColor = Color.White;
                        radioMultipleBorrow.ForeColor = Color.Yellow;
                        btnMultipleBorrow.Visible = true;
                        btnCancleBorrow.Visible = true;

                        txtScanSetNo.Text = "";
                        txtScanSetNo.Enabled = false;
                        txtScanSetNo.BackColor = Color.Gray;
                        txtScanID.Enabled = true;
                        txtScanID.ForeColor = Color.Black;
                        txtScanID.BackColor = Color.Yellow;
                        dtGridViwerHitory.DataSource = null;
                        dt2.Clear();
                        tables.Clear();

                        txtScanID.Text = "";
                        txtScanID.SelectAll();
                        txtScanID.Focus();
                    }

                }
                else
                {
                    radioMultipleBorrow.Checked = false;
                    MessageBox.Show("Please Select Borrow or Return", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            finally
            {
                PleaseWait.Destroy();
            }

        }

        public void InsertMultiBorrowDocument(string strSetInsert, string strID, string strName, string strDepartment)
        {
            Connection();
            string command = "";
            command = "insert into agreement_document(SET_NO, PERSONAL_ID, PERSONAL_NAME, DEPARTMENT) values('" + strSetInsert + "','"+strID+"', '"+strName+"', '"+strDepartment+"')";

            MySqlCommand cmd = new MySqlCommand(command,conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }


        string EIAFirst;
        string EIALast;
        string EIARange;
        string DocumentNoSampleMulti;
        string DepartmentSampleMulti;
        DataTable tab = new DataTable();
        int[] C = new int[501];
        string[] D = new string[501];
       // string[] ArrRange = new string[1000];
        string modelBor;
        string eiaBor;
        string Unit = "";
        string N;

        private void btnMultipleBorrow_Click(object sender, EventArgs e)
        {
            
            int CountBorrow2 = 0;
            int AUTO_ID = 0;
            string strDueDateMultiBorrow = datePickerReturn.Value.Date.ToString("yyyy-MM-dd");
            DateTime d = DateTime.Now;
            string dateDoc = d.ToString("yyyyMM");
            DateTime df = DateTime.Now;
            string DateBorrow = df.ToString("yyyy-MM-dd");

            toolStripStatusLabel1.Text = "";
            statusStrip1.Refresh();
            
            if (Counter == 0)
            {
                toolStripStatusLabel1.Text = "Status: No Item Borrrow, plaease add data.";
                statusStrip1.Refresh();

                txtScanSetNo.SelectAll();
                txtScanSetNo.Focus();
            }
            else
            {
                if (cmd_multiborrow != "")
                {
                    if (radioIssueDocument.Checked == true)
                    {

                        if (b == "SAMPLE SET CONTROL")
                        {
                            newdata.Rows.Clear();
                            int v;
                            IniFile Gen;
                            Gen = new IniFile(Application.StartupPath + "\\generate.ini");
                            DateTime dd = DateTime.Now;
                            //dateDoc = dd.ToString("yyyyMM");

                            datecheck();
                            v = Convert.ToInt32(Gen.IniReadValue("generate", "gen_ssc"));
                            v = v + 1;
                            number = v.ToString();
                            Gen.IniWriteValue("generate", "gen_ssc", number);
                            string DocumentNumber = "SSC" + dateDoc + "-" + number;


                            DialogResult result = MessageBox.Show("Borrow item: " + Counter + "- Return Date: " + datePickerReturn.Value.ToShortDateString() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (result == DialogResult.Yes)
                            {
                                Connection();
                                MySqlCommand cmd = new MySqlCommand(cmd_multiborrow, conn);
                                cmd.ExecuteNonQuery();
                                closeCon();

                                if (checkBoxInternal.Checked == true)
                                {

                                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                                    string IP = iniconfig.IniReadValue("Employee_Server", "IP");
                                    string DB = iniconfig.IniReadValue("Employee_Server", "DB");

                                    string ConnectionString = "Server=" + IP + ";";
                                    ConnectionString += "Uid=root;";
                                    ConnectionString += "Password=123456;";
                                    ConnectionString += "Database=" + DB + ";";

                                    conn = new MySqlConnection(ConnectionString);

                                    conn.Open();
                                    string strSQL = "SELECT serial,dept FROM `employee` where serial = '" + txtScanID.Text.Trim() + "'";

                                    MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
                                    DataTable dt = new DataTable();
                                    da.Fill(dt);
                                    foreach (DataRow drData in dt.Rows)
                                    {
                                        string IDtext = txtScanID.Text.Trim();
                                        string IDcheck = Convert.ToString(drData["serial"]);
                                        DepartmentSampleMulti = Convert.ToString(drData["dept"]);

                                        int min = 999999;
                                        int max = 0;

                                        Array.Sort(B);
                                        Array.Reverse(B);
                                        Array.Sort(A);
                                        Array.Reverse(A);

                                        int i = 1;
                                        foreach (var r in B)
                                        {
                                            int recieve = Convert.ToInt32(r);
                                            C[i] = recieve;
                                            i++;
                                        }

                                        int p = 0;
                                        foreach (var h in B)
                                        {
                                            D[p] = h;
                                            string sett = "EIA" + D[p];
                                            Connection();
                                            com = "update set_data_update set DOC_NO = '" + DocumentNumber + "' where set_no ='" + sett + "' and flg_status = '1'";
                                            MySqlCommand cm = new MySqlCommand(com, conn);
                                            cm.ExecuteNonQuery();
                                            closeCon();
                                            p++;
                                        }

                                        for (int s = 0; s <= 17; s++)
                                        {
                                           
                                            string one = A[s];
                                            var EW = from S in C
                                                     where S != 0 
                                                     //where S.ToString().Substring(0, 2) == one
                                                    select S;

                                            int K = EW.Count();
                                            
                                            for (int r = 1; r <= K; r++)
                                            {
                                                if ((C[r-1] - C[r]) == 1 || r == 1)
                                                {
                                                    int bbb = C[r];

                                                   // aaa = Convert.ToString(bbb);
                                                    N = bbb.ToString("000000"); //003310

                                                    if (N != "")
                                                    {
                                                        //string N = "00" + aaa;
                                                        //string N = bbb.ToString("000000");

                                                        string check_eia = "EIA" + N; //EIA003310
                                                        foreach (DataRow dr in dt2.Rows)
                                                        {
                                                            eiaBor = Convert.ToString(dr["SET_NO"]);
                                                            if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                            {
                                                                if (N.Substring(0, 4) == one) //N.Substring = 33 เปรียบเทียบกับ one = 33
                                                                {
                                                                    modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                                    int int_cut = Convert.ToInt32(N); //convert aaa string to int aaa = 3310

                                                                    if (min > int_cut) // 9999 > 3311,3311 > 3310
                                                                    {
                                                                        min = int_cut; //3311,3310
                                                                    }
                                                                    if (max < int_cut)
                                                                    {
                                                                        max = int_cut;

                                                                    }

                                                                    EIAFirst = min.ToString("000000");
                                                                    EIALast = max.ToString("000000");
                                                                    if (EIAFirst == EIALast)
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst;
                                                                        Unit = "Set";
                                                                    }
                                                                    else
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst + "-" + "EIA" + EIALast;
                                                                        Unit = "Set";
                                                                    }

                                                                    CountBorrow2++;
                                                                    AUTO_ID++;
                                                                }

                                                            }

                                                        }
                                                    }

                                                }
                                                else if ((C[r - 1] - C[r]) != 1 && r != 1)
                                                {

                                                   if (r != 0 && one == N.Substring(0, 4))
                                                  {
                                                    newdata.Rows.Add(Unit, DocumentNumber, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, modelBor, EIARange, CountBorrow2, "Internal", DateBorrow);
                                                    min = 999999;
                                                    max = 0;
                                                    CountBorrow2 = 0;
                                                    EIARange = "";
                                                   
                                                  }

                                                    int bbb = C[r];
                                                    N = bbb.ToString("000000"); //003310
                                                  
                                                    if (N != "")
                                                    {
                                                        string check_eia = "EIA" + N; //EIA003310
                                                        foreach (DataRow dr in dt2.Rows)
                                                        {
                                                            eiaBor = Convert.ToString(dr["SET_NO"]);
                                                            if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                            {
                                                                if (N.Substring(0, 4) == one) //cut aaa = 0033 เปรียบเทียบกับ one = 0033
                                                                {
                                                                    modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                                    int int_cut = Convert.ToInt32(N); //convert aaa string to int aaa = 3310

                                                                    if (min > int_cut) // 9999 > 3311,3311 > 3310
                                                                    {
                                                                        min = int_cut; //3311,3310
                                                                    }
                                                                    if (max < int_cut)
                                                                    {
                                                                        max = int_cut;

                                                                    }

                                                                    EIAFirst = min.ToString("000000");
                                                                    EIALast = max.ToString("000000");
                                                                    if (EIAFirst == EIALast)
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst;
                                                                        Unit = "Set";
                                                                    }
                                                                    else
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst + "-" + "EIA" + EIALast;
                                                                        Unit = "Set";
                                                                    }

                                                                    
                                                                    CountBorrow2++;
                                                                    AUTO_ID++;
                                                                }

                                                            }

                                                        }
                                                    }

                                                }

                                            }

                                            if (one != "")
                                            {
                                                if (EIARange != "")
                                                {
                                                    newdata.Rows.Add(Unit, DocumentNumber, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, modelBor, EIARange, CountBorrow2, "Internal", DateBorrow);
                                                    min = 999999;
                                                    max = 0;
                                                    CountBorrow2 = 0;
                                                    AUTO_ID = 0 ;
                                                    one = "";
                                                }
                                            }
                                           

                                        }
                                        l = 0;
                                        n = 0;
                                        k = 0;
                                        getModel = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                                        A = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                                        B = new string[500];
                                        C = new int[501];
                                        D = new string[501];

                                        frmPdfExport expPDF = new frmPdfExport();
                                        expPDF.select_db = b;
                                        expPDF.exportPDF(newdata);

                                    }

                                }
                                else if (checkBoxExternal.Checked == true)
                                {

                                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                                    string IP = iniconfig.IniReadValue("Employee_Server", "IP");
                                    string DB = iniconfig.IniReadValue("Employee_Server", "DB");

                                    string ConnectionString = "Server=" + IP + ";";
                                    ConnectionString += "Uid=root;";
                                    ConnectionString += "Password=123456;";
                                    ConnectionString += "Database=" + DB + ";";

                                    conn = new MySqlConnection(ConnectionString);

                                    conn.Open();
                                    string strSQL = "SELECT serial,dept FROM `employee` where serial = '" + txtScanID.Text.Trim() + "'";

                                    MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
                                    DataTable dt = new DataTable();
                                    da.Fill(dt);
                                    foreach (DataRow drData in dt.Rows)
                                    {
                                        string IDtext = txtScanID.Text.Trim();
                                        string IDcheck = Convert.ToString(drData["serial"]);
                                        DepartmentSampleMulti = Convert.ToString(drData["dept"]);

                                        int min = 999999;
                                        int max = 0;

                                        Array.Sort(B);
                                        Array.Reverse(B);
                                        Array.Sort(A);
                                        Array.Reverse(A);

                                        int i = 1;
                                        foreach (var r in B)
                                        {
                                            int recieve = Convert.ToInt32(r);
                                            C[i] = recieve;
                                            i++;
                                        }

                                        int p = 0;
                                        foreach (var h in B)
                                        {
                                            D[p] = h;
                                            string sett = "EIA" + D[p];
                                            Connection();
                                            com = "update set_data_update set DOC_NO = '" + DocumentNumber + "' where set_no ='" + sett + "' and flg_status = '1'";
                                            MySqlCommand cm = new MySqlCommand(com, conn);
                                            cm.ExecuteNonQuery();
                                            closeCon();
                                            p++;
                                        }

                                        for (int s = 0; s <= 17; s++)
                                        {

                                            string one = A[s];
                                            var EW = from S in C
                                                     where S != 0
                                                     //where S.ToString().Substring(0, 2) == one
                                                     select S;

                                            int K = EW.Count();

                                            for (int r = 1; r <= K; r++)
                                            {
                                                if ((C[r - 1] - C[r]) == 1 || r == 1)
                                                {
                                                    int bbb = C[r];
                                                    //N = Convert.ToString(bbb);
                                                    N = bbb.ToString("000000"); //003310
                                                    if (N != "")
                                                    {
                                                        //string N = "00" + aaa;
                                                        string check_eia = "EIA" + N;
                                                        foreach (DataRow dr in dt2.Rows)
                                                        {
                                                            eiaBor = Convert.ToString(dr["SET_NO"]);
                                                            if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                            {
                                                                if (N.Substring(0, 4) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                                {
                                                                    modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                                    int int_cut = Convert.ToInt32(N); //convert aaa string to int aaa = 331

                                                                    if (min > int_cut) // 9999 > 3311,3311 > 3310
                                                                    {
                                                                        min = int_cut; //3311,3310
                                                                    }
                                                                    if (max < int_cut)
                                                                    {
                                                                        max = int_cut;

                                                                    }

                                                                    EIAFirst = min.ToString("000000");
                                                                    EIALast = max.ToString("000000");
                                                                    if (EIAFirst == EIALast)
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst;
                                                                        Unit = "Set";
                                                                    }
                                                                    else
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst + "-" + "EIA" + EIALast;
                                                                        Unit = "Set";
                                                                    }
                                                                    DocumentNoSampleMulti = "SSC";
                                                                    CountBorrow2++;
                                                                    AUTO_ID++;
                                                                }

                                                            }

                                                        }
                                                    }

                                                }
                                                else if ((C[r - 1] - C[r]) != 1 && r != 1)
                                                {

                                                    if (r != 0 && one == N.Substring(0, 4))
                                                    {
                                                        newdata.Rows.Add(Unit, DocumentNumber, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, modelBor, EIARange, CountBorrow2, "External", DateBorrow);
                                                        min = 999999;
                                                        max = 0;
                                                        CountBorrow2 = 0;
                                                        EIARange = "";

                                                    }

                                                    int bbb = C[r];
                                                    N = bbb.ToString("000000"); //003310
                                                    if (N != "")
                                                    {
                                                        //string N = "00" + aaa;
                                                        string check_eia = "EIA" + N;
                                                        foreach (DataRow dr in dt2.Rows)
                                                        {
                                                            eiaBor = Convert.ToString(dr["SET_NO"]);
                                                            if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                            {
                                                                if (N.Substring(0, 4) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                                {
                                                                    modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                                    int int_cut = Convert.ToInt32(N); //convert aaa string to int aaa = 3310

                                                                    if (min > int_cut) // 9999 > 3311,3311 > 3310
                                                                    {
                                                                        min = int_cut; //3311,3310
                                                                    }
                                                                    if (max < int_cut)
                                                                    {
                                                                        max = int_cut;

                                                                    }

                                                                    EIAFirst = min.ToString("000000");
                                                                    EIALast = max.ToString("000000");
                                                                    if (EIAFirst == EIALast)
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst;
                                                                        Unit = "Set";
                                                                    }
                                                                    else
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst + "-" + "EIA" + EIALast;
                                                                        Unit = "Set";
                                                                    }

                                                                    DocumentNoSampleMulti = "SSC";
                                                                    CountBorrow2++;
                                                                    AUTO_ID++;
                                                                }

                                                            }

                                                        }
                                                    }

                                                }

                                            }

                                            if (one != "")
                                            {
                                                if (EIARange != "")
                                                {
                                                    newdata.Rows.Add(Unit, DocumentNumber, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, modelBor, EIARange, CountBorrow2, "External", DateBorrow);
                                                    min = 999999;
                                                    max = 0;
                                                    CountBorrow2 = 0;
                                                    one = "";
                                                }
                                            }

                                        }

                                        l = 0;
                                        n = 0;
                                        k = 0;
                                        getModel = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                                        A = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                                        B = new string[500];
                                        C = new int[501];
                                        D = new string[501];
                                        frmPdfExport expPDF = new frmPdfExport();
                                        expPDF.select_db = b;
                                        expPDF.exportPDF(newdata);

                                    }
 
                                }

                                getHistory("NORMAL BORROW");
                                getCount();

                                Unit = "";
                                DocumentNumber = "";
                                lblRequestName.Text = "";
                                DepartmentSampleMulti = "";
                                modelBor = "";
                                EIARange = "";
                                CountBorrow2 = 0;
                                DateBorrow = "";

                                txtScanID.Text = "";
                                txtScanID.Focus();
                                txtScanID.Select();
                                txtScanSetNo.Text = "";
                                txtScanSetNo.Enabled = false;
                                txtScanSetNo.BackColor = Color.Gray;
                                checkBoxInternal.Checked = false;
                                checkBoxExternal.Checked = false;
                                radioMultipleBorrow.Checked = false;
                                radioMultipleBorrow.ForeColor = Color.White;
                                radioDefaultBorrow.Checked = true;
                                radioDefaultBorrow.ForeColor = Color.Yellow;
                                radioIssueDocument.Checked = false;
                                radioIssueDocument.ForeColor = Color.White;
                                btnMultipleBorrow.Visible = false;
                                btnCancleBorrow.Visible = false;
                                newdata.Clear();
                                dt2.Clear();
                                cmd_multiborrow = "";
                            }
                            else if (result == DialogResult.No)
                            {
                                txtScanSetNo.Focus();
                                txtScanSetNo.SelectAll();
                            }
                            
                        }
                        else if (b == "ENGINEERING TRAINING CENTER")
                        {
                            int v;
                            IniFile Gen;
                            Gen = new IniFile(Application.StartupPath + "\\generate.ini");
                            DateTime dd = DateTime.Now;
                            //dateDoc = dd.ToString("yyyyMM");
                            datecheck();

                            v = Convert.ToInt32(Gen.IniReadValue("generate", "gen_etc"));
                            v = v + 1;
                            number = v.ToString();
                            Gen.IniWriteValue("generate", "gen_etc", number);
                            string DocumentNumber = "ETC" + dateDoc + "-" + number;

                            DialogResult result = MessageBox.Show("Borrow item: " + Counter + "- Return Date: " + datePickerReturn.Value.ToShortDateString() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (result == DialogResult.Yes)
                            {
                                Connection();
                                MySqlCommand cmd = new MySqlCommand(cmd_multiborrow, conn);
                                cmd.ExecuteNonQuery();
                                closeCon();

                                if (checkBoxInternal.Checked == true)
                                {
                                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                                    string IP = iniconfig.IniReadValue("Employee_Server", "IP");
                                    string DB = iniconfig.IniReadValue("Employee_Server", "DB");

                                    string ConnectionString = "Server=" + IP + ";";
                                    ConnectionString += "Uid=root;";
                                    ConnectionString += "Password=123456;";
                                    ConnectionString += "Database=" + DB + ";";

                                    conn = new MySqlConnection(ConnectionString);

                                    conn.Open();
                                    string strSQL = "SELECT serial,dept FROM `employee` where serial = '" + txtScanID.Text.Trim() + "'";

                                    MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
                                    DataTable dt = new DataTable();
                                    da.Fill(dt);
                                    foreach (DataRow drData in dt.Rows)
                                    {
                                        string IDtext = txtScanID.Text.Trim();
                                        string IDcheck = Convert.ToString(drData["serial"]);
                                        DepartmentSampleMulti = Convert.ToString(drData["dept"]);

                                        int min = 999999;
                                        int max = 0;

                                        Array.Sort(B);
                                        Array.Reverse(B);
                                        Array.Sort(A);
                                        Array.Reverse(A);

                                        int i = 1;
                                        foreach (var r in B)
                                        {
                                            int recieve = Convert.ToInt32(r);
                                            C[i] = recieve;
                                            i++;
                                        }

                                        int p = 0;
                                        foreach (var h in B)
                                        {
                                            D[p] = h;
                                            string sett = "EIA" + D[p];
                                            Connection();
                                            com = "update set_data_update set DOC_NO = '" + DocumentNumber + "' where set_no ='" + sett + "' and flg_status = '1'";
                                            MySqlCommand cm = new MySqlCommand(com, conn);
                                            cm.ExecuteNonQuery();
                                            closeCon();
                                            p++;
                                        }

                                        for (int s = 0; s <= 17; s++)
                                        {

                                            string one = A[s];
                                            var EW = from S in C
                                                     where S != 0
                                                     //where S.ToString().Substring(0, 2) == one
                                                     select S;

                                            int K = EW.Count();

                                            for (int r = 1; r <= K; r++)
                                            {
                                                if ((C[r - 1] - C[r]) == 1 || r == 1)
                                                {
                                                    int bbb = C[r];
                                                    //aaa = Convert.ToString(bbb);
                                                    N = bbb.ToString("000000"); //003310
                                                    if (N != "")
                                                    {
                                                        //string N = "00" + aaa;
                                                        string check_eia = "EIA" + N;
                                                        foreach (DataRow dr in dt2.Rows)
                                                        {
                                                            eiaBor = Convert.ToString(dr["SET_NO"]);
                                                            if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                            {
                                                                if (N.Substring(0, 4) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                                {
                                                                    modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                                    int int_cut = Convert.ToInt32(N); //convert aaa string to int aaa = 3310

                                                                    if (min > int_cut) // 9999 > 3311,3311 > 3310
                                                                    {
                                                                        min = int_cut; //3311,3310
                                                                    }
                                                                    if (max < int_cut)
                                                                    {
                                                                        max = int_cut;

                                                                    }

                                                                    EIAFirst = min.ToString("000000");
                                                                    EIALast = max.ToString("000000");
                                                                    if (EIAFirst == EIALast)
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst;
                                                                        Unit = "Set";
                                                                    }
                                                                    else
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst + "-" + "EIA" + EIALast;
                                                                        Unit = "Set";
                                                                    }

                                                                    DocumentNoSampleMulti = "ETC";
                                                                    CountBorrow2++;
                                                                }

                                                            }

                                                        }
                                                    }

                                                }
                                                else if ((C[r - 1] - C[r]) != 1 && r != 1)
                                                {

                                                    if (r != 0 && one == N.Substring(0, 4))
                                                    {
                                                        newdata.Rows.Add(Unit, DocumentNumber, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, modelBor, EIARange, CountBorrow2, "Internal", DateBorrow);
                                                        min = 999999;
                                                        max = 0;
                                                        CountBorrow2 = 0;
                                                        EIARange = "";

                                                    }

                                                    int bbb = C[r];
                                                    //aaa = Convert.ToString(bbb);
                                                    N = bbb.ToString("000000"); //003310
                                                    if (N != "")
                                                    {
                                                        //string N = "00" + aaa;
                                                        string check_eia = "EIA" + N;
                                                        foreach (DataRow dr in dt2.Rows)
                                                        {
                                                            eiaBor = Convert.ToString(dr["SET_NO"]);
                                                            if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                            {
                                                                if (N.Substring(0, 4) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                                {
                                                                    modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                                    int int_cut = Convert.ToInt32(N); //convert aaa string to int aaa = 3310

                                                                    if (min > int_cut) // 9999 > 3311,3311 > 3310
                                                                    {
                                                                        min = int_cut; //3311,3310
                                                                    }
                                                                    if (max < int_cut)
                                                                    {
                                                                        max = int_cut;
                                                                    }

                                                                    EIAFirst = min.ToString("000000");
                                                                    EIALast = max.ToString("000000");
                                                                    if (EIAFirst == EIALast)
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst;
                                                                        Unit = "Set";
                                                                    }
                                                                    else
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst + "-" + "EIA" + EIALast;
                                                                        Unit = "Set";
                                                                    }

                                                                    DocumentNoSampleMulti = "ETC";
                                                                    CountBorrow2++;
                                                                }

                                                            }

                                                        }
                                                    }

                                                }

                                            }

                                            if (one != "")
                                            {
                                                if (EIARange != "")
                                                {
                                                    newdata.Rows.Add(Unit, DocumentNumber, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, modelBor, EIARange, CountBorrow2, "Internal", DateBorrow);
                                                    min = 999999;
                                                    max = 0;
                                                    CountBorrow2 = 0;
                                                    one = "";
                                                }
                                            }


                                        }
                                        l = 0;
                                        n = 0;
                                        k = 0;
                                        l = 0;
                                        n = 0;
                                        k = 0;
                                        getModel = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                                        A = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                                        B = new string[500];
                                        C = new int[501];
                                        D = new string[501];

                                        frmPdfExport expPDF = new frmPdfExport();
                                        expPDF.select_db = b;
                                        expPDF.exportPDF(newdata);

                                    }
                                }
                                else if (checkBoxExternal.Checked == true)
                                {

                                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                                    string IP = iniconfig.IniReadValue("Employee_Server", "IP");
                                    string DB = iniconfig.IniReadValue("Employee_Server", "DB");

                                    string ConnectionString = "Server=" + IP + ";";
                                    ConnectionString += "Uid=root;";
                                    ConnectionString += "Password=123456;";
                                    ConnectionString += "Database=" + DB + ";";

                                    conn = new MySqlConnection(ConnectionString);

                                    conn.Open();
                                    string strSQL = "SELECT serial,dept FROM `employee` where serial = '" + txtScanID.Text.Trim() + "'";

                                    MySqlDataAdapter da = new MySqlDataAdapter(strSQL, conn);
                                    DataTable dt = new DataTable();
                                    da.Fill(dt);
                                    foreach (DataRow drData in dt.Rows)
                                    {
                                        string IDtext = txtScanID.Text.Trim();
                                        string IDcheck = Convert.ToString(drData["serial"]);
                                        DepartmentSampleMulti = Convert.ToString(drData["dept"]);

                                        int min = 999999;
                                        int max = 0;

                                        Array.Sort(B);
                                        Array.Reverse(B);
                                        Array.Sort(A);
                                        Array.Reverse(A);

                                        int i = 1;
                                        foreach (var r in B)
                                        {
                                            int recieve = Convert.ToInt32(r);
                                            C[i] = recieve;
                                            i++;
                                        }

                                        int p = 0;
                                        foreach (var h in B)
                                        {
                                            D[p] = h;
                                            string sett = "EIA" + D[p];
                                            Connection();
                                            com = "update set_data_update set DOC_NO = '" + DocumentNumber + "' where set_no ='" + sett + "' and flg_status = '1'";
                                            MySqlCommand cm = new MySqlCommand(com, conn);
                                            cm.ExecuteNonQuery();
                                            closeCon();
                                            p++;
                                        }

                                        for (int s = 0; s <= 17; s++)
                                        {

                                            string one = A[s];
                                            var EW = from S in C
                                                     where S != 0
                                                     //where S.ToString().Substring(0, 2) == one
                                                     select S;

                                            int K = EW.Count();

                                            for (int r = 1; r <= K; r++)
                                            {
                                                if ((C[r - 1] - C[r]) == 1 || r == 1)
                                                {
                                                    int bbb = C[r];
                                                    //aaa = Convert.ToString(bbb);
                                                    N = bbb.ToString("000000"); //003310
                                                    if (N != "")
                                                    {
                                                        //string N = "00" + aaa;
                                                        string check_eia = "EIA" + N;
                                                        foreach (DataRow dr in dt2.Rows)
                                                        {
                                                            eiaBor = Convert.ToString(dr["SET_NO"]);
                                                            if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                            {
                                                                if (N.Substring(0, 4) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                                {
                                                                    modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                                    int int_cut = Convert.ToInt32(N); //convert aaa string to int aaa = 3310

                                                                    if (min > int_cut) // 9999 > 3311,3311 > 3310
                                                                    {
                                                                        min = int_cut; //3311,3310
                                                                    }
                                                                    if (max < int_cut)
                                                                    {
                                                                        max = int_cut;

                                                                    }

                                                                    EIAFirst = min.ToString("000000");
                                                                    EIALast = max.ToString("000000");
                                                                    if (EIAFirst == EIALast)
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst;
                                                                        Unit = "Set";
                                                                    }
                                                                    else
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst + "-" + "EIA" + EIALast;
                                                                        Unit = "Set";
                                                                    }

                                                                    DocumentNoSampleMulti = "ETC";
                                                                    CountBorrow2++;
                                                                }

                                                            }

                                                        }
                                                    }

                                                }
                                                else if ((C[r - 1] - C[r]) != 1 && r != 1)
                                                {

                                                    if (r != 0 && one == N.Substring(0, 4))
                                                    {
                                                        newdata.Rows.Add(Unit, DocumentNumber, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, modelBor, EIARange, CountBorrow2, "External", DateBorrow);
                                                        min = 999999;
                                                        max = 0;
                                                        CountBorrow2 = 0;
                                                        EIARange = "";

                                                    }

                                                    int bbb = C[r];
                                                    //aaa = Convert.ToString(bbb);
                                                    N = bbb.ToString("000000"); //003310
                                                    if (N != "")
                                                    {
                                                        //string N = "00" + aaa;
                                                        string check_eia = "EIA" + N;
                                                        foreach (DataRow dr in dt2.Rows)
                                                        {
                                                            eiaBor = Convert.ToString(dr["SET_NO"]);
                                                            if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                            {
                                                                if (N.Substring(0, 4) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                                {
                                                                    modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                                    int int_cut = Convert.ToInt32(N); //convert aaa string to int aaa = 3310

                                                                    if (min > int_cut) // 9999 > 3311,3311 > 3310
                                                                    {
                                                                        min = int_cut; //3311,3310
                                                                    }
                                                                    if (max < int_cut)
                                                                    {
                                                                        max = int_cut;
                                                                    }

                                                                    EIAFirst = min.ToString("000000");
                                                                    EIALast = max.ToString("000000");
                                                                    if (EIAFirst == EIALast)
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst;
                                                                        Unit = "Set";
                                                                    }
                                                                    else
                                                                    {
                                                                        EIARange = "EIA" + EIAFirst + "-" + "EIA" + EIALast;
                                                                        Unit = "Set";
                                                                    }

                                                                    DocumentNoSampleMulti = "ETC";
                                                                    CountBorrow2++;
                                                                }

                                                            }

                                                        }
                                                    }

                                                }

                                            }

                                            if (one != "")
                                            {
                                                if (EIARange != "")
                                                {
                                                    newdata.Rows.Add(Unit, DocumentNumber, txtScanID.Text.Trim(), lblRequestName.Text, DepartmentSampleMulti, modelBor, EIARange, CountBorrow2, "External", DateBorrow);
                                                    min = 999999;
                                                    max = 0;
                                                    CountBorrow2 = 0;
                                                    one = "";
                                                }
                                            }


                                        }
                                        l = 0;
                                        n = 0;
                                        k = 0;
                                        getModel = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                                        A = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                                        B = new string[500];
                                        C = new int[501];
                                        D = new string[501];


                                        frmPdfExport expPDF = new frmPdfExport();
                                        expPDF.select_db = b;
                                        expPDF.exportPDF(newdata);

                                    }
                                }
                                getHistory("NORMAL BORROW");
                                getCount();

                                //newdata.Rows.Clear();
                                Unit = "";
                                DocumentNumber = "";
                                lblRequestName.Text = "";
                                DepartmentSampleMulti = "";
                                modelBor = "";
                                EIARange = "";
                                CountBorrow2 = 0;
                                DateBorrow = "";

                                txtScanID.Text = "";
                                txtScanID.Focus();
                                txtScanID.Select();
                                txtScanSetNo.Text = "";
                                checkBoxInternal.Checked = false;
                                checkBoxExternal.Checked = false;
                                txtScanSetNo.Enabled = false;
                                txtScanSetNo.BackColor = Color.Gray;
                                radioMultipleBorrow.Checked = false;
                                radioMultipleBorrow.ForeColor = Color.White;
                                radioDefaultBorrow.Checked = true;
                                radioDefaultBorrow.ForeColor = Color.Yellow;
                                radioIssueDocument.Checked = false;
                                radioIssueDocument.ForeColor = Color.White;
                                btnMultipleBorrow.Visible = false;
                                btnCancleBorrow.Visible = false;
                                newdata.Clear();
                                dt2.Clear();
                                newdata.Clear();
                                cmd_multiborrow = "";


                            }
                            else if (result == DialogResult.No)
                            {
                                txtScanSetNo.Focus();
                                txtScanSetNo.SelectAll();

                            }

                            getHistory("NORMAL BORROW");
                            getCount();

                            txtScanID.Text = "";
                            txtScanID.Focus();
                            txtScanID.Select();
                            txtScanSetNo.Text = "";
                            checkBoxInternal.Checked = false;
                            checkBoxExternal.Checked = false;
                            txtScanSetNo.Enabled = false;
                            txtScanSetNo.BackColor = Color.Gray;
                            radioMultipleBorrow.Checked = false;
                            radioMultipleBorrow.ForeColor = Color.White;
                            radioDefaultBorrow.Checked = true;
                            radioDefaultBorrow.ForeColor = Color.Yellow;
                            radioIssueDocument.Checked = false;
                            radioIssueDocument.ForeColor = Color.White;
                            btnMultipleBorrow.Visible = false;
                            btnCancleBorrow.Visible = false;
                            dt2.Clear();
                            newdata.Clear();
                            cmd_multiborrow = "";

                         
                        }

                    }

                }
                else if (cmd_multiborrowNotIssue != "")
                {
                    if (radioNotIssue.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Borrow item: " + Counter + "- Return Date: " + datePickerReturn.Value.ToShortDateString() + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            PleaseWait.Create();
                            try
                            {
                                Connection();
                                MySqlCommand cmd = new MySqlCommand(cmd_multiborrowNotIssue, conn);
                                cmd.ExecuteNonQuery();
                                closeCon();

                                getHistory("NORMAL BORROW");
                                getCount();
                                txtScanID.Text = "";
                                txtScanID.Focus();
                                txtScanID.Select();
                                txtScanSetNo.Text = "";
                                lblRequestName.Text = "";
                                txtScanSetNo.Enabled = false;
                                txtScanSetNo.BackColor = Color.Gray;
                                radioMultipleBorrow.Checked = true;
                                radioMultipleBorrow.ForeColor = Color.Yellow;
                                dt2.Clear();
                                t.Clear();
                                cmd_multiborrowNotIssue = "";
                            }
                            finally
                            {
                                PleaseWait.Destroy();
                            }

                        }

                    }
                }
                else if (cmd_multiReturn != "")
                {
                    DialogResult result = MessageBox.Show("Return item: " + Counter + "- By: " + lblRequestName.Text + "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        PleaseWait.Create();
                        try
                        {
                            Connection();
                            MySqlCommand cmdd = new MySqlCommand(cmd_multiReturn, conn);
                            cmdd.ExecuteNonQuery();
                            closeCon();

                            getHistory("NORMAL BORROW");
                            getCount();
                            txtScanID.Text = "";
                            txtScanID.Focus();
                            txtScanSetNo.Text = "";
                            txtScanSetNo.Enabled = false;
                            txtScanSetNo.BackColor = Color.Gray;
                            radioDefaultBorrow.Checked = true;
                            radioDefaultBorrow.ForeColor = Color.Yellow;
                            radioMultipleBorrow.Checked = false;
                            radioMultipleBorrow.ForeColor = Color.White;
                            btnCancleBorrow.Visible = false;
                            btnMultipleBorrow.Visible = false;
                            dt2.Clear();
                            cmd_multiReturn = "";
                        }
                        finally
                        {
                            PleaseWait.Destroy();
                        }

                    }

                }
                else
                {
                    MessageBox.Show("No data Borrow or Return, Please Check", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

            }

        }

        public void InsertDoc(string strDocumentNumber,string strSetNo)
        {
            Connection();
            com = "update agreement_document set DOC_NO = '" + strDocumentNumber + "' where set_no ='" + strSetNo + "' and flg_status = '1'";
        }

        public void InsertMultiple(string strSetNo, string strRequestName, string strSpacial, string strDueDate)
        {
            Connection();
            
            if (strSpacial == "")
            {
                cmd_multiborrow += "insert into set_data_update (set_no,request_name,request_date,due_date,flg_status)values('" + strSetNo + "','" + strRequestName + "',sysdate()," + strDueDate + ",'1');";
            } 
            else
            {
                cmd_multiborrow += "insert into set_data_update (set_no,request_name,request_date,duedate,flg_status,flg_spacial)values('" + strSetNo + "','" + strRequestName + "',sysdate(),'" + strDueDate + "','1','" + strSpacial + "');";
            }
            
        }


        public void InsertMultipleNotIssue(string strSetNo, string strRequestName, string strSpacial, string strDueDate)
        {
            Connection();

            if (strSpacial == "")
            {
                cmd_multiborrowNotIssue += "insert into set_data_update (set_no,request_name,request_date,due_date,flg_status,DOC_NO)values('" + strSetNo + "','" + strRequestName + "',sysdate()," + strDueDate + ",'1','Not Issue');";
            }
            else
            {
                cmd_multiborrowNotIssue += "insert into set_data_update (set_no,request_name,request_date,duedate,flg_status,flg_spacial,DOC_NO)values('" + strSetNo + "','" + strRequestName + "',sysdate(),'" + strDueDate + "','1','" + strSpacial + "','Not Issue');";
            }

        }

        public void UpdateReturnMultiBorrow(string strSetNo, string strRequestName)
        {
            Connection();

            cmd_multiReturn += "update set_data_update set return_name = '" + strRequestName + "' ,return_date = sysdate(),flg_status= '0'"
            + " where set_no = '" + strSetNo + "' and flg_status = '1';";
        }


        private void btnCancleBorrow_Click(object sender, EventArgs e)
        {
            txtScanSetNo.Focus();
            txtScanSetNo.SelectAll();
            dtGridViwerHitory.DataSource = null;
            dt2.Clear();

        }

        
        private void AgreementBorrowSet(object sender, EventArgs e)
        {
            value = Interaction.InputBox("Please enter your id [Last 8 digit]", "PASSWORD", "");

            if (value == "")
            {
                return;
            }

            if (getAuthenAdmin(value) == true)
            {
                frmAgreementBorrowSet agreement = new frmAgreementBorrowSet();
                agreement.select_db = b;
                agreement.checkname(value.ToString());
                agreement.ShowDialog();
            }
            else
            {
                MessageBox.Show("Access denied:" + value);
            }

        }

        public void checkuser(string select_user)
        {
            textuser = select_user.ToString();
        }

        private void lblOverDue_Click(object sender, EventArgs e)
        {

        }

        private void checkBoxInternal_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxInternal.Checked == true)
            {
                if (radioBorrow.Checked == true || radioReturn.Checked == true)
                {
                    if (radioDefaultBorrow.Checked == true || radioMultipleBorrow.Checked == true)
                    {
                        checkBoxExternal.Checked = false;
                        txtScanID.Text = "";
                        txtScanID.SelectAll();
                        txtScanID.Focus();

                        txtScanSetNo.Text = "";
                        
                    }
                }

            }

        }

        private void radioNotUseRFID_Click(object sender, EventArgs e)
        {
            if (radioNotUseRFID.Checked == true)
            {
                txtScanID.Focus();
                txtScanID.SelectAll();

                txtRFID.Visible = false;
                radioRFID.Checked = false;
                radioRFID.ForeColor = Color.White;
                radioNotUseRFID.ForeColor = Color.Yellow;

                lblRequestName.Text = "";
                txtScanID.Text = "";
                txtScanSetNo.Text = "";

                txtScanSetNo.Enabled = false;
                txtScanSetNo.BackColor = Color.Gray;
            }
        }

        private void radioRFID_Click(object sender, EventArgs e)
        {
            if (radioRFID.Checked == true)
            {
                lblRequestName.Text = "";
                txtScanID.Text = "";
                txtScanID.Enabled = false;
                txtScanID.BackColor = Color.Yellow;

                txtScanSetNo.Text = "";
                txtScanSetNo.Enabled = false;
                txtScanSetNo.BackColor = Color.Gray;

                txtRFID.Text = "";

                radioRFID.ForeColor = Color.Yellow;
                radioNotUseRFID.Checked = false;
                radioNotUseRFID.ForeColor = Color.White;

                txtRFID.Visible = true;
                txtRFID.Focus();
                txtRFID.SelectAll();

                //checkRFID();
            }
        }

        private void txtRFID_TextChanged(object sender, EventArgs e)
        {
            if (txtRFID.Text != "")
            {
                if (txtRFID.Text.Length == 8)
                {
                    checkRFID();

                    txtScanID.Focus();
                    txtScanID.SelectAll();
                    txtScanID.Enabled = true;
                }
            }
        }

        public void checkRFID()
        {
            txtRFID.Visible = false;
            string empRFID, empNAME, empID;

            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
            string IP = iniconfig.IniReadValue("Server", "IP");
            string DB = iniconfig.IniReadValue("Server", "DB");
            string ConnectionString = "Server='" + IP + "';";
            ConnectionString += "User ID=sa;";
            ConnectionString += "Password=s;";
            ConnectionString += "Database='" + DB + "';";
            SQLConnection = new SqlConnection(ConnectionString);
            SQLConnection.Open();
            SqlDataAdapter adt = new SqlDataAdapter("SELECT * FROM [STTC_HUMAN_RESOURCE].[dbo].[TBL_MANPOW_EMPID_RFID]", SQLConnection);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            foreach (DataRow row in dt.Rows)
            {
                empRFID = Convert.ToString(row["RFID"]);
                empNAME = Convert.ToString(row["REMARK"]);
                empID = Convert.ToString(row["EMPID"]);
                if (txtRFID.Text == empRFID)
                {

                    txtScanID.Text = empID;
                }
            }
        }

        private void checkBoxExternal_CheckedChanged(object sender, EventArgs e)
        {
           if (checkBoxExternal.Checked == true)
            {
                if (radioBorrow.Checked == true || radioReturn.Checked == true)
                {
                    if (radioDefaultBorrow.Checked == true || radioMultipleBorrow.Checked == true)
                    {
                        checkBoxInternal.Checked = false;
                        txtScanID.Text = "";
                        txtScanID.SelectAll();
                        txtScanID.Focus();

                        txtScanSetNo.Text = "";
                    }
                }
            }
        }

        private void radioIssueDocument_Click(object sender, EventArgs e)
        {
            if (radioBorrow.Checked == true)
            {
                if (radioDefaultBorrow.Checked == true)
                {
                    if (radioIssueDocument.Checked == true)
                    {
                        radioIssueDocument.ForeColor = Color.Yellow;
                        radioNotIssue.ForeColor = Color.White;

                        groupBox6.Visible = true;

                        txtScanID.Text = "";
                        txtScanID.Focus();
                        txtScanID.SelectAll();

                    }
                }
                else if (radioMultipleBorrow.Checked == true)
                {
                    if (radioIssueDocument.Checked == true)
                    {
                        radioIssueDocument.ForeColor = Color.Yellow;
                        radioNotIssue.ForeColor = Color.White;

                        groupBox6.Visible = true;

                        txtScanID.Text = "";
                        txtScanID.Focus();
                        txtScanID.SelectAll();

                    }
                }
                else
                {
                    MessageBox.Show("Please select One by One or Multiple Borrow!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if (radioReturn.Checked == true)
            {
                if (radioMultipleBorrow.Checked == true)
                {
                    if (radioIssueDocument.Checked == true)
                    {
                        radioIssueDocument.ForeColor = Color.Yellow;
                        radioNotIssue.ForeColor = Color.White;

                        groupBox6.Visible = true;

                        txtScanID.Text = "";
                        txtScanID.Focus();
                        txtScanID.SelectAll();

                    }
                }
                else if (radioReturn.Checked == true)
                {
                    if (radioIssueDocument.Checked == true)
                    {
                        radioIssueDocument.ForeColor = Color.Yellow;
                        radioNotIssue.ForeColor = Color.White;

                        groupBox6.Visible = true;

                        txtScanID.Text = "";
                        txtScanID.Focus();
                        txtScanID.SelectAll();

                    }
                }
            }
            else
            {
                MessageBox.Show("Please Select Borrow or Return!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void radioNotIssue_Click(object sender, EventArgs e)
        {
            if (radioBorrow.Checked == true)
            {
                if (radioDefaultBorrow.Checked == true)
                {
                    if (radioNotIssue.Checked == true)
                    {
                        radioIssueDocument.ForeColor = Color.White;
                        radioNotIssue.ForeColor = Color.Yellow;

                        groupBox6.Visible = false;
                        checkBoxInternal.Checked = false;
                        checkBoxExternal.Checked = false;

                        txtScanSetNo.Enabled = false;
                        txtScanSetNo.BackColor = Color.Gray;
                        txtScanSetNo.Text = "";

                        txtScanID.Enabled = true;
                        txtScanID.BackColor = Color.Yellow;
                        txtScanID.Text = "";

                        txtScanID.Text = "";
                        txtScanID.Focus();
                        txtScanID.SelectAll();

                        //txtScanSetNo.ForeColor = Color.Gray;
                    }
                }
                else if (radioMultipleBorrow.Checked == true)
                {
                    if (radioNotIssue.Checked == true)
                    {
                        radioIssueDocument.ForeColor = Color.White;
                        radioNotIssue.ForeColor = Color.Yellow;

                        groupBox6.Visible = false;
                        checkBoxInternal.Checked = false;
                        checkBoxExternal.Checked = false;

                        txtScanSetNo.Enabled = false;
                        txtScanSetNo.BackColor = Color.Gray;
                        txtScanSetNo.Text = "";

                        txtScanID.Enabled = true;
                        txtScanID.BackColor = Color.Yellow;
                        txtScanID.Text = "";
                        txtScanID.Focus();
                        txtScanID.SelectAll();
                        //txtScanSetNo.ForeColor = Color.Gray;
                    }
                }
                else
                {
                    MessageBox.Show("Please select One by One or Multiple Borrow!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if (radioReturn.Checked == true)
            {
                if (radioDefaultBorrow.Checked == true)
                {
                    if (radioNotIssue.Checked == true)
                    {
                        radioIssueDocument.ForeColor = Color.White;
                        radioNotIssue.ForeColor = Color.Yellow;

                        groupBox6.Visible = false;
                        checkBoxInternal.Checked = false;
                        checkBoxExternal.Checked = false;

                        txtScanID.Text = "";
                        txtScanID.Focus();
                        txtScanID.SelectAll();
                    }
                }
                else if (radioMultipleBorrow.Checked == true)
                {
                    if (radioNotIssue.Checked == true)
                    {
                        radioIssueDocument.ForeColor = Color.White;
                        radioNotIssue.ForeColor = Color.Yellow;

                        groupBox6.Visible = false;
                        checkBoxInternal.Checked = false;
                        checkBoxExternal.Checked = false;

                        txtScanID.Text = "";
                        txtScanID.Focus();
                        txtScanID.SelectAll();
                    }
                }
                else
                {
                    MessageBox.Show("Please select One by One or Multiple Borrow!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            else
            {
                MessageBox.Show("Please Select Borrow or Return!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        public void datecheck()
        {
            IniFile Gen;
            Gen = new IniFile(Application.StartupPath + "\\generate.ini");
            DateTime dd = DateTime.Now;
            string dateDoc = dd.ToString("yyyyMM");

            if (dateDoc == Gen.IniReadValue("generate", "date"))
            {

            }
            else
            {
                Gen.IniWriteValue("generate", "date", dateDoc);
                Gen.IniWriteValue("generate", "gen_ssc", "0");
                Gen.IniWriteValue("generate", "gen_etc", "0");
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBoxRestore_Click(object sender, EventArgs e)
        {

        }


    }
}
