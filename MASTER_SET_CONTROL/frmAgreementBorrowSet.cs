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
using System.Data.SqlClient;
using INI;
using Microsoft.VisualBasic;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using System.Web;
using System.Net;
using System.Diagnostics;


namespace MASTER_SET_CONTROL
{
    public partial class frmAgreementBorrowSet : Form
    {
        MySqlConnection conn;
        //MySqlConnection connect;
        //SqlConnection SQLConnection;
        //IniFile iniconfig;
        IniFile iniconfig;
        SqlConnection SQLConnection;

        //WebRequest Rd;
        //WebResponse Rp;

        BindingSource bs = new BindingSource();
        BindingList<DataTable> tables = new BindingList<DataTable>();

        DataTable newdata = new DataTable();
        DataTable dtBorrow = new DataTable();
        DataTable dtMaster = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable t = new DataTable();

        public string select_db;
        string empID;
        string strID;
        string strPassword;
        string b;
        string strCon;

        string EIALast;
        string EIARange;
        string Unit;
        int CountBorrow2 = 0;
        string eiaBor;
        string modelBor;
        string eia;
        string check_cut;
        string checkcutA;
        int l;
        
        public frmAgreementBorrowSet()
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

        private void frmAgreementBorrowSet_Load(object sender, EventArgs e)
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            if (select_db == "SAMPLE SET CONTROL")
            {
                b = select_db;

                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                strCon = "Server=" + IP + ";";
                strCon += "Uid=root;";
                strCon += "Password=123456*;";
                strCon += "Database=" + DB + ";";

                conn = new MySqlConnection(strCon);

                conn.Open();

                //ConnectionString = "Server=localhost;Database=eia_master_set_control;Uid=root;Password=123456*;";
                //strCon = "host=localhost;Database=eia_master_set_control;Uid=root;Password=123456;Convert Zero Datetime=True;";
                this.Text = "SAMPLE SET CONTROL FORM AGREEMENT BORROW SET" + String.Format(" --- Version {0}", version) + " - Server : " + IP;
            }
            else if (select_db == "ENGINEERING TRAINING CENTER")
            {
                b = select_db;
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB = iniconfig.IniReadValue("MySQL_Server", "DB2");

                strCon = "Server=" + IP + ";";
                strCon += "Uid=root;";
                strCon += "Password=123456*;";
                strCon += "Database=" + DB + ";";

                conn = new MySqlConnection(strCon);

                conn.Open();

                //ConnectionString2 = "Server=localhost;Database=eia_master_set_control_spacial;Uid=root;Password=123456;";
                //strCon = "host=localhost;Database=eia_master_set_control_spacial;Uid=root;Password=123456;Convert Zero Datetime=True;";
                this.Text = "ENGINEERING TRAINING CENTER AGREEMENT BORROW SET" + String.Format(" --- Version {0}", version) + " - Server : " + IP;
            }

           
            comboSearchType.Items.Add("Pending");
            comboSearchType.Items.Add("Complete");

            comboStatusReprint.Items.Add("Pending");
            comboStatusReprint.Items.Add("Complete");

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

            dtBorrow.Columns.Add("SET_NO");

            dtMaster.Columns.Add("SET_NO");
            dtMaster.Columns.Add("MODEL");

            dt2.Columns.Add("SET_NO");
            dt2.Columns.Add("MODEL");

            strID = checkCredential.username;
            strPassword = checkCredential.password;

            PleaseWait.Create();
            try
            {
                getMaster("NOTCOMPLETE");
                getCount("NOTCOMPLETE");
                getCount("COMPLETE");
                getCount("MONTH");
                getCount("TODAY");
                getCount("TOTAL");
            }
            finally
            {
                PleaseWait.Destroy();
            }
           
        }


        string Namee;
        string default_empID;
        public void checkname(string name)
        {
            empID = name;

            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
            string IP = iniconfig.IniReadValue("Server", "IP");
            string DB = iniconfig.IniReadValue("Server", "DB");

            string ConnectionString = "Server='" + IP + "';";
            ConnectionString += "User ID=sa;";
            ConnectionString += "Password=s;";
            ConnectionString += "Database='" + DB + "';";

            SQLConnection = new SqlConnection(ConnectionString);

            SQLConnection.Open();
            string strSQL = "SELECT * FROM [STTC_HUMAN_RESOURCE].[dbo].[TBL_MANPOW_EMPID_RFID] where EMPID = '" + empID + "' ";

            SqlDataAdapter da = new SqlDataAdapter(strSQL, SQLConnection);
            DataTable dt = new DataTable();
            da.Fill(dt);

            foreach (DataRow row in dt.Rows)
            {
                default_empID = Convert.ToString(row["EMPID"]);
                if (empID == default_empID)
                {
                    Namee = Convert.ToString(row["Remark"]);
                }
               
            }

        }

        public void closeCon()
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }

        }
        
        public void getStatusDocument()
        {
            if ((string)comboStatusReprint.SelectedItem == "Pending")
            {
                comboSearchReprint.Items.Clear();
                comboSearchByReprint.Items.Clear();
                txtSearchReprint.Text = "";

                comboSearchReprint.Items.Add("TODAY");
                comboSearchReprint.Items.Add("MONTH");
                comboSearchReprint.Items.Add("TOTAL");
            }
            else if ((string)comboStatusReprint.SelectedItem == "Complete")
            {
                comboSearchReprint.Items.Clear();
                comboSearchByReprint.Items.Clear();
                txtSearchReprint.Text = "";

                comboSearchReprint.Items.Add("TODAY");
                comboSearchReprint.Items.Add("MONTH");
                comboSearchReprint.Items.Add("TOTAL");
            }
        }

        public void getSearchReprint()
        {
            if ((string)comboSearchReprint.SelectedItem == "TODAY")
            {
                comboSearchByReprint.Items.Clear();

                Connection();
                //MySqlDataAdapter adapter = new MySqlDataAdapter("select distinct agreement_document.AUTO_ID,agreement_document.DOC_NO,agreement_document.PERSONAL_ID,PERSONAL_NAME,agreement_document.DEPARTMENT,agreement_document.BORROW_DATE,agreement_document.PURPOSE,agreement_document.COMPLETE_BY,agreement_document.COMPLETE_DATE,agreement_document.flg_status,set_data_update.flg_status from agreement_document,set_data_update where set_data_update.DOC_NO = agreement_document.DOC_NO and set_data_update.flg_status = '1' order by rec_date desc", conn);
                MySqlDataAdapter adapter = new MySqlDataAdapter("select DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where BORROW_DATE = CURRENT_DATE() order by BORROW_DATE desc", conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                foreach (DataColumn col in dt.Columns)
                {
                    comboSearchByReprint.Items.Add(col.ColumnName);
                }
                closeCon();
                comboSearchByReprint.Items.Add("MODEL");

            }
            else if ((string)comboSearchReprint.SelectedItem == "MONTH")
            {
                comboSearchByReprint.Items.Clear();

                Connection();
                MySqlDataAdapter adapter = new MySqlDataAdapter("select DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where MONTH(BORROW_DATE) = MONTH(CURRENT_DATE()) order by BORROW_DATE desc", conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                foreach (DataColumn col in dt.Columns)
                {
                    comboSearchByReprint.Items.Add(col.ColumnName);
                }
                closeCon();
                comboSearchByReprint.Items.Add("MODEL");

            }
            else if ((string)comboSearchReprint.SelectedItem == "TOTAL")
            {
                comboSearchByReprint.Items.Clear();

                Connection();
                MySqlDataAdapter adapter = new MySqlDataAdapter("select DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document order by BORROW_DATE desc", conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                foreach (DataColumn col in dt.Columns)
                {
                    comboSearchByReprint.Items.Add(col.ColumnName);
                }
                closeCon();
                comboSearchByReprint.Items.Add("MODEL");

            }
        }

        public void getSearchMaster()
        {
            if ((string)comboStatusReprint.SelectedItem == "Pending")
            {
                if ((string)comboSearchByReprint.SelectedItem == "DOC_NO")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where DOC_NO like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "PERSONAL_ID")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PERSONAL_ID like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "PERSONAL_NAME")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PERSONAL_NAME like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "DEPARTMENT")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where DEPARTMENT like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "PURPOSE")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PURPOSE like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "COMPLETE_BY")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where COMPLETE_BY like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "COMPLETE_DATE")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where COMPLETE_DATE like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "MODEL")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select distinct agreement_document.AUTO_ID,agreement_document.DOC_NO,v_agreement_info.MODEL,agreement_document.PERSONAL_ID,agreement_document.PERSONAL_NAME,agreement_document.DEPARTMENT,agreement_document.BORROW_DATE,agreement_document.PURPOSE,agreement_document.COMPLETE_BY,agreement_document.COMPLETE_DATE,v_agreement_info.Flg_Doc,agreement_document.REMARK from agreement_document,v_agreement_info where v_agreement_info.MODEL = '" + txtSearchReprint.Text.Trim() + "' and v_agreement_info.Flg_Doc = 'Pending'", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
            }
            else if ((string)comboStatusReprint.SelectedItem == "Complete")
            {
                if ((string)comboSearchByReprint.SelectedItem == "DOC_NO")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where DOC_NO like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "PERSONAL_ID")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PERSONAL_ID like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "PERSONAL_NAME")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PERSONAL_NAME like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "DEPARTMENT")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where DEPARTMENT like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "PURPOSE")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PURPOSE like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "COMPLETE_BY")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where COMPLETE_BY like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "COMPLETE_DATE")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where COMPLETE_DATE like '%" + txtSearchReprint.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchByReprint.SelectedItem == "MODEL")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select distinct agreement_document.AUTO_ID,agreement_document.DOC_NO,v_agreement_info.MODEL,agreement_document.PERSONAL_ID,agreement_document.PERSONAL_NAME,agreement_document.DEPARTMENT,agreement_document.BORROW_DATE,agreement_document.PURPOSE,agreement_document.COMPLETE_BY,agreement_document.COMPLETE_DATE,v_agreement_info.Flg_Doc,agreement_document.REMARK from agreement_document,v_agreement_info where v_agreement_info.MODEL = '" + txtSearchReprint.Text.Trim() + "' and v_agreement_info.Flg_Doc = 'Complete'", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
            }
            
            
        }

        public void getMaster(string master)
        {
            dtGridViewAgreement.DataSource = null;
            dtGridViewAgreement.Rows.Clear();
            dtGridViewAgreement.Refresh();
            bindingNavigator.BindingSource = null;
            tables.Clear();
            t.Clear();

            string strSQLL = "";
            if (master == "TOTAL")
            {
                Connection();
                strSQLL = "select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document order by rec_date desc";
                MySqlDataAdapter adtt = new MySqlDataAdapter(strSQLL, conn);
                DataTable tb = new DataTable();
                adtt.Fill(tb);
                dtGridViewAgreement.DataSource = tb;
                conn.Close();
            }
            else if (master == "TODAY")
            {
                Connection();
                strSQLL = "select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where BORROW_DATE = CURRENT_DATE() order by BORROW_DATE desc";
                MySqlDataAdapter adapp = new MySqlDataAdapter(strSQLL, conn);
                DataTable table = new DataTable();
                adapp.Fill(table);
                dtGridViewAgreement.DataSource = table;
                conn.Close();
            }
            else if (master == "MONTH")
            {
                Connection();
                strSQLL = "select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where MONTH(BORROW_DATE) = MONTH(CURRENT_DATE()) order by BORROW_DATE desc";
                MySqlDataAdapter adapterr = new MySqlDataAdapter(strSQLL, conn);
                DataTable tab = new DataTable();
                adapterr.Fill(tab);
                dtGridViewAgreement.DataSource = tab;
                conn.Close();
            }
            else if (master == "COMPLETE")
            {
                Connection();
                strSQLL = "select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where flg_status = 'Complete' order by BORROW_DATE desc";
                MySqlDataAdapter adapterr = new MySqlDataAdapter(strSQLL, conn);
                DataTable tab = new DataTable();
                adapterr.Fill(tab);
                dtGridViewAgreement.DataSource = tab;
                conn.Close();

            }
            else if (master == "NOTCOMPLETE")
            {
                Connection();
                strSQLL = "SELECT AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK FROM agreement_document where flg_status = 'Pending' order by BORROW_DATE desc";
                MySqlDataAdapter adapterr = new MySqlDataAdapter(strSQLL, conn);
                DataTable tab = new DataTable();
                adapterr.Fill(tab);
                dtGridViewAgreement.DataSource = tab;
                conn.Close();
                
            }

            // 1
            // Open connection
            Connection();
            {

                // 2
                // Create new DataAdapter
                using (MySqlDataAdapter da = new MySqlDataAdapter(
                    strSQLL, conn))
                {
                    // 3
                    // Use DataAdapter to fill DataTable
                    da.Fill(t);
                    // 4
                    // Render data onto the screen
                    dtGridViewAgreement.DataSource = t;

                    if (t.Rows.Count != 0)
                    {
                        int countt = 0;
                        DataTable dt = null;

                        foreach (DataRow dr in t.Rows)
                        {

                            if (countt == 0)
                            {
                                dt = t.Clone();
                                tables.Add(dt);
                            }
                            dt.Rows.Add(dr.ItemArray);
                            countt++;
                            if (countt >= 50)
                            {
                                countt = 0;
                            }
                            //dt.Rows.Clear();
                            dtGridViewAgreement.DataSource = null;
                        }

                        bindingNavigator.BindingSource = bs;
                        bs.DataSource = tables;
                        bs.PositionChanged += bs_PositionChanged;
                        bs_PositionChanged(bs, EventArgs.Empty);

                    }

                }
            }
        }

        public void searchtype()
        {
            if ((string)comboSearchType.SelectedItem == "Pending")
            {
        
                Connection();
                MySqlDataAdapter adapter = new MySqlDataAdapter("select DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document", conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                foreach (DataColumn col in dt.Columns)
                {
                    comboSearchBy.Items.Add(col.ColumnName);
                }
                closeCon();
            }
            else if ((string)comboSearchType.SelectedItem == "Complete")
            {
                comboSearchBy.Items.Clear();
                txtSearch.Text = "";

                Connection();
                MySqlDataAdapter adapter = new MySqlDataAdapter("select DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document", conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                foreach (DataColumn col in dt.Columns)
                {
                    comboSearchBy.Items.Add(col.ColumnName);
                }
                closeCon();
            }
        }

        public void searchby()
        {

            if ((string)comboSearchType.SelectedItem == "Pending")
            {
                if ((string)comboSearchBy.SelectedItem == "DOC_NO")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where DOC_NO like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "PERSONAL_ID")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PERSONAL_ID like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "PERSONAL_NAME")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PERSONAL_NAME like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "DEPARTMENT")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where DEPARTMENT like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "PURPOSE")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PURPOSE like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "COMPLETE_BY")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where COMPLETE_BY like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "COMPLETE_DATE")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where COMPLETE_DATE like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Pending' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
            }
            else if ((string)comboSearchType.SelectedItem == "Complete")
            {
                if ((string)comboSearchBy.SelectedItem == "DOC_NO")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where DOC_NO like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "PERSONAL_ID")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PERSONAL_ID like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "PERSONAL_NAME")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PERSONAL_NAME like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "DEPARTMENT")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where DEPARTMENT like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "PURPOSE")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where PURPOSE like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "COMPLETE_BY")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where COMPLETE_BY like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
                else if ((string)comboSearchBy.SelectedItem == "COMPLETE_DATE")
                {
                    MySqlDataAdapter adap = new MySqlDataAdapter("select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where COMPLETE_DATE like '%" + txtSearch.Text.Trim() + "%' and flg_status = 'Complete' order by BORROW_DATE desc", conn);
                    DataTable table = new DataTable();
                    adap.Fill(table);
                    dtGridViewAgreement.DataSource = table;
                }
            }
        }

        public void getCount(string count)
        {

            string strSQL = "";
            if (count == "TOTAL")
            {
                Connection();
                strSQL = "select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document order by BORROW_DATE desc";
                MySqlDataAdapter adtt = new MySqlDataAdapter(strSQL, conn);
                DataTable tb = new DataTable();
                adtt.Fill(tb);

                foreach (DataRow row in tb.Rows)
                {
                    lblTotal.Text = tb.Rows.Count.ToString();
                }
                conn.Close();
            }
            else if (count == "TODAY")
            {
                Connection();
                strSQL = "select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where BORROW_DATE = CURRENT_DATE() order by BORROW_DATE desc";
                MySqlDataAdapter adapp = new MySqlDataAdapter(strSQL, conn);
                DataTable table = new DataTable();
                adapp.Fill(table);

                foreach (DataRow rw in table.Rows)
                {
                    lblToday.Text = table.Rows.Count.ToString();
                }
                conn.Close();

            }
            else if (count == "MONTH")
            {
                Connection();
                strSQL = "select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where MONTH(BORROW_DATE) = MONTH(CURRENT_DATE()) order by BORROW_DATE desc";
                MySqlDataAdapter adapterr = new MySqlDataAdapter(strSQL, conn);
                DataTable tab = new DataTable();
                adapterr.Fill(tab);

                foreach (DataRow dr in tab.Rows)
                {
                    lblMonth.Text = tab.Rows.Count.ToString();
                }
                conn.Close();
            }
            else if (count == "COMPLETE")
            {
                Connection();
                strSQL = "select AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK from agreement_document where flg_status = 'Complete' order by BORROW_DATE desc";
                MySqlDataAdapter adapterr = new MySqlDataAdapter(strSQL, conn);
                DataTable tab = new DataTable();
                adapterr.Fill(tab);

                foreach (DataRow dr in tab.Rows)
                {
                    lblComplete.Text = tab.Rows.Count.ToString();
                }
                conn.Close();

            }
            else if (count == "NOTCOMPLETE")
            {
                Connection();
                strSQL = "SELECT AUTO_ID,DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,COMPLETE_BY,COMPLETE_DATE,flg_status,REMARK FROM agreement_document where flg_status = 'Pending' ";
                MySqlDataAdapter adapterr = new MySqlDataAdapter(strSQL, conn);
                DataTable tab = new DataTable();
                adapterr.Fill(tab);

                lblNotComplete.Text = tab.Rows.Count.ToString();
                
                closeCon();

                if (Int32.Parse(lblNotComplete.Text) > 0)
                {
                    TimeToday.Enabled = true;
                }
                else
                {
                    TimeToday.Enabled = false;
                    lblNotComplete.BackColor = Color.Orange;
                }
                
            }
   
        }

        void bs_PositionChanged(object sender, EventArgs e)
        {
            if (tables.Count != 0)
            {
                dtGridViewAgreement.DataSource = tables[bs.Position];
            }
        }

        public void getReprint()
        {
            Connection();
            string strSQLReprint = "select * from agreement_document";
            MySqlDataAdapter adapterr = new MySqlDataAdapter(strSQLReprint, conn);
            DataTable tab = new DataTable();
            adapterr.Fill(tab);
            dtGridViewAgreement.DataSource = tab;

        }

        private void lblTODAY(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster("TODAY");
                getCount("TODAY");

                txtDocumentNo.Text = "";
                txtPersonalID.Text = "";
                txtPersonalName.Text = "";
                txtDepartment.Text = "";
                txtBorrowDate.Text = "";
                txtPurpose.Text = "";
                txtBrowse.Text = "";
                txtComplete.Text = "";
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void lblMONTH(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster("MONTH");
                getCount("MONTH");

                txtDocumentNo.Text = "";
                txtPersonalID.Text = "";
                txtPersonalName.Text = "";
                txtDepartment.Text = "";
                txtBorrowDate.Text = "";
                txtPurpose.Text = "";
                txtBrowse.Text = "";
                txtComplete.Text = "";
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void lblTOTAL(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster("TOTAL");
                getCount("TOTAL");

                txtDocumentNo.Text = "";
                txtPersonalID.Text = "";
                txtPersonalName.Text = "";
                txtDepartment.Text = "";
                txtBorrowDate.Text = "";
                txtPurpose.Text = "";
                txtBrowse.Text = "";
                txtComplete.Text = "";
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }


        private void TimeToday_Tick(object sender, EventArgs e)
        {
            if (lblNotComplete.BackColor == Color.Red)
            {
                lblNotComplete.BackColor = Color.Orange;
            }
            else
            {
                lblNotComplete.BackColor = Color.Red;
            }
        }

        string docno;
        private void dtGridViewAgreement_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dtGridViewAgreement.SelectedCells.Count > 0 & e.ColumnIndex >= 0 & e.RowIndex >= 0)
            {
                dtGridViewAgreement.CurrentCell = dtGridViewAgreement.Rows[e.RowIndex].Cells[e.ColumnIndex];
                dtGridViewAgreement.Rows[e.RowIndex].Selected = true;
                dtGridViewAgreement.Focus();

                int selectedrowindex = dtGridViewAgreement.SelectedCells[0].RowIndex;
                DataGridViewRow selectedrow = dtGridViewAgreement.Rows[selectedrowindex];

                txtDocumentNo.Text = Convert.ToString(selectedrow.Cells[1].Value.ToString());
                txtPersonalID.Text = Convert.ToString(selectedrow.Cells[2].Value.ToString());
                txtPersonalName.Text = Convert.ToString(selectedrow.Cells[3].Value.ToString());
                txtDepartment.Text = Convert.ToString(selectedrow.Cells[4].Value.ToString());
                txtBorrowDate.Text = Convert.ToString(selectedrow.Cells[5].Value.ToString());
                txtPurpose.Text = Convert.ToString(selectedrow.Cells[6].Value.ToString());
                txtComplete.Text = Convert.ToString(selectedrow.Cells[7].Value.ToString());
                txtCompleteDate.Text = Convert.ToString(selectedrow.Cells[8].Value.ToString());
                txtStatus.Text = Convert.ToString(selectedrow.Cells[9].Value.ToString());

                txtDocNoReprint.Text = Convert.ToString(selectedrow.Cells[1].Value.ToString());
                txtPersonalIDReprint.Text = Convert.ToString(selectedrow.Cells[2].Value.ToString());
                txtPersonalNameReprint.Text = Convert.ToString(selectedrow.Cells[3].Value.ToString());
                txtDepartmentReprint.Text = Convert.ToString(selectedrow.Cells[4].Value.ToString());
                txtBorrowDateReprint.Text = Convert.ToString(selectedrow.Cells[5].Value.ToString());
                txtPurposeReprint.Text = Convert.ToString(selectedrow.Cells[6].Value.ToString());
                txtCompleteByReprint.Text = Convert.ToString(selectedrow.Cells[7].Value.ToString());
                txtCompleteDateReprint.Text = Convert.ToString(selectedrow.Cells[8].Value.ToString());
                txtStatusReprint.Text = Convert.ToString(selectedrow.Cells[9].Value.ToString());


                docno = Convert.ToString(selectedrow.Cells[1].Value.ToString());

            }
            
        }

        private void lblComplete_DoubleClick(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster("COMPLETE");
                getCount("COMPLETE");
            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        private void lblNotComplete_DoubleClick(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster("NOTCOMPLETE");
                getCount("NOTCOMPLETE");
            }
            finally
            {
                PleaseWait.Destroy();
            }

        }


        public void InsertAgreementUpdate(string DocumentNo, string PersonalID, string PersonalName, string Department, string BorrowDate, string QTY, string UMO, string Purpose, string SetNo, string Model, string CompleteBy)
        {
            Connection();
            string command = "";
            command = "insert into agreement_borrow_set_update (DocumentNo,personal_id,personal_name,dept,Borrow_Date,QTY,UMO,purpose,set_no,model,complete_by,flg_status)values('" + DocumentNo + "','" + PersonalID + "','" + PersonalName + "','" + Department + "','" + BorrowDate + "','" + QTY + "','" + UMO + "','" + Purpose + "','" + SetNo + "','" + Model + "','" + CompleteBy + "','Complete')";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }

        public void InsertAgreementHistory(string DocumentNo,string PersonalID,string PersonalName,string Department, string BorrowDate, string QTY, string UMO, string Purpose, string SetNo, string Model,string CompleteBy)
        {
            Connection();
            string command = "";
            command = "insert into agreement_borrow_set_history (DocumentNo,personal_id,personal_name,dept,Borrow_Date,QTY,UMO,purpose,set_no,model,complete_by,flg_status)values('" + DocumentNo + "','" + PersonalID + "','" + PersonalName + "','" + Department + "','"+BorrowDate+"','"+QTY+"','"+UMO+"','"+Purpose+"','"+SetNo+"','"+Model+"','"+CompleteBy+"','Complete')";

            MySqlCommand cmd = new MySqlCommand(command,conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }

        public void UpdateAgreementUpdate(string DocumentNo, string PersonalID, string PersonalName, string Department, string BorrowDate, string QTY, string UMO, string Purpose, string SetNo, string Model, string CompleteBy)
        {
            Connection();
            string command = "";

            command = "update agreement_borrow_set_update set DocumentNo = '" + DocumentNo + "' ,personal_id = '" + PersonalID + "',personal_name= '" + PersonalName + "',dept= '" + Department + "',Borrow_Date='" + BorrowDate + "',QTY='" + QTY + "',UMO='" + UMO + "',purpose='" + Purpose + "',set_no='" + SetNo + "',model='" + Model + "',complete_by='" + CompleteBy + "',flg_status='Complete'" + " where set_no = '"+txtDocumentNo+"'";

            MySqlCommand cmd = new MySqlCommand(command, conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }

        public bool checkDocumentNumber(string strDocument)
        {
            Connection();
            string strCheckDoc = "select distinct DOC_NO, flg_status from set_data_update where DOC_NO = '"+txtDocumentNo.Text.Trim()+"' and set_data_update.flg_status = '0'";
            MySqlDataAdapter adap = new MySqlDataAdapter(strCheckDoc, conn);
            DataTable tb = new DataTable();
            adap.Fill(tb);

            foreach (DataRow row in tb.Rows)
            {
                return true;
            }
            closeCon();

            return false;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (checkDocumentNumber(txtDocumentNo.Text.Trim()) == true)
            {
                if (txtComplete.Text.Trim() == "" | txtCompleteDate.Text.Trim() == "")
                {
                    if (txtDocumentNo.Text.Trim() != "" && txtPersonalID.Text.Trim() != "" && txtPersonalName.Text.Trim() != "")
                    {
                        if (docno != "")
                        {
                            //upload file complete
                            try
                            {
                                string newURLs = "http://43.72.52.12/uploadfile/upload_pdf.php?docno=" + docno + "&name=" + Namee;
                                System.Diagnostics.Process.Start("IExplore.exe", newURLs);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }

                    }
                    else
                    {
                        MessageBox.Show("Please Select Cell.", "Warrning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Please Select Document Number status Pending!", "Warrning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            else
            {
                MessageBox.Show("Can Not Upload File, Please Return all Set Number in Document Number : " +txtDocumentNo.Text , "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            PleaseWait.Create();
            try
            {
                getMaster("NOTCOMPLETE");
                getCount("TOTAL");
                getCount("TODAY");
                getCount("MONTH");
                getCount("COMPLETE");
                getCount("NOTCOMPLETE");

                comboSearchBy.Text = "";
                comboSearchType.Text = "";
                txtSearch.Text = "";

                txtDocumentNo.Text = "";
                txtPersonalID.Text = "";
                txtPersonalName.Text = "";
                txtDepartment.Text = "";
                txtBorrowDate.Text = "";
                txtPurpose.Text = "";
                txtBrowse.Text = "";
                txtComplete.Text = "";
                txtCompleteDate.Text = "";

            }
            finally
            {
                PleaseWait.Destroy();
            }
        }

        string strBrowse;
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (txtDocumentNo.Text.Trim() != "" && txtPersonalID.Text.Trim() != "" && txtPersonalName.Text.Trim() != "")
            {
                OpenFileDialog fdlg = new OpenFileDialog();
                fdlg.Filter = "PDF Files (*.pdf)|*.pdf|All files (*.*)|*.*";
                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    txtBrowse.Text = fdlg.FileName;
                    strBrowse = txtBrowse.Text;
                }
            }
            else
            {
                MessageBox.Show("Please Select Cell.","Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void getTab2()
        {
            string strTab2 = "select distinct agreement_document.AUTO_ID,agreement_document.DOC_NO,agreement_document.PERSONAL_ID,PERSONAL_NAME,agreement_document.DEPARTMENT,agreement_document.BORROW_DATE,agreement_document.PURPOSE,agreement_document.COMPLETE_BY,agreement_document.COMPLETE_DATE,agreement_document.flg_status,agreement_document.REMARK from agreement_document,set_data_update where set_data_update.DOC_NO = agreement_document.DOC_NO and set_data_update.flg_status = '1'";
            MySqlDataAdapter adap = new MySqlDataAdapter(strTab2, conn);
            DataTable tb = new DataTable();
            adap.Fill(tb);
            dtGridViewAgreement.DataSource = tb;
        }

        private void tabControl1_Click(object sender, EventArgs e)
        {
            checktab();
        }

        public void checktab()
        {
            if (tabControl1.SelectedIndex == 0)
            {
                PleaseWait.Create();
                try
                {
                    
                    getMaster("NOTCOMPLETE");
                    getCount("TOTAL");
                    getCount("MONTH");
                    getCount("TODAY");
                    getCount("COMPLETE");
                    getCount("NOTCOMPLETE");

                    comboSearchBy.Text = "";
                    comboSearchType.Text = "";
                    txtSearch.Text = "";

                    txtDocumentNo.Text = "";
                    txtPersonalID.Text = "";
                    txtPersonalName.Text = "";
                    txtDepartment.Text = "";
                    txtBorrowDate.Text = "";
                    txtPurpose.Text = "";
                    txtStatus.Text = "";
                    txtComplete.Text = "";
                    txtCompleteDate.Text = "";
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
                    getMaster("COMPLETE");
                    comboStatusReprint.Text = "";
                    comboSearchReprint.Text = "";
                    comboSearchByReprint.Text = "";
                    txtSearchReprint.Text = "";

                    txtDocNoReprint.Text = "";
                    txtPersonalIDReprint.Text = "";
                    txtPersonalNameReprint.Text = "";
                    txtDepartmentReprint.Text = "";
                    txtBorrowDateReprint.Text = "";
                    txtPurposeReprint.Text = "";
                    txtStatusReprint.Text = "";
                    txtCompleteByReprint.Text = "";
                    txtCompleteDateReprint.Text = "";
                }
                finally
                {
                    PleaseWait.Destroy();
                }
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (comboSearchType.Text == "Pending")
            {
                if (comboSearchBy.Text != "")
                {
                    if (txtSearch.Text != "")
                    {
                        PleaseWait.Create();
                        try
                        {
                            searchby();
                        }
                        finally
                        {
                            PleaseWait.Destroy();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please Fill word in textSearch","Warning",MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    PleaseWait.Create();
                    try
                    {
                        getMaster("NOTCOMPLETE");
                    }
                    finally
                    {
                        PleaseWait.Destroy();
                    }
                }

            }
            else if (comboSearchType.Text == "Complete")
            {
                if (comboSearchBy.Text != "")
                {
                    if (txtSearch.Text != "")
                    {
                        PleaseWait.Create();
                        try
                        {
                            searchby();
                        }
                        finally
                        {
                            PleaseWait.Destroy();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please Fill word in textSearch", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }
                else
                {
                    PleaseWait.Create();
                    try
                    {
                        getMaster("COMPLETE");
                    }
                    finally
                    {
                        PleaseWait.Destroy();
                    }
                }
            }
            else
            {
                MessageBox.Show("No Data", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }


        private void btnSearchReprint_Click(object sender, EventArgs e)
        {
            if (comboStatusReprint.Text == "Pending")
            {
                if (comboSearchReprint.Text != "")
                {
                    if (comboSearchByReprint.Text != "")
                    {
                        if (txtSearchReprint.Text != "")
                        {
                            PleaseWait.Create();
                            try
                            {
                                getSearchMaster();
                            }
                            finally
                            {
                                PleaseWait.Destroy();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please Fill word in textSearch!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select information","Warning",MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    PleaseWait.Create();
                    try
                    {
                        getTab2();
                    }
                    finally
                    {
                        PleaseWait.Destroy();
                    }
                }

            }
            else if (comboStatusReprint.Text == "Complete")
            {
                if (comboSearchReprint.Text != "")
                {
                    if (comboSearchByReprint.Text != "")
                    {
                        if (txtSearchReprint.Text != "")
                        {
                            PleaseWait.Create();
                            try
                            {
                                getSearchMaster();
                            }
                            finally
                            {
                                PleaseWait.Destroy();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please Fill word in textSearch!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Please select information", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }
                else
                {
                    PleaseWait.Create();
                    try
                    {
                        getMaster("COMPLETE");
                    }
                    finally
                    {
                        PleaseWait.Destroy();
                    }
                }
            }
            else
            {
                MessageBox.Show("No Data", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void comboStatusReprint_SelectedIndexChanged(object sender, EventArgs e)
        {
            getStatusDocument(); 
        }

        private void comboSearchReprint_SelectedIndexChanged(object sender, EventArgs e)
        {
            getSearchReprint();
        }


        string[] getModel = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] A = new string[18] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] B = new string[500];
        int[] C = new int[501];
        string a = "";
        int n;
        string aaa;

        private void btnReprint_Click(object sender, EventArgs e)
        {
     
            int flg = 0;
            string EIAFirst;

            if (txtDocNoReprint.Text.Trim() == "" | txtPersonalIDReprint.Text.Trim() == "" | txtPersonalNameReprint.Text.Trim() == "")
            {
                MessageBox.Show("Please Select Row.!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                if (txtStatusReprint.Text == "Pending")
                {
                    Connection();
                    int i = 0;
                    //MySqlCommand cmd = new MySqlCommand("select distinct set_no from set_data_update where DOC_NO = '" + txtDocNoReprint.Text.Trim() + "'", conn);
                    MySqlCommand cmd = new MySqlCommand("select set_no from set_data_update where DOC_NO = '" + txtDocNoReprint.Text.Trim() + "' and flg_status = '1'", conn);
                    MySqlDataReader rad = cmd.ExecuteReader();
                    while (rad.Read())
                    {
                        Connection();
                        MySqlDataAdapter adap = new MySqlDataAdapter("SELECT SET_NO,MODEL FROM master_set ", conn);
                        DataTable results = new DataTable();
                        adap.Fill(results);

                        foreach (DataRow rows in results.Rows)
                        {
                            string flg3 = "";
                            string flg2 = "";
                            string flg1 = "";

                            string _SetNo = Convert.ToString(rows["SET_NO"]);
                            if (rad["set_no"].ToString() == _SetNo)
                            {
                                string _Model = Convert.ToString(rows["MODEL"]);
                                if (flg == 0)
                                {
                                    for (int k = 0; k < dt2.Rows.Count; k++)
                                    {
                                        if (rad["set_no"].ToString() != dt2.Rows[k][0].ToString()) //set ที่จะยืมต้องไม่เท่ากับ set ที่มีใน dt2
                                        {
                                        }
                                        else
                                        {
                                            MessageBox.Show("Document Number : " + txtDocNoReprint.Text + " was return set all of information");
                                            a = "Dup";
                                            return;
                                        }
                                    }
                                    if (a != "Dup")
                                    {
                                        eia = rad["set_no"].ToString(); //read string
                                        check_cut = eia.Substring(3, 6); //003310
                                        checkcutA = check_cut.Substring(2, 2); //33
                               
                                        for (int t = 0; t <= 17; t++)
                                        {
                                            if (A[t] != "")
                                            {
                                                l++;
                                            }
                                        }

                                        for (int j = 0; j <= 499; j++)
                                        {
                                            if (B[j] != null)
                                            {
                                                n++;
                                            }
                                        }

                                        if (l == 0)
                                        {
                                            A[0] = checkcutA;
                                        }

                                        if (n == 0)
                                        {
                                            B[0] = check_cut;
                                        }

                                        if (l != 0)
                                        {
                                            for (int t = 0; t < l; t++)
                                            {
                                                if (A[t] == checkcutA)
                                                {
                                                    flg1 = "noadd";
                                                }
                                            }
                                            if (flg1 != "noadd" && l <= 18) //เชค Array A เกิน limit ไหม
                                            {
                                                if (l < 18)
                                                {
                                                    A[l] = checkcutA;
                                                }
                                                else
                                                {
                                                    MessageBox.Show("model EIA failed ");
                                                    flg2 = "no show";
                                                }
                                            }
                                            l = 0;

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

                                        if (flg2 != "no show")
                                        {
                                            dt2.Rows.Add(rad["set_no"].ToString(),_Model);
                                        }

                                    }

                                }

                            }
                        }
                        i++;
                        
                    }
                    closeCon();
 
                    if (b == "SAMPLE SET CONTROL")
                    {
                        DateTime df = DateTime.Now;
                        string DateBorrow = df.ToString("yyyy-MM-dd");

                        if (txtPurposeReprint.Text == "Internal")
                        {
                            int min = 999999;
                            int max = 0;

                            Array.Sort(B);
                            Array.Reverse(B);
                            Array.Sort(A);
                            Array.Reverse(A);

                            int a = 1;
                            foreach (var r in B)
                            {
                                int recieve = Convert.ToInt32(r);
                                C[a] = recieve;
                                a++;
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
                                        aaa = Convert.ToString(bbb);

                                        if (aaa != "")
                                        {
                                            string N = "00" + aaa;
                                            string check_eia = "EIA" + N;
                                            foreach (DataRow dr in dt2.Rows)
                                            {
                                                eiaBor = Convert.ToString(dr["SET_NO"]);
                                                if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                {
                                                    if (aaa.Substring(0, 2) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                    {
                                                        modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                        int int_cut = Convert.ToInt32(aaa); //convert aaa string to int aaa = 3310
                                                        int int_cut2 = Convert.ToInt32(aaa);


                                                        if (min > int_cut) //3310 > 3315
                                                        {
                                                            min = int_cut; // min = 3310
                                                        }
                                                        if (max < int_cut)
                                                        {
                                                            max = int_cut;

                                                        }

                                                        //convert min,max to String
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
                                                    }

                                                }

                                            }
                                        }

                                    }
                                    else if ((C[r - 1] - C[r]) != 1 && r != 1)
                                    {
                                        if (r != 0 && one == aaa.Substring(0, 2))
                                        {
                                            newdata.Rows.Add(Unit, txtDocNoReprint.Text, txtPersonalIDReprint.Text, txtPersonalNameReprint.Text, txtDepartmentReprint.Text, modelBor, EIARange, CountBorrow2, "Internal", DateBorrow);
                                            min = 999999;
                                            max = 0;
                                            CountBorrow2 = 0;
                                            EIARange = "";
                                        }

                                        int bbb = C[r];
                                        aaa = Convert.ToString(bbb);
                                        if (aaa != "")
                                        {
                                            string N = "00" + aaa;
                                            string check_eia = "EIA" + N;
                                            foreach (DataRow dr in dt2.Rows)
                                            {

                                                eiaBor = Convert.ToString(dr["SET_NO"]);
                                                if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                {
                                                    if (aaa.Substring(0, 2) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                    {
                                                        modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                        int int_cut = Convert.ToInt32(aaa); //convert aaa string to int aaa = 3310
                                                        int int_cut2 = Convert.ToInt32(aaa);


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
                                        newdata.Rows.Add(Unit, txtDocNoReprint.Text, txtPersonalIDReprint.Text, txtPersonalNameReprint.Text, txtDepartmentReprint.Text, modelBor, EIARange, CountBorrow2, "Internal", DateBorrow);
                                        min = 999999;
                                        max = 0;
                                        CountBorrow2 = 0;
                                        one = "";
                                    }
                                }

                            }
                            Reprint expPDF = new Reprint();
                            expPDF.select_db = b;
                            expPDF.exportPDF(newdata);

                        }
                        else if (txtPurposeReprint.Text == "External")
                        {
                            int min = 999999;
                            int max = 0;

                            Array.Sort(B);
                            Array.Reverse(B);
                            Array.Sort(A);
                            Array.Reverse(A);

                            int a = 1;
                            foreach (var r in B)
                            {
                                int recieve = Convert.ToInt32(r);
                                C[a] = recieve;
                                a++;
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
                                        aaa = Convert.ToString(bbb);

                                        if (aaa != "")
                                        {
                                            string N = "00" + aaa;
                                            string check_eia = "EIA" + N;
                                            foreach (DataRow dr in dt2.Rows)
                                            {

                                                eiaBor = Convert.ToString(dr["SET_NO"]);
                                                if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                {
                                                    if (aaa.Substring(0, 2) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                    {
                                                        modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                        int int_cut = Convert.ToInt32(aaa); //convert aaa string to int aaa = 3310
                                                        int int_cut2 = Convert.ToInt32(aaa);


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
                                                    }

                                                }

                                            }
                                        }

                                    }
                                    else if ((C[r - 1] - C[r]) != 1 && r != 1)
                                    {

                                        if (r != 0 && one == aaa.Substring(0, 2))
                                        {
                                            newdata.Rows.Add(Unit, txtDocNoReprint.Text, txtPersonalIDReprint.Text, txtPersonalNameReprint.Text, txtDepartmentReprint.Text, modelBor, EIARange, CountBorrow2, "External", DateBorrow);
                                            min = 999999;
                                            max = 0;
                                            CountBorrow2 = 0;
                                            EIARange = "";

                                        }

                                        int bbb = C[r];
                                        aaa = Convert.ToString(bbb);
                                        if (aaa != "")
                                        {
                                            string N = "00" + aaa;
                                            string check_eia = "EIA" + N;
                                            foreach (DataRow dr in dt2.Rows)
                                            {

                                                eiaBor = Convert.ToString(dr["SET_NO"]);
                                                if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                {
                                                    if (aaa.Substring(0, 2) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                    {
                                                        modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                        int int_cut = Convert.ToInt32(aaa); //convert aaa string to int aaa = 3310
                                                        int int_cut2 = Convert.ToInt32(aaa);


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
                                        newdata.Rows.Add(Unit, txtDocNoReprint.Text, txtPersonalIDReprint.Text, txtPersonalNameReprint.Text, txtDepartmentReprint.Text, modelBor, EIARange, CountBorrow2, "External", DateBorrow);
                                        min = 999999;
                                        max = 0;
                                        CountBorrow2 = 0;
                                        one = "";
                                    }
                                }

                            }
                            Reprint expPDF = new Reprint();
                            expPDF.select_db = b;
                            expPDF.exportPDF(newdata);
                        }

                    }
                    else if (b == "ENGINEERING TRAINING CENTER")
                    {
                        DateTime df = DateTime.Now;
                        string DateBorrow = df.ToString("yyyy-MM-dd");

                        if (txtPurposeReprint.Text == "Internal")
                        {
                            int min = 999999;
                            int max = 0;

                            Array.Sort(B);
                            Array.Reverse(B);
                            Array.Sort(A);
                            Array.Reverse(A);

                            int a = 1;
                            foreach (var r in B)
                            {
                                int recieve = Convert.ToInt32(r);
                                C[a] = recieve;
                                a++;
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
                                        aaa = Convert.ToString(bbb);

                                        if (aaa != "")
                                        {
                                            string N = "00" + aaa;
                                            string check_eia = "EIA" + N;
                                            foreach (DataRow dr in dt2.Rows)
                                            {

                                                eiaBor = Convert.ToString(dr["SET_NO"]);
                                                if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                {
                                                    if (aaa.Substring(0, 2) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                    {
                                                        modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                        int int_cut = Convert.ToInt32(aaa); //convert aaa string to int aaa = 3310
                                                        int int_cut2 = Convert.ToInt32(aaa);


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
                                                    }

                                                }

                                            }
                                        }

                                    }
                                    else if ((C[r - 1] - C[r]) != 1 && r != 1)
                                    {

                                        if (r != 0 && one == aaa.Substring(0, 2))
                                        {
                                            newdata.Rows.Add(Unit, txtDocNoReprint.Text, txtPersonalIDReprint.Text, txtPersonalNameReprint.Text, txtDepartmentReprint.Text, modelBor, EIARange, CountBorrow2, "Internal", DateBorrow);
                                            min = 999999;
                                            max = 0;
                                            CountBorrow2 = 0;
                                            EIARange = "";

                                        }

                                        int bbb = C[r];
                                        aaa = Convert.ToString(bbb);
                                        if (aaa != "")
                                        {
                                            string N = "00" + aaa;
                                            string check_eia = "EIA" + N;
                                            foreach (DataRow dr in dt2.Rows)
                                            {

                                                eiaBor = Convert.ToString(dr["SET_NO"]);
                                                if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                {
                                                    if (aaa.Substring(0, 2) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                    {
                                                        modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                        int int_cut = Convert.ToInt32(aaa); //convert aaa string to int aaa = 3310
                                                        int int_cut2 = Convert.ToInt32(aaa);


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
                                        newdata.Rows.Add(Unit, txtDocNoReprint.Text, txtPersonalIDReprint.Text, txtPersonalNameReprint.Text, txtDepartmentReprint.Text, modelBor, EIARange, CountBorrow2, "Internal", DateBorrow);
                                        min = 999999;
                                        max = 0;
                                        CountBorrow2 = 0;
                                        one = "";
                                    }
                                }

                            }
                            Reprint expPDF = new Reprint();
                            expPDF.select_db = b;
                            expPDF.exportPDF(newdata);

                        }
                        else if (txtPurposeReprint.Text == "External")
                        {
                            int min = 999999;
                            int max = 0;

                            Array.Sort(B);
                            Array.Reverse(B);
                            Array.Sort(A);
                            Array.Reverse(A);

                            int a = 1;
                            foreach (var r in B)
                            {
                                int recieve = Convert.ToInt32(r);
                                C[a] = recieve;
                                a++;
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
                                        aaa = Convert.ToString(bbb);

                                        if (aaa != "")
                                        {
                                            string N = "00" + aaa;
                                            string check_eia = "EIA" + N;
                                            foreach (DataRow dr in dt2.Rows)
                                            {

                                                eiaBor = Convert.ToString(dr["SET_NO"]);
                                                if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                {
                                                    if (aaa.Substring(0, 2) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                    {
                                                        modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                        int int_cut = Convert.ToInt32(aaa); //convert aaa string to int aaa = 3310
                                                        int int_cut2 = Convert.ToInt32(aaa);


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
                                                    }

                                                }

                                            }
                                        }

                                    }
                                    else if ((C[r - 1] - C[r]) != 1 && r != 1)
                                    {

                                        if (r != 0 && one == aaa.Substring(0, 2))
                                        {
                                            newdata.Rows.Add(Unit, txtDocNoReprint.Text, txtPersonalIDReprint.Text, txtPersonalNameReprint.Text, txtDepartmentReprint.Text, modelBor, EIARange, CountBorrow2, "External", DateBorrow);
                                            min = 999999;
                                            max = 0;
                                            CountBorrow2 = 0;
                                            EIARange = "";

                                        }

                                        int bbb = C[r];
                                        aaa = Convert.ToString(bbb);
                                        if (aaa != "")
                                        {
                                            string N = "00" + aaa;
                                            string check_eia = "EIA" + N;
                                            foreach (DataRow dr in dt2.Rows)
                                            {

                                                eiaBor = Convert.ToString(dr["SET_NO"]);
                                                if (eiaBor == check_eia) //check SET_NO ที่ยืมเหมือนกันไหม
                                                {
                                                    if (aaa.Substring(0, 2) == one) //cut aaa = 33 เปรียบเทียบกับ one = 33
                                                    {
                                                        modelBor = Convert.ToString(dr["MODEL"]); //Model 
                                                        int int_cut = Convert.ToInt32(aaa); //convert aaa string to int aaa = 3310
                                                        int int_cut2 = Convert.ToInt32(aaa);


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
                                        newdata.Rows.Add(Unit, txtDocNoReprint.Text, txtPersonalIDReprint.Text, txtPersonalNameReprint.Text, txtDepartmentReprint.Text, modelBor, EIARange, CountBorrow2, "External", DateBorrow);
                                        min = 999999;
                                        max = 0;
                                        CountBorrow2 = 0;
                                        one = "";
                                    }
                                }

                            }
                            Reprint expPDF = new Reprint();
                            expPDF.select_db = b;
                            expPDF.exportPDF(newdata);
                        }
                    }
                }
                else if (txtStatusReprint.Text == "Complete")
                {
                    string path = "http://43.72.52.12/uploadfile/agreement_document/complete_document/";
                    string filename = txtDocNoReprint.Text.Trim();
                    string loadfile = path + filename + ".pdf";

                    try
                    {
                        WebClient client = new WebClient();
                        NetworkCredential nc = new NetworkCredential(strID, strPassword);

                        Uri addy = new Uri("http://43.72.52.12/uploadfile/agreement_document/complete_document/" + txtDocNoReprint.Text.Trim() + ".pdf");
                        client.Credentials = nc;
                        System.Diagnostics.Process.Start("IExplore.exe", loadfile);
  
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }              

                }
            }
        }

        private void comboSearchType_SelectedIndexChanged(object sender, EventArgs e)
        {
            searchtype();
        }

       
        private void pictureBoxRestore_DoubleClick(object sender, EventArgs e)
        {
            string value = Interaction.InputBox("Do you want to restore Document No. : " + txtDocNoReprint.Text + "?\n" + "Please enter your id.", "Confirm", "");

            if (value == "")
            {
                return;
            }

            if (value == empID)
            {
                PleaseWait.Create();
                try
                {
                    ReturnHistory(txtDocNoReprint.Text);
                    getMaster("COMPLETE");
                    
                }
                finally
                {
                    PleaseWait.Destroy();
                }

                MessageBox.Show("Restore Completed!");

            }
            else
            {
                MessageBox.Show("Access Denied:" + value);
            }
        }

        public void ReturnHistory(string strDOC)
        {
            Connection();
            string strSql = "";
            strSql = "update agreement_document set flg_status = 'Pending', COMPLETE_BY = '',COMPLETE_DATE = '',PATH_COMPLETE = '',REMARK = 'Restore_Upload' where DOC_NO = '"+strDOC+"' ";
            MySqlCommand cmd = new MySqlCommand(strSql,conn);
            cmd.ExecuteNonQuery();
            closeCon();
        }



    }
}
