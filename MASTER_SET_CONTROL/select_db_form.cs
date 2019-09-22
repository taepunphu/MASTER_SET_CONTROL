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
    public partial class select_db_form : Form
    {
        string a;
        public string select_DB;
        public string member_DB;
        public string strCon;
        IniFile iniconfig;
        MySqlConnection conn;

        //public string strCon = "Server=43.72.52.12;Database=eia_master_set_control;Uid=root;Password=123456*;Convert Zero Datetime=True;";
        //public string strCon = "host=localhost;Database=eia_master_set_control;Uid=root;Password=123456;Convert Zero Datetime=True;";

        public select_db_form()
        {
            InitializeComponent();
        }

        private void select_db_form_Load(object sender, EventArgs e)
        {

            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
            string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
            string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

            strCon = "Server=" + IP + ";";
            strCon += "Uid=root;";
            strCon += "Password=123456*;";
            strCon += "Database=" + DB + ";";

            conn = new MySqlConnection(strCon);

            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = "MASTER SET CONTROL" + String.Format(" --- Version {0}", version);

            Connection();
            MySqlDataAdapter adapter = new MySqlDataAdapter("select * from personal_login", conn);
            DataTable table = new DataTable();
            adapter.Fill(table);
            foreach (DataRow dr in table.Rows)
            {
                string ID = Convert.ToString(dr["userID"]);
                string groupcheck = Convert.ToString(dr["Group"]);
                if (user == ID)
                {
                    if (groupcheck == "SSC")
                    {
                        comboBox1.Items.Add("SAMPLE SET CONTROL");
                    }
                    else if (groupcheck == "ETC")
                    {
                        comboBox1.Items.Add("ENGINEERING TRAINING CENTER");
                    }
                    else if (groupcheck == "SSC" && groupcheck == "ETC")
                    {
                        comboBox1.Items.Add("SAMPLE SET CONTROL");
                        comboBox1.Items.Add("ENGINEERING TRAINING CENTER");
                    }

                }

            }

        }

        public void checkSelect()
        {

            if ((string)comboBox1.SelectedItem == "SAMPLE SET CONTROL")
            {
                this.Hide();
                a = comboBox1.Text;
                select_DB = comboBox1.Text;
                Form1 fm = new Form1();
                fm.ab(a.ToString());
                fm.Show();
            }
            else if ((string)comboBox1.SelectedItem == "ENGINEERING TRAINING CENTER")
            {
                this.Hide();
                a = comboBox1.Text;
                select_DB = comboBox1.Text;
                Form1 fm = new Form1();
                fm.ab(a.ToString());
                fm.Show();
            }
            else
            {
                MessageBox.Show("Please Select Database", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_selectdb_Click(object sender, EventArgs e)
        {
            checkSelect();
        }

        private void comboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                checkSelect();
            }
        }

        string user;
        public void checkuser(string b, string P)
        {
            user = b.ToString();
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

    }
}
