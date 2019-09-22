using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using STTADUser;
using INI;
using System.Data.SqlClient;
using System.IO.Ports;


namespace MASTER_SET_CONTROL
{
    public partial class frmLogin : Form
    {
        public string b;
        public string P;
        IniFile iniconfig;
        string usernameid;
        SqlConnection SQLConnection;
        SqlDataReader SQLDataReader;
        clsLDAPAuthentication clsAuthen;


        public frmLogin()
        {
            InitializeComponent();
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = "LOGIN MASTER SET CONTROL" + String.Format(" --- Version {0}", version);
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {

            CheckLogin();
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            txuser.Focus();
            txuser.Select();

            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
            clsAuthen = new clsLDAPAuthentication();

        }

        private string Authenticate(string user, string password)
        {

            string Res = clsAuthen.checkAuthorizedInDomain("APTHCHOADS201.AP.SONY.COM", user, password);
            return Res.ToString();

        }
        public void CheckLogin()
        {
 
            string IP = iniconfig.IniReadValue("Server", "IP");
            string DB = iniconfig.IniReadValue("Server", "DB");

            string ConnectionString = "Server='" + IP + "';";
            ConnectionString += "User ID=sa;";
            ConnectionString += "Password=s;";
            ConnectionString += "Database='" + DB + "';";

            SQLConnection = new SqlConnection(ConnectionString);


            SQLConnection.Open();
            string username, Gid;

            string resultAccess = "";
            resultAccess = Authenticate(txuser.Text, txpass.Text);
            if (resultAccess == "PASS")
            {
                usernameid = txuser.Text;
                string SQLStatement = "";

                SQLStatement += "SELECT * FROM [TBL_MANPOW_MASTER] where GID = '" + txuser.Text.Trim() + "' AND EMP_STATUS_EN != 'Termination'";
                
                SqlCommand cmd = new SqlCommand(SQLStatement, SQLConnection);

                SQLDataReader = cmd.ExecuteReader();
                if (SQLDataReader.Read())
                {
                    Gid = SQLDataReader["GID"].ToString();
                    username = SQLDataReader["EMP_NAME_EN"].ToString();
                    this.Hide();
                    b = txuser.Text.Trim();
                    P = txpass.Text.Trim();
                    checkCredential.username = b;
                    checkCredential.password = P;

                    select_db_form select = new select_db_form();
                    select.checkuser(b.ToString(),P.ToString());
                    select.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Login Fail!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txuser.Text = "";
                    txpass.Text = "";
                }

            }
            else
            {
                MessageBox.Show("Can not access domain AP!" + resultAccess, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txpass.Text = "";
            }
        }

        private void txuser_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CheckLogin();
            }
        }

        private void txpass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CheckLogin();
            }
        }

        private void txuser_TextChanged(object sender, EventArgs e)
        {
            if (txuser.Text != "")
            {

                if (txuser.Text.Length == 10)
                {
                    txpass.Focus();

                }

            }
        }

        private void txuser_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < 48 || e.KeyChar > 57) && (e.KeyChar != 8))
            {
                e.Handled = true;
            }
        }

    }

}
