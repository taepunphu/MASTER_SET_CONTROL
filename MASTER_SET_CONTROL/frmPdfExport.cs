using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Net;
using System.Security.Permissions;
using System.Security.Principal;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using System.Runtime.ConstrainedExecution;
using System.Security;

using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using MySql.Data.MySqlClient;

using INI;
using STTADUser;

namespace MASTER_SET_CONTROL
{
    public partial class frmPdfExport : Form
    {
        MySqlConnection condb;
        public static string scan_D;
        public string select_db;
        public string strCon;
        public string command;
        public string com;
        public string com1;
        public string com2;
        public string com3;
        public string com4;
        public string com5;
        public string com6;
        string number;
        //string dateDoc2;
        string DocumentNumber;
        string ID;
        string namesurname;
        string Department;
        string Model;
        string EIA;
        string Count;
        string cb;
        string Unit;
        string strID;
        string strPassword;
        string DateBorrow;
        IniFile iniconfig;
        DataTable table = new DataTable();
        
        private void frmPdfExport_Load(object sender, EventArgs e)
        {

        }

        public frmPdfExport()
        {
           InitializeComponent();
        }

        
        

        public frmPdfExport(string a, string b, string c, string d, string e, string f, string g, string h, string i, string j,string k)
        {
            strID = checkCredential.username;
            strPassword = checkCredential.password;

            if (j == "SAMPLE SET CONTROL")
            {
                IniFile Gen;
                Gen = new IniFile(Application.StartupPath + "\\generate.ini");
                DateTime dt = DateTime.Now;
                DateTime dt2 = DateTime.Now;
                string datedoc = dt.ToString("yyyy-MM-dd");
                string datedoc2 = dt2.ToString("yyyyMM");

                string recieveDocumentNo = a;
                if (h == "1")
                {
                    if (recieveDocumentNo == "SSC")
                    {
                        datecheck();
                        int v = Convert.ToInt32(Gen.IniReadValue("generate", "gen_ssc"));
                        v = v + 1;
                        number = v.ToString();
                        Gen.IniWriteValue("generate", "gen_ssc", number);

                        if (i == "Internal")
                        {
                            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                            string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                            string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                            string ConnectionString = "Server=" + IP + ";";
                            ConnectionString += "Uid=root;";
                            ConnectionString += "Password=123456*;";
                            ConnectionString += "Database=" + DB + ";";
                            condb = new MySqlConnection(ConnectionString);
                            condb.Open();
                
                            string DocumentNumber1 = recieveDocumentNo + datedoc2 + "-" + number;
                            string Model1 = f;
                            string EIA1 = g;
                            string ID1 = c;
                            string count1 = h;
                            string DueDate1 = b;
                            string Department1 = e;
                            string name1 = d;
                            string EIA = k;

                            SaveFileDialog save = new SaveFileDialog();
                            save.Filter = "PDF Files (*.pdf)|*.pdf|All files (*.*)|*.*";
                            save.FileName = recieveDocumentNo + datedoc2 + "-" + number;
                            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {

                                string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                                string path = save.FileName;

                                PdfReader reader = new PdfReader(fromfileInternal);
                                PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                                AcroFields fields = stamper.AcroFields;

                                fields.SetField("writeDocumentNo", DocumentNumber1);
                                fields.SetField("writeBorrowDate", datedoc);
                                fields.SetField("writeENNo", ID1);
                                fields.SetField("writeNameSurname", name1);
                                fields.SetField("writeDepartment", Department1);
                                fields.SetField("writeBorrowDate", datedoc);

                                fields.SetField("NoRow1", "1");
                                fields.SetField("TypeRow1", Model1);
                                fields.SetField("EIANoRow1", EIA1);
                                fields.SetField("UMORow1", "Set");
                                fields.SetField("QTYRow1", count1);

                                stamper.FormFlattening = true;
                                stamper.Close();

                                command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,purpose,flg_status,rec_date) values('" + DocumentNumber1 + "','" + ID1 + "', '" + name1 + "', '" + Department1 + "','" + datedoc + "', '" + i + "', 'Pending',sysdate());";
                                MySqlCommand cmd = new MySqlCommand(command, condb);
                                cmd.ExecuteNonQuery();

                                com = "update set_data_update set DOC_NO = '" + DocumentNumber1 + "' where set_no = '" + EIA + "' and flg_status = '1'";
                                MySqlCommand cm = new MySqlCommand(com, condb);
                                cm.ExecuteNonQuery();
                              
                                MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                System.Diagnostics.Process.Start(path);
                            }
                            condb.Close();
                        }
                        else if (i == "External")
                        {
                            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                            string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                            string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                            string ConnectionString = "Server=" + IP + ";";
                            ConnectionString += "Uid=root;";
                            ConnectionString += "Password=123456*;";
                            ConnectionString += "Database=" + DB + ";";
                            condb = new MySqlConnection(ConnectionString);
                            condb.Open();

                            string DocumentNumber1 = recieveDocumentNo + datedoc2 + "-" + number;
                            string Model1 = f;
                            string EIA1 = g;
                            string ID1 = c;
                            string count1 = h;
                            string DueDate1 = b;
                            string Department1 = e;
                            string name1 = d;
                            string EIA = k;

                            SaveFileDialog save = new SaveFileDialog();
                            save.Filter = "PDF Files (*.pdf)|*.pdf";
                            save.FileName = recieveDocumentNo + datedoc2 + "-" + number;
                            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                                string path = save.FileName;

                                PdfReader reader = new PdfReader(fromfileExternal);
                                PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                                AcroFields fields = stamper.AcroFields;

                                fields.SetField("writeDocumentNo", DocumentNumber1);
                                fields.SetField("writeBorrowDate", datedoc);
                                fields.SetField("writeENNo", ID1);
                                fields.SetField("writeNameSurname", name1);
                                fields.SetField("writeDepartment", Department1);
                                fields.SetField("writeBorrowDate", datedoc);

                                fields.SetField("NoRow1", "1");
                                fields.SetField("TypeRow1", Model1);
                                fields.SetField("EIANoRow1", EIA1);
                                fields.SetField("UMORow1", "Set");
                                fields.SetField("QTYRow1", count1);

                                stamper.FormFlattening = true;
                                stamper.Close();

                                command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,purpose,flg_status,rec_date) values('" + DocumentNumber1 + "','" + ID1 + "', '" + name1 + "', '" + Department1 + "','" + datedoc + "', '" + i + "', 'Pending',sysdate());";
                                MySqlCommand cmd = new MySqlCommand(command, condb);
                                cmd.ExecuteNonQuery();

                                com = "update set_data_update set DOC_NO = '" + DocumentNumber1 + "' where set_no = '" + EIA + "' and flg_status = '1' ";
                                MySqlCommand cm = new MySqlCommand(com, condb);
                                cm.ExecuteNonQuery();

                                MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                System.Diagnostics.Process.Start(path);
                            }
                            condb.Close();

                        }

                    }
                    else if (recieveDocumentNo == "ETC")
                    {
                        datecheck();
                        int v = Convert.ToInt32(Gen.IniReadValue("generate", "gen_etc"));
                        v = v + 1;
                        number = v.ToString();
                        Gen.IniWriteValue("generate", "gen_etc", number);

                        if (i == "Internal")
                        {
                            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                            string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                            string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                            string ConnectionString = "Server=" + IP + ";";
                            ConnectionString += "Uid=root;";
                            ConnectionString += "Password=123456*;";
                            ConnectionString += "Database=" + DB + ";";
                            condb = new MySqlConnection(ConnectionString);
                            condb.Open();

                            string DocumentNumber1 = recieveDocumentNo + datedoc2 + "-" + number;
                            string Model1 = f;
                            string EIA1 = g;
                            string ID1 = c;
                            string count1 = h;
                            string DueDate1 = b;
                            string Department1 = e;
                            string name1 = d;
                            string EIA = k;

                            SaveFileDialog save = new SaveFileDialog();
                            save.Filter = "PDF Files (*.pdf)|*.pdf";
                            save.FileName = recieveDocumentNo + datedoc2 + "-" + number;
                            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                                string path = save.FileName;

                                PdfReader reader = new PdfReader(fromfileInternal);
                                PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                                AcroFields fields = stamper.AcroFields;

                                fields.SetField("writeDocumentNo", DocumentNumber1);
                                fields.SetField("writeBorrowDate", datedoc);
                                fields.SetField("writeENNo", ID1);
                                fields.SetField("writeNameSurname", name1);
                                fields.SetField("writeDepartment", Department1);
                                fields.SetField("writeBorrowDate", datedoc);

                                fields.SetField("NoRow1", "1");
                                fields.SetField("TypeRow1", Model1);
                                fields.SetField("EIANoRow1", EIA1);
                                fields.SetField("UMORow1", "Set");
                                fields.SetField("QTYRow1", count1);

                                stamper.FormFlattening = true;
                                stamper.Close();

                                command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,purpose,flg_status,rec_date) values('" + DocumentNumber1 + "','" + ID1 + "', '" + name1 + "', '" + Department1 + "','" + datedoc + "', '" + i + "', 'Pending',sysdate());";
                                MySqlCommand cmd = new MySqlCommand(command, condb);
                                cmd.ExecuteNonQuery();

                                com = "update set_data_update set DOC_NO = '" + DocumentNumber1 + "' where set_no = '" + EIA + "' and flg_status = '1' ";
                                MySqlCommand cm = new MySqlCommand(com, condb);
                                cm.ExecuteNonQuery();

                                MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                System.Diagnostics.Process.Start(path);

                            }
                            condb.Close();

                        }
                        else if (i == "External")
                        {
                            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                            string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                            string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                            string ConnectionString = "Server=" + IP + ";";
                            ConnectionString += "Uid=root;";
                            ConnectionString += "Password=123456*;";
                            ConnectionString += "Database=" + DB + ";";
                            condb = new MySqlConnection(ConnectionString);
                            condb.Open();


                            string DocumentNumber1 = recieveDocumentNo + datedoc2 + "-" + number;
                            string Model1 = f;
                            string EIA1 = g;
                            string ID1 = c;
                            string count1 = h;
                            string DueDate1 = b;
                            string Department1 = e;
                            string name1 = d;

                            SaveFileDialog save = new SaveFileDialog();
                            save.Filter = "PDF Files (*.pdf)|*.pdf";
                            save.FileName = recieveDocumentNo + datedoc2 + "-" + number;
                            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                                string path = save.FileName;

                                PdfReader reader = new PdfReader(fromfileExternal);
                                PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                                AcroFields fields = stamper.AcroFields;

                                fields.SetField("writeDocumentNo", DocumentNumber1);
                                fields.SetField("writeBorrowDate", datedoc);
                                fields.SetField("writeENNo", ID1);
                                fields.SetField("writeNameSurname", name1);
                                fields.SetField("writeDepartment", Department1);
                                fields.SetField("writeBorrowDate", datedoc);

                                fields.SetField("NoRow1", "1");
                                fields.SetField("TypeRow1", Model1);
                                fields.SetField("EIANoRow1", EIA1);
                                fields.SetField("UMORow1", "Set");
                                fields.SetField("QTYRow1", count1);

                                stamper.FormFlattening = true;
                                stamper.Close();


                                command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,purpose,flg_status,rec_date) values('" + DocumentNumber1 + "','" + ID1 + "', '" + name1 + "', '" + Department1 + "','" + datedoc + "', '" + i + "', 'Pending',sysdate());";
                                MySqlCommand cmd = new MySqlCommand(command, condb);
                                cmd.ExecuteNonQuery();

                                com = "update set_data_update set DOC_NO = '" + DocumentNumber1 + "' where set_no ='" + EIA + "' and flg_status = '1'";
                                MySqlCommand cm = new MySqlCommand(com, condb);
                                cm.ExecuteNonQuery();

                                MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                System.Diagnostics.Process.Start(path);

                            }
                            condb.Close();
                        }

                    }

                }
                condb.Close();
          
            }
            else if (j == "ENGINEERING TRAINING CENTER")
            {

                IniFile Gen;
                Gen = new IniFile(Application.StartupPath + "\\generate.ini");
                DateTime dt = DateTime.Now;
                DateTime dt2 = DateTime.Now;
                string datedoc = dt.ToString("yyyy-MM-dd");
                string datedoc2 = dt2.ToString("yyyyMM");

                string recieveDocumentNo = a;
                if (h == "1")
                {
                    if (recieveDocumentNo == "SSC")
                    {
                        datecheck();
                        int v = Convert.ToInt32(Gen.IniReadValue("generate", "gen_ssc"));
                        v = v + 1;
                        number = v.ToString();
                        Gen.IniWriteValue("generate", "gen_ssc", number);

                        if (i == "Internal")
                        {
                            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                            string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                            string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                            string ConnectionString = "Server=" + IP + ";";
                            ConnectionString += "Uid=root;";
                            ConnectionString += "Password=123456*;";
                            ConnectionString += "Database=" + DB + ";";
                            condb = new MySqlConnection(ConnectionString);
                            condb.Open();

                            string DocumentNumber1 = recieveDocumentNo + datedoc2 + "-" + number;
                            string Model1 = f;
                            string EIA1 = g;
                            string ID1 = c;
                            string count1 = h;
                            string DueDate1 = b;
                            string Department1 = e;
                            string name1 = d;

                            SaveFileDialog save = new SaveFileDialog();
                            save.Filter = "PDF Files (*.pdf)|*.pdf";
                            save.FileName = recieveDocumentNo + datedoc2 + "-" + number;
                            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {

                                string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                                string path = save.FileName;

                                PdfReader reader = new PdfReader(fromfileInternal);
                                PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                                AcroFields fields = stamper.AcroFields;

                                fields.SetField("writeDocumentNo", DocumentNumber1);
                                fields.SetField("writeBorrowDate", datedoc);
                                fields.SetField("writeENNo", ID1);
                                fields.SetField("writeNameSurname", name1);
                                fields.SetField("writeDepartment", Department1);
                                fields.SetField("writeBorrowDate", datedoc);

                                fields.SetField("NoRow1", "1");
                                fields.SetField("TypeRow1", Model1);
                                fields.SetField("EIANoRow1", EIA1);
                                fields.SetField("UMORow1", "Set");
                                fields.SetField("QTYRow1", count1);

                                stamper.FormFlattening = true;
                                stamper.Close();

                                command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,purpose,flg_status,rec_date) values('" + DocumentNumber1 + "','" + ID1 + "', '" + name1 + "', '" + Department1 + "','" + datedoc + "', '" + i + "', 'Pending',sysdate());";
                                MySqlCommand cmd = new MySqlCommand(command, condb);
                                cmd.ExecuteNonQuery();

                                com = "update set_data_update set DOC_NO = '" + DocumentNumber1 + "' where set_no ='" + EIA + "' and flg_status = '1' ";
                                MySqlCommand cm = new MySqlCommand(com, condb);
                                cm.ExecuteNonQuery();

                                MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                System.Diagnostics.Process.Start(path);
                            }
                            condb.Close();

                        }
                        else if (i == "External")
                        {
                            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                            string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                            string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                            string ConnectionString = "Server=" + IP + ";";
                            ConnectionString += "Uid=root;";
                            ConnectionString += "Password=123456*;";
                            ConnectionString += "Database=" + DB + ";";
                            condb = new MySqlConnection(ConnectionString);
                            condb.Open();


                            string DocumentNumber1 = recieveDocumentNo + datedoc2 + "-" + number;
                            string Model1 = f;
                            string EIA1 = g;
                            string ID1 = c;
                            string count1 = h;
                            string DueDate1 = b;
                            string Department1 = e;
                            string name1 = d;

                            SaveFileDialog save = new SaveFileDialog();
                            save.Filter = "PDF Files (*.pdf)|*.pdf";
                            save.FileName = recieveDocumentNo + datedoc2 + "-" + number;
                            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                                string path = save.FileName;

                                PdfReader reader = new PdfReader(fromfileExternal);
                                PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                                AcroFields fields = stamper.AcroFields;

                                fields.SetField("writeDocumentNo", DocumentNumber1);
                                fields.SetField("writeBorrowDate", datedoc);
                                fields.SetField("writeENNo", ID1);
                                fields.SetField("writeNameSurname", name1);
                                fields.SetField("writeDepartment", Department1);
                                fields.SetField("writeBorrowDate", datedoc);

                                fields.SetField("NoRow1", "1");
                                fields.SetField("TypeRow1", Model1);
                                fields.SetField("EIANoRow1", EIA1);
                                fields.SetField("UMORow1", "Set");
                                fields.SetField("QTYRow1", count1);

                                stamper.FormFlattening = true;
                                stamper.Close();

                                command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,purpose,flg_status,rec_date) values('" + DocumentNumber1 + "','" + ID1 + "', '" + name1 + "', '" + Department1 + "','" + datedoc + "', '" + i + "', 'Pending',sysdate());";
                                MySqlCommand cmd = new MySqlCommand(command, condb);
                                cmd.ExecuteNonQuery();

                                com = "update set_data_update set DOC_NO = '" + DocumentNumber1 + "' where set_no ='" + EIA + "' and flg_status = '1'";
                                MySqlCommand cm = new MySqlCommand(com, condb);
                                cm.ExecuteNonQuery();

                                MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                System.Diagnostics.Process.Start(path);
                            }
                            condb.Close();
                        }

                    }
                    else if (recieveDocumentNo == "ETC")
                    {
                        datecheck();
                        int v = Convert.ToInt32(Gen.IniReadValue("generate", "gen_etc"));
                        v = v + 1;
                        number = v.ToString();
                        Gen.IniWriteValue("generate", "gen_etc", number);

                        if (i == "Internal")
                        {
                            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                            string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                            string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                            string ConnectionString = "Server=" + IP + ";";
                            ConnectionString += "Uid=root;";
                            ConnectionString += "Password=123456*;";
                            ConnectionString += "Database=" + DB + ";";
                            condb = new MySqlConnection(ConnectionString);
                            condb.Open();


                            string DocumentNumber1 = recieveDocumentNo + datedoc2 + "-" + number;
                            string Model1 = f;
                            string EIA1 = g;
                            string ID1 = c;
                            string count1 = h;
                            string DueDate1 = b;
                            string Department1 = e;
                            string name1 = d;

                            SaveFileDialog save = new SaveFileDialog();
                            save.Filter = "PDF Files (*.pdf)|*.pdf";
                            save.FileName = recieveDocumentNo + datedoc2 + "-" + number;
                            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                                string path = save.FileName;

                                PdfReader reader = new PdfReader(fromfileInternal);
                                PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                                AcroFields fields = stamper.AcroFields;

                                fields.SetField("writeDocumentNo", DocumentNumber1);
                                fields.SetField("writeBorrowDate", datedoc);
                                fields.SetField("writeENNo", ID1);
                                fields.SetField("writeNameSurname", name1);
                                fields.SetField("writeDepartment", Department1);
                                fields.SetField("writeBorrowDate", datedoc);

                                fields.SetField("NoRow1", "1");
                                fields.SetField("TypeRow1", Model1);
                                fields.SetField("EIANoRow1", EIA1);
                                fields.SetField("UMORow1", "Set");
                                fields.SetField("QTYRow1", count1);

                                stamper.FormFlattening = true;
                                stamper.Close();

                                command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,purpose,flg_status) values('" + DocumentNumber1 + "','" + ID1 + "', '" + name1 + "', '" + Department1 + "',sysdate(), '" + i + "', 'Pending');";
                                MySqlCommand cmd = new MySqlCommand(command, condb);
                                cmd.ExecuteNonQuery();

                                MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                System.Diagnostics.Process.Start(path);

                            }
                            condb.Close();

                        }
                        else if (i == "External")
                        {
                            iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                            string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                            string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                            string ConnectionString = "Server=" + IP + ";";
                            ConnectionString += "Uid=root;";
                            ConnectionString += "Password=123456*;";
                            ConnectionString += "Database=" + DB + ";";
                            condb = new MySqlConnection(ConnectionString);
                            condb.Open();


                            string DocumentNumber1 = recieveDocumentNo + datedoc2 + "-" + number;
                            string Model1 = f;
                            string EIA1 = g;
                            string ID1 = c;
                            string count1 = h;
                            string DueDate1 = b;
                            string Department1 = e;
                            string name1 = d;

                            SaveFileDialog save = new SaveFileDialog();
                            save.Filter = "PDF Files (*.pdf)|*.pdf";
                            save.FileName = recieveDocumentNo + datedoc2 + "-" + number;
                            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                                string path = save.FileName;

                                PdfReader reader = new PdfReader(fromfileExternal);
                                PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                                AcroFields fields = stamper.AcroFields;

                                fields.SetField("writeDocumentNo", DocumentNumber1);
                                fields.SetField("writeBorrowDate", datedoc);
                                fields.SetField("writeENNo", ID1);
                                fields.SetField("writeNameSurname", name1);
                                fields.SetField("writeDepartment", Department1);
                                fields.SetField("writeBorrowDate", datedoc);

                                fields.SetField("NoRow1", "1");
                                fields.SetField("TypeRow1", Model1);
                                fields.SetField("EIANoRow1", EIA1);
                                fields.SetField("UMORow1", "Set");
                                fields.SetField("QTYRow1", count1);

                                stamper.FormFlattening = true;
                                stamper.Close();

                                command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,purpose,flg_status,rec_date) values('" + DocumentNumber1 + "','" + ID1 + "', '" + name1 + "', '" + Department1 + "','" + datedoc + "', '" + i + "', 'Pending',sysdate());";
                                MySqlCommand cmd = new MySqlCommand(command, condb);
                                cmd.ExecuteNonQuery();

                                com = "update set_data_update set DOC_NO = '" + DocumentNumber1 + "' where set_no ='" + EIA + "' and flg_status = '1'";
                                MySqlCommand cm = new MySqlCommand(com, condb);
                                cm.ExecuteNonQuery();

                                MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                System.Diagnostics.Process.Start(path);

                            }
                            condb.Close();

                        }

                    }

                } 
          
            }
            else
            {
                MessageBox.Show("Cannot Connect to server!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        public void CountRow1(DataTable dt)
        {
           foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

              
            if (cb == "Internal")
            {
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                string ConnectionString = "Server=" + IP + ";";
                ConnectionString += "Uid=root;";
                ConnectionString += "Password=123456*;";
                ConnectionString += "Database=" + DB + ";";
                condb = new MySqlConnection(ConnectionString);
                condb.Open();

                string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF Files (*.pdf)|*.pdf";
                save.FileName = DocumentNumber;
                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string path = save.FileName;

                    PdfReader reader = new PdfReader(fromfileInternal);
                    PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                    AcroFields fields = stamper.AcroFields;

                    fields.SetField("writeDocumentNo", DocumentNumber);
                    fields.SetField("writeBorrowDate", DateBorrow);
                    fields.SetField("writeENNo", ID);
                    fields.SetField("writeNameSurname", namesurname);
                    fields.SetField("writeDepartment", Department);
                    fields.SetField("writeBorrowDate", DateBorrow);
                    fields.SetField("writeComment", txtComment.Text.Trim());

                    fields.SetField("NoRow1", "1");
                    fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                    fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                    fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                    fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());

                    stamper.FormFlattening = true;
                    stamper.Close();

                    command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                    MySqlCommand cmd = new MySqlCommand(command, condb);
                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    System.Diagnostics.Process.Start(path);
                    return;

                }  
         
            }
            else if (cb == "External")
            {
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                string ConnectionString = "Server=" + IP + ";";
                ConnectionString += "Uid=root;";
                ConnectionString += "Password=123456*;";
                ConnectionString += "Database=" + DB + ";";
                condb = new MySqlConnection(ConnectionString);
                condb.Open();

                string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF Files (*.pdf)|*.pdf";
                save.FileName = DocumentNumber;
                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string path = save.FileName;

                    PdfReader reader = new PdfReader(fromfileExternal);
                    PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                    AcroFields fields = stamper.AcroFields;

                    fields.SetField("writeDocumentNo", DocumentNumber);
                    fields.SetField("writeBorrowDate", DateBorrow);
                    fields.SetField("writeENNo", ID);
                    fields.SetField("writeNameSurname", namesurname);
                    fields.SetField("writeDepartment", Department);
                    fields.SetField("writeBorrowDate", DateBorrow);
                    fields.SetField("writeComment", txtComment.Text.Trim());

                    fields.SetField("NoRow1", "1");
                    fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                    fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                    fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                    fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());

                    stamper.FormFlattening = true;
                    stamper.Close();

                    command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                    MySqlCommand cmd = new MySqlCommand(command, condb);
                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    System.Diagnostics.Process.Start(path);
                    return;
                }
            }
            
          }

        }

        public void CountRow2(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);


                if (cb == "Internal")
            {
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                string ConnectionString = "Server=" + IP + ";";
                ConnectionString += "Uid=root;";
                ConnectionString += "Password=123456*;";
                ConnectionString += "Database=" + DB + ";";
                condb = new MySqlConnection(ConnectionString);
                condb.Open();

                string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF Files (*.pdf)|*.pdf";
                save.FileName = DocumentNumber;
                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                 {
                    string path = save.FileName;

                    PdfReader reader = new PdfReader(fromfileInternal);
                    PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                    AcroFields fields = stamper.AcroFields;

                    fields.SetField("writeDocumentNo", DocumentNumber);
                    fields.SetField("writeBorrdowDate", DateBorrow);
                    fields.SetField("writeENNo", ID);
                    fields.SetField("writeNameSurname", namesurname);
                    fields.SetField("writeDepartment", Department);
                    fields.SetField("writeBorrowDate", DateBorrow);
                    fields.SetField("writeComment", txtComment.Text.Trim());

                    fields.SetField("NoRow1", "1");
                    fields.SetField("NoRow2", "2");
                    fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                    fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                    fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                    fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                    fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                    fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                    fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                    fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());

                    stamper.FormFlattening = true;
                    stamper.Close();

                    command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                    MySqlCommand cmd = new MySqlCommand(command, condb);
                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    System.Diagnostics.Process.Start(path);
                    return;
                }

            }
            else if (cb == "External")
            {
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                string ConnectionString = "Server=" + IP + ";";
                ConnectionString += "Uid=root;";
                ConnectionString += "Password=123456*;";
                ConnectionString += "Database=" + DB + ";";
                condb = new MySqlConnection(ConnectionString);
                condb.Open();

                string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF Files (*.pdf)|*.pdf";
                save.FileName = DocumentNumber;
                   if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                    string path = save.FileName;

                    PdfReader reader = new PdfReader(fromfileExternal);
                    PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                    AcroFields fields = stamper.AcroFields;

                    fields.SetField("writeDocumentNo", DocumentNumber);
                    fields.SetField("writeBorrdowDate", DateBorrow);
                    fields.SetField("writeENNo", ID);
                    fields.SetField("writeNameSurname", namesurname);
                    fields.SetField("writeDepartment", Department);
                    fields.SetField("writeBorrowDate", DateBorrow);
                    fields.SetField("writeComment", txtComment.Text.Trim());

                    fields.SetField("NoRow1", "1");
                    fields.SetField("NoRow2", "2");
                    fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                    fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                    fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                    fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                    fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                    fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                    fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                    fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());

                    stamper.FormFlattening = true;
                    stamper.Close();

                    command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                    MySqlCommand cmd = new MySqlCommand(command, condb);
                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    System.Diagnostics.Process.Start(path);
                    return;
                   }
                }

            }
            
        }

        public void CountRow3(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                 if (cb == "Internal")
                  {
                      iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                      string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                      string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                      string ConnectionString = "Server=" + IP + ";";
                      ConnectionString += "Uid=root;";
                      ConnectionString += "Password=123456*;";
                      ConnectionString += "Database=" + DB + ";";
                      condb = new MySqlConnection(ConnectionString);
                      condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        dt.Clear();
                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);


                        return;
            
                      }


            }
            else if (cb == "External")
            {
                iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                string ConnectionString = "Server=" + IP + ";";
                ConnectionString += "Uid=root;";
                ConnectionString += "Password=123456*;";
                ConnectionString += "Database=" + DB + ";";
                condb = new MySqlConnection(ConnectionString);
                condb.Open();

                string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF Files (*.pdf)|*.pdf";
                save.FileName = DocumentNumber;
                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string path = save.FileName;

                    PdfReader reader = new PdfReader(fromfileExternal);
                    PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                    AcroFields fields = stamper.AcroFields;

                    fields.SetField("writeDocumentNo", DocumentNumber);
                    fields.SetField("writeBorrowDate", DateBorrow);
                    fields.SetField("writeENNo", ID);
                    fields.SetField("writeNameSurname", namesurname);
                    fields.SetField("writeDepartment", Department);
                    fields.SetField("writeBorrowDate", DateBorrow);
                    fields.SetField("writeComment", txtComment.Text.Trim());

                    fields.SetField("NoRow1", "1");
                    fields.SetField("NoRow2", "2");
                    fields.SetField("NoRow3", "3");
                    fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                    fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                    fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                    fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                    fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                    fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                    fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                    fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                    fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                    fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                    fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                    fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());

                    stamper.FormFlattening = true;
                    stamper.Close();

                    command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                    MySqlCommand cmd = new MySqlCommand(command, condb);
                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    System.Diagnostics.Process.Start(path);
                    return;
                }
              
              }
            }
           
        }

        public void CountRow4(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }

                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }

                }
            }
        }

        public void CountRow5(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }

                }
 
            }
           
        }

        public void CountRow6(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal1.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External1.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
 
            }
 
        }

        public void CountRow7(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow8(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow9(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);
                int Count2 = Int32.Parse(Count);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow10(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow11(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow12(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External2.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow13(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow14(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow15(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("NoRow15", "15");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("TypeRow15", dt.Rows[14]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("EIANoRow15", dt.Rows[14]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("UMORow15", dt.Rows[14]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());
                        fields.SetField("QTYRow15", dt.Rows[14]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("NoRow15", "15");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("TypeRow15", dt.Rows[14]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("EIANoRow15", dt.Rows[14]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("UMORow15", dt.Rows[14]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());
                        fields.SetField("QTYRow15", dt.Rows[14]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow16(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("NoRow15", "15");
                        fields.SetField("NoRow16", "16");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("TypeRow15", dt.Rows[14]["Model"].ToString());
                        fields.SetField("TypeRow16", dt.Rows[15]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("EIANoRow15", dt.Rows[14]["EIA"].ToString());
                        fields.SetField("EIANoRow16", dt.Rows[15]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("UMORow15", dt.Rows[14]["unit"].ToString());
                        fields.SetField("UMORow16", dt.Rows[15]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());
                        fields.SetField("QTYRow15", dt.Rows[14]["Count"].ToString());
                        fields.SetField("QTYRow16", dt.Rows[15]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("NoRow15", "15");
                        fields.SetField("NoRow16", "16");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("TypeRow15", dt.Rows[14]["Model"].ToString());
                        fields.SetField("TypeRow16", dt.Rows[15]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("EIANoRow15", dt.Rows[14]["EIA"].ToString());
                        fields.SetField("EIANoRow16", dt.Rows[15]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("UMORow15", dt.Rows[14]["unit"].ToString());
                        fields.SetField("UMORow16", dt.Rows[15]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());
                        fields.SetField("QTYRow15", dt.Rows[14]["Count"].ToString());
                        fields.SetField("QTYRow16", dt.Rows[15]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow17(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("NoRow15", "15");
                        fields.SetField("NoRow16", "16");
                        fields.SetField("NoRow17", "17");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("TypeRow15", dt.Rows[14]["Model"].ToString());
                        fields.SetField("TypeRow16", dt.Rows[15]["Model"].ToString());
                        fields.SetField("TypeRow17", dt.Rows[16]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("EIANoRow15", dt.Rows[14]["EIA"].ToString());
                        fields.SetField("EIANoRow16", dt.Rows[15]["EIA"].ToString());
                        fields.SetField("EIANoRow17", dt.Rows[16]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("UMORow15", dt.Rows[14]["unit"].ToString());
                        fields.SetField("UMORow16", dt.Rows[15]["unit"].ToString());
                        fields.SetField("UMORow17", dt.Rows[16]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());
                        fields.SetField("QTYRow15", dt.Rows[14]["Count"].ToString());
                        fields.SetField("QTYRow16", dt.Rows[15]["Count"].ToString());
                        fields.SetField("QTYRow17", dt.Rows[16]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("NoRow15", "15");
                        fields.SetField("NoRow16", "16");
                        fields.SetField("NoRow17", "17");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("TypeRow15", dt.Rows[14]["Model"].ToString());
                        fields.SetField("TypeRow16", dt.Rows[15]["Model"].ToString());
                        fields.SetField("TypeRow17", dt.Rows[16]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("EIANoRow15", dt.Rows[14]["EIA"].ToString());
                        fields.SetField("EIANoRow16", dt.Rows[15]["EIA"].ToString());
                        fields.SetField("EIANoRow17", dt.Rows[16]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("UMORow15", dt.Rows[14]["unit"].ToString());
                        fields.SetField("UMORow16", dt.Rows[15]["unit"].ToString());
                        fields.SetField("UMORow17", dt.Rows[16]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());
                        fields.SetField("QTYRow15", dt.Rows[14]["Count"].ToString());
                        fields.SetField("QTYRow16", dt.Rows[15]["Count"].ToString());
                        fields.SetField("QTYRow17", dt.Rows[16]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }

        public void CountRow18(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

                if (cb == "Internal")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileInternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_Internal3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileInternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("NoRow15", "15");
                        fields.SetField("NoRow16", "16");
                        fields.SetField("NoRow17", "17");
                        fields.SetField("NoRow18", "18");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("TypeRow15", dt.Rows[14]["Model"].ToString());
                        fields.SetField("TypeRow16", dt.Rows[15]["Model"].ToString());
                        fields.SetField("TypeRow17", dt.Rows[16]["Model"].ToString());
                        fields.SetField("TypeRow18", dt.Rows[17]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("EIANoRow15", dt.Rows[14]["EIA"].ToString());
                        fields.SetField("EIANoRow16", dt.Rows[15]["EIA"].ToString());
                        fields.SetField("EIANoRow17", dt.Rows[16]["EIA"].ToString());
                        fields.SetField("EIANoRow18", dt.Rows[17]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("UMORow15", dt.Rows[14]["unit"].ToString());
                        fields.SetField("UMORow16", dt.Rows[15]["unit"].ToString());
                        fields.SetField("UMORow17", dt.Rows[16]["unit"].ToString());
                        fields.SetField("UMORow18", dt.Rows[17]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());
                        fields.SetField("QTYRow15", dt.Rows[14]["Count"].ToString());
                        fields.SetField("QTYRow16", dt.Rows[15]["Count"].ToString());
                        fields.SetField("QTYRow17", dt.Rows[16]["Count"].ToString());
                        fields.SetField("QTYRow18", dt.Rows[17]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
                else if (cb == "External")
                {
                    iniconfig = new IniFile(Application.StartupPath + "\\Config.ini");
                    string IP = iniconfig.IniReadValue("MySQL_Server", "IP");
                    string DB = iniconfig.IniReadValue("MySQL_Server", "DB");

                    string ConnectionString = "Server=" + IP + ";";
                    ConnectionString += "Uid=root;";
                    ConnectionString += "Password=123456*;";
                    ConnectionString += "Database=" + DB + ";";
                    condb = new MySqlConnection(ConnectionString);
                    condb.Open();

                    string fromfileExternal = System.AppDomain.CurrentDomain.BaseDirectory + @"\doc\baseFile\SAMPLE_SET_AGREEMENT_External3.pdf";
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF Files (*.pdf)|*.pdf";
                    save.FileName = DocumentNumber;
                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string path = save.FileName;

                        PdfReader reader = new PdfReader(fromfileExternal);
                        PdfStamper stamper = new PdfStamper(reader, new FileStream(path, FileMode.Create));
                        AcroFields fields = stamper.AcroFields;

                        fields.SetField("writeDocumentNo", DocumentNumber);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeENNo", ID);
                        fields.SetField("writeNameSurname", namesurname);
                        fields.SetField("writeDepartment", Department);
                        fields.SetField("writeBorrowDate", DateBorrow);
                        fields.SetField("writeComment", txtComment.Text.Trim());

                        fields.SetField("NoRow1", "1");
                        fields.SetField("NoRow2", "2");
                        fields.SetField("NoRow3", "3");
                        fields.SetField("NoRow4", "4");
                        fields.SetField("NoRow5", "5");
                        fields.SetField("NoRow6", "6");
                        fields.SetField("NoRow7", "7");
                        fields.SetField("NoRow8", "8");
                        fields.SetField("NoRow9", "9");
                        fields.SetField("NoRow10", "10");
                        fields.SetField("NoRow11", "11");
                        fields.SetField("NoRow12", "12");
                        fields.SetField("NoRow13", "13");
                        fields.SetField("NoRow14", "14");
                        fields.SetField("NoRow15", "15");
                        fields.SetField("NoRow16", "16");
                        fields.SetField("NoRow17", "17");
                        fields.SetField("NoRow18", "18");
                        fields.SetField("TypeRow1", dt.Rows[0]["Model"].ToString());
                        fields.SetField("TypeRow2", dt.Rows[1]["Model"].ToString());
                        fields.SetField("TypeRow3", dt.Rows[2]["Model"].ToString());
                        fields.SetField("TypeRow4", dt.Rows[3]["Model"].ToString());
                        fields.SetField("TypeRow5", dt.Rows[4]["Model"].ToString());
                        fields.SetField("TypeRow6", dt.Rows[5]["Model"].ToString());
                        fields.SetField("TypeRow7", dt.Rows[6]["Model"].ToString());
                        fields.SetField("TypeRow8", dt.Rows[7]["Model"].ToString());
                        fields.SetField("TypeRow9", dt.Rows[8]["Model"].ToString());
                        fields.SetField("TypeRow10", dt.Rows[9]["Model"].ToString());
                        fields.SetField("TypeRow11", dt.Rows[10]["Model"].ToString());
                        fields.SetField("TypeRow12", dt.Rows[11]["Model"].ToString());
                        fields.SetField("TypeRow13", dt.Rows[12]["Model"].ToString());
                        fields.SetField("TypeRow14", dt.Rows[13]["Model"].ToString());
                        fields.SetField("TypeRow15", dt.Rows[14]["Model"].ToString());
                        fields.SetField("TypeRow16", dt.Rows[15]["Model"].ToString());
                        fields.SetField("TypeRow17", dt.Rows[16]["Model"].ToString());
                        fields.SetField("TypeRow18", dt.Rows[17]["Model"].ToString());
                        fields.SetField("EIANoRow1", dt.Rows[0]["EIA"].ToString());
                        fields.SetField("EIANoRow2", dt.Rows[1]["EIA"].ToString());
                        fields.SetField("EIANoRow3", dt.Rows[2]["EIA"].ToString());
                        fields.SetField("EIANoRow4", dt.Rows[3]["EIA"].ToString());
                        fields.SetField("EIANoRow5", dt.Rows[4]["EIA"].ToString());
                        fields.SetField("EIANoRow6", dt.Rows[5]["EIA"].ToString());
                        fields.SetField("EIANoRow7", dt.Rows[6]["EIA"].ToString());
                        fields.SetField("EIANoRow8", dt.Rows[7]["EIA"].ToString());
                        fields.SetField("EIANoRow9", dt.Rows[8]["EIA"].ToString());
                        fields.SetField("EIANoRow10", dt.Rows[9]["EIA"].ToString());
                        fields.SetField("EIANoRow11", dt.Rows[10]["EIA"].ToString());
                        fields.SetField("EIANoRow12", dt.Rows[11]["EIA"].ToString());
                        fields.SetField("EIANoRow13", dt.Rows[12]["EIA"].ToString());
                        fields.SetField("EIANoRow14", dt.Rows[13]["EIA"].ToString());
                        fields.SetField("EIANoRow15", dt.Rows[14]["EIA"].ToString());
                        fields.SetField("EIANoRow16", dt.Rows[15]["EIA"].ToString());
                        fields.SetField("EIANoRow17", dt.Rows[16]["EIA"].ToString());
                        fields.SetField("EIANoRow18", dt.Rows[17]["EIA"].ToString());
                        fields.SetField("UMORow1", dt.Rows[0]["unit"].ToString());
                        fields.SetField("UMORow2", dt.Rows[1]["unit"].ToString());
                        fields.SetField("UMORow3", dt.Rows[2]["unit"].ToString());
                        fields.SetField("UMORow4", dt.Rows[3]["unit"].ToString());
                        fields.SetField("UMORow5", dt.Rows[4]["unit"].ToString());
                        fields.SetField("UMORow6", dt.Rows[5]["unit"].ToString());
                        fields.SetField("UMORow7", dt.Rows[6]["unit"].ToString());
                        fields.SetField("UMORow8", dt.Rows[7]["unit"].ToString());
                        fields.SetField("UMORow9", dt.Rows[8]["unit"].ToString());
                        fields.SetField("UMORow10", dt.Rows[9]["unit"].ToString());
                        fields.SetField("UMORow11", dt.Rows[10]["unit"].ToString());
                        fields.SetField("UMORow12", dt.Rows[11]["unit"].ToString());
                        fields.SetField("UMORow13", dt.Rows[12]["unit"].ToString());
                        fields.SetField("UMORow14", dt.Rows[13]["unit"].ToString());
                        fields.SetField("UMORow15", dt.Rows[14]["unit"].ToString());
                        fields.SetField("UMORow16", dt.Rows[15]["unit"].ToString());
                        fields.SetField("UMORow17", dt.Rows[16]["unit"].ToString());
                        fields.SetField("UMORow18", dt.Rows[17]["unit"].ToString());
                        fields.SetField("QTYRow1", dt.Rows[0]["Count"].ToString());
                        fields.SetField("QTYRow2", dt.Rows[1]["Count"].ToString());
                        fields.SetField("QTYRow3", dt.Rows[2]["Count"].ToString());
                        fields.SetField("QTYRow4", dt.Rows[3]["Count"].ToString());
                        fields.SetField("QTYRow5", dt.Rows[4]["Count"].ToString());
                        fields.SetField("QTYRow6", dt.Rows[5]["Count"].ToString());
                        fields.SetField("QTYRow7", dt.Rows[6]["Count"].ToString());
                        fields.SetField("QTYRow8", dt.Rows[7]["Count"].ToString());
                        fields.SetField("QTYRow9", dt.Rows[8]["Count"].ToString());
                        fields.SetField("QTYRow10", dt.Rows[9]["Count"].ToString());
                        fields.SetField("QTYRow11", dt.Rows[10]["Count"].ToString());
                        fields.SetField("QTYRow12", dt.Rows[11]["Count"].ToString());
                        fields.SetField("QTYRow13", dt.Rows[12]["Count"].ToString());
                        fields.SetField("QTYRow14", dt.Rows[13]["Count"].ToString());
                        fields.SetField("QTYRow15", dt.Rows[14]["Count"].ToString());
                        fields.SetField("QTYRow16", dt.Rows[15]["Count"].ToString());
                        fields.SetField("QTYRow17", dt.Rows[16]["Count"].ToString());
                        fields.SetField("QTYRow18", dt.Rows[17]["Count"].ToString());

                        stamper.FormFlattening = true;
                        stamper.Close();

                        command = "insert into agreement_document (DOC_NO,PERSONAL_ID,PERSONAL_NAME,DEPARTMENT,BORROW_DATE,PURPOSE,flg_status,rec_date) values('" + DocumentNumber + "','" + ID + "', '" + namesurname + "', '" + Department + "','" + DateBorrow + "', '" + cb + "','Pending',sysdate());";
                        MySqlCommand cmd = new MySqlCommand(command, condb);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Save PDF Complete! ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        System.Diagnostics.Process.Start(path);
                        return;
                    }
                }
            }
        }
 
        public void exportPDF(DataTable dt)
        {

            foreach (DataRow row in dt.Rows)
            {
                DocumentNumber = Convert.ToString(row["DocumentNo"]);
                ID = Convert.ToString(row["ScanID"]);
                namesurname = Convert.ToString(row["Name"]);
                Department = Convert.ToString(row["Department"]);
                Model = Convert.ToString(row["Model"]);
                EIA = Convert.ToString(row["EIA"]);
                Count = Convert.ToString(row["Count"]);
                cb = Convert.ToString(row["checkBox"]);
                Unit = Convert.ToString(row["unit"]);
                DateBorrow = Convert.ToString(row["DateBorrow"]);

            }

            if (dt.Rows.Count == 1)
            {
                CountRow1(dt);
            }
            else if (dt.Rows.Count == 2)
            {
                CountRow2(dt);
            }
            else if (dt.Rows.Count == 3)
            {
                CountRow3(dt);

            }
            else if (dt.Rows.Count == 4)
            {
                CountRow4(dt);
            }
            else if (dt.Rows.Count == 5)
            {
                CountRow5(dt);
            }
            else if (dt.Rows.Count == 6)
            {
                CountRow6(dt);
            }
            else if (dt.Rows.Count == 7)
            {
                CountRow7(dt);
            }
            else if (dt.Rows.Count == 8)
            {
                CountRow8(dt);
            }
            else if (dt.Rows.Count == 9)
            {
                CountRow9(dt);
            }
            else if (dt.Rows.Count == 10)
            {
                CountRow10(dt);
            }
            else if (dt.Rows.Count == 11)
            {
                CountRow11(dt);
            }
            else if (dt.Rows.Count == 12)
            {
                CountRow12(dt);
            }
            else if (dt.Rows.Count == 13)
            {
                CountRow13(dt);
            }
            else if (dt.Rows.Count == 14)
            {
                CountRow14(dt);
            }
            else if (dt.Rows.Count == 15)
            {
                CountRow15(dt);
            }
            else if (dt.Rows.Count == 16)
            {
                CountRow16(dt);
            }
            else if (dt.Rows.Count == 17)
            {
                CountRow17(dt);
            }
            else if (dt.Rows.Count == 18)
            {
                CountRow18(dt);
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

    }
}
