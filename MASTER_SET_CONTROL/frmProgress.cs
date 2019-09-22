using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MASTER_SET_CONTROL
{
    public partial class PleaseWait : Form
    {
        private static PleaseWait mInstance;
        public static void Create()
        {
            var t = new System.Threading.Thread(() =>
            {
                mInstance = new PleaseWait();
                mInstance.FormClosed += (s, e) => mInstance = null;
                Application.Run(mInstance);
            });
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.IsBackground = true;
            t.Start();
        }

        public static void Destroy()
        {
            if (mInstance != null) mInstance.Invoke(new Action(() => mInstance.Close()));
            
        }

        public PleaseWait()
        {
            InitializeComponent();
        }
    }
}
