using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;

namespace MASTER_SET_CONTROL
{
    public class ExportToExcel
    {

        public void export2excel(DataTable dtGridview)
        {
            string[] Arr = new string[100];
            int s, k = 0;

            if (dtGridview.Rows.Count > 0)
            {
                foreach (DataColumn column in dtGridview.Columns)
                {
                    Arr[k] = column.ColumnName;
                    k++;
                }

                Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();

                XcelApp.Application.Workbooks.Add(Type.Missing);

                for (s = 0; s < dtGridview.Columns.Count; s++)
                {
                    XcelApp.Cells[1, s + 1] = Arr[s];
                }

                for (int i = 0; i < dtGridview.Rows.Count; i++)
                {
                    for (int j = 0; j < dtGridview.Columns.Count; j++)
                    {
                        XcelApp.Cells[i + 2, j + 1] = (dtGridview.Rows[i][j] ?? "").ToString();
                    }
                }

                XcelApp.Columns.AutoFit();
                XcelApp.Visible = true;
            }
        }
    }


}
