using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace excelTemplate
{
    public partial class mainForm : Form
    {
        public mainForm()
        {
            InitializeComponent();

        }

        List<List<string>> excelData = new List<List<string>>();

        private void btBrowseExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog browseExcel = new OpenFileDialog();

            browseExcel.Title = "Open Contact Excel File";
            browseExcel.Filter = "xls files|*.xls";

            if (browseExcel.ShowDialog() == DialogResult.OK)
            {
                excelData.Clear();

                tbBrowseExcel.Text = browseExcel.FileName.ToString();

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(tbBrowseExcel.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);               

                Excel.Range range = xlWorkSheet.UsedRange;
                //Excel.Range range = xlWorkSheet.get_Range("A1", last);

                int excelRow = range.Rows.Count;
                int excelCol = range.Columns.Count;

                int lastRow = xlWorkSheet.Cells[excelRow, 1].End(Excel.XlDirection.xlUp).Row;                

                for (int i=4; i<=lastRow; i++) // Start index at 4th row
                {
                    List<string> subData = new List<string>();

                    for (int j=1; j<=excelCol; j++)
                    {
                        subData.Add(Convert.ToString(xlWorkSheet.Cells[i, j].Value2));
                    }

                    excelData.Add(subData);
                }

                //Clean up
                xlWorkBook.Close(false);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

            }
        }

        private void btBrowseTemplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog browseTemplate = new OpenFileDialog();

            browseTemplate.Title = "Open Template File";
            browseTemplate.Filter = "docx files|*.docx";

            if (browseTemplate.ShowDialog() == DialogResult.OK)
            {
                tbBrowseTemplate.Text = browseTemplate.FileName.ToString();
            }
        }

        private void btSearchID_Click(object sender, EventArgs e)
        {
            int searchIndex = excelData.Count + 1;

            for(int i=0; i<excelData.Count; i++)
            {
                if(excelData[i][1].ToString() == tbSearchID.Text)
                {
                    searchIndex = i;
                    break;
                }
            }

            if(searchIndex == excelData.Count + 1)
            {
                MessageBox.Show("ไม่พบข้อมูล");
            }

            else
            {

            }
        }
    }
}
