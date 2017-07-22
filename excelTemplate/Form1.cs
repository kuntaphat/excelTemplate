using Novacode;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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

        CultureInfo ThaiCulture = new CultureInfo("th-TH");
        CultureInfo UsaCulture = new CultureInfo("en-US");

        List<List<string>> excelData = new List<List<string>>();
        string witnessSign;

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
                        if (j == 17)
                        {
                            DateTime dateTime = Convert.ToDateTime(xlWorkSheet.Cells[i, j].Value2);

                            //double date = double.Parse(dateTime);

                            var conv = dateTime.ToString("dd MMMM yyyy", ThaiCulture);

                            subData.Add(conv);
                        }

                        else
                        {
                            subData.Add(Convert.ToString(xlWorkSheet.Cells[i, j].Value2));
                        }                       
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

        private void btSearchID_Click(object sender, EventArgs e)
        {
            clearResult();

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
                tbRegisID.Text = excelData[searchIndex][1];
                tbIDCard.Text = excelData[searchIndex][2];
                tbNameTitle.Text = excelData[searchIndex][3];
                tbName.Text = excelData[searchIndex][4];
                tbSurname.Text = excelData[searchIndex][5];
                tbPosition.Text = excelData[searchIndex][6];

                tbAddress.Text = excelData[searchIndex][7];
                tbAddress2.Text = excelData[searchIndex][8];
                tbAddress3.Text = excelData[searchIndex][9];
                tbAddress4.Text = excelData[searchIndex][10];
                tbZipCode.Text = excelData[searchIndex][11];

                tbTel.Text = excelData[searchIndex][12];
                tbDepartment.Text = excelData[searchIndex][13];
                tbVice.Text = excelData[searchIndex][14];
                tbWorkCode.Text = excelData[searchIndex][15];
                tbStartDate.Text = excelData[searchIndex][16];
                tbID.Text = excelData[searchIndex][17];

                DateTime now = DateTime.Today;

                DateTime birthDate = Convert.ToDateTime(excelData[searchIndex][27]).AddYears(-543);

                int birthyear = Convert.ToInt32(birthDate.ToString("yyyy"));

                int age = now.Year - birthyear;

                if (now < birthDate.AddYears(age)) age--;

                tbAge.Text = age.ToString();
                tbWitness1.Text = excelData[searchIndex][24];
                tbWitness2.Text = excelData[searchIndex][25];
                witnessSign = excelData[searchIndex][26];

                tbEmName.Text = excelData[searchIndex][19];
                tbEmPosition.Text = excelData[searchIndex][21];
                tbEmWorkCode.Text = excelData[searchIndex][23];

            }
        }

        private void clearResult()
        {
            tbRegisID.Clear();
            tbIDCard.Clear();
            tbNameTitle.Clear();
            tbName.Clear();
            tbSurname.Clear();
            tbPosition.Clear();

            tbAddress.Clear();
            tbAddress2.Clear();
            tbAddress3.Clear();
            tbAddress4.Clear();
            tbZipCode.Clear();

            tbTel.Clear();
            tbDepartment.Clear();
            tbVice.Clear();
            tbWorkCode.Clear();
            tbStartDate.Clear();
            tbID.Clear();

            tbAge.Clear();
            tbSection.Clear();
            tbDivision.Clear();
            tbWitness1.Clear();
            tbWitness2.Clear();

            tbEmName.Clear();
            tbEmPosition.Clear();
            tbEmWorkCode.Clear();
        }

        private void btCreateContact_Click(object sender, EventArgs e)
        {
            SaveFileDialog browseContact = new SaveFileDialog();

            browseContact.Title = "Open Template File";
            browseContact.Filter = "docx files|*.docx";

            if (browseContact.ShowDialog() == DialogResult.OK)
            {
                DocX document = DocX.Load(tbBrowseTemplate.Text);

                document.ReplaceText("work", "การไฟฟ้าฝ่ายผลิตแห่งประเทศไทย");

                string[] startWork = tbStartDate.Text.Split(' ');

                document.ReplaceText("startDate", startWork[0]);
                document.ReplaceText("startMonth", startWork[1]);
                document.ReplaceText("startYear", startWork[2]);

                document.ReplaceText("bossName", tbEmName.Text);
                document.ReplaceText("bossPos", tbEmPosition.Text);

                document.ReplaceText("emTitle", tbNameTitle.Text);
                document.ReplaceText("emName", tbName.Text);
                document.ReplaceText("emSurname", tbSurname.Text);
                document.ReplaceText("emAge", tbAge.Text);

                string[] address1 = tbAddress.Text.Split(new[] { "หมู่ที่" }, StringSplitOptions.None);     //บ้านเลขที่
                string[] address2 = address1[1].Split(new[] { "ตรอก/ซอย" }, StringSplitOptions.None);    //หมู่ที่
                string[] address3 = address2[1].Split(new[] { "ถนน" }, StringSplitOptions.None);        //ซอย, ถนน


                document.ReplaceText("emHomeNo", address1[0]);
                document.ReplaceText("emVillageNo", address2[0]);
                document.ReplaceText("emSoi", address3[0]);
                document.ReplaceText("emStreet", address3[1]);
                document.ReplaceText("em2Address", tbAddress2.Text);
                document.ReplaceText("em3Address", tbAddress3.Text);
                document.ReplaceText("emProvince", tbAddress4.Text);
                document.ReplaceText("emZipCode", tbZipCode.Text);

                document.ReplaceText("emTel", tbTel.Text);
                document.ReplaceText("emPos", tbPosition.Text);
                document.ReplaceText("emNo", tbID.Text);
                document.ReplaceText("emDepartment", tbDepartment.Text);
                document.ReplaceText("emWorkCode", tbWorkCode.Text);

                document.ReplaceText("bossSign", tbEmName.Text);
                document.ReplaceText("bossPos", tbEmPosition.Text);

                document.ReplaceText("emX", tbNameTitle.Text);
                document.ReplaceText("emY", tbName.Text);
                document.ReplaceText("emZ", tbSurname.Text);

                document.ReplaceText("witness1", tbWitness1.Text);
                document.ReplaceText("witness2", tbWitness2.Text);
                document.ReplaceText("wit2Sign", witnessSign);

                document.SaveAs(browseContact.FileName);

                MessageBox.Show("สร้างสัญญาเสร็จเรียบร้อย");

                document.
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
    }
}
