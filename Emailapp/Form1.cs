using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Mail;

namespace Emailapp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int i;
        private void button1_Click(object sender, EventArgs e)
        {
            //string filePath = string.Empty;
            //string fileExt = string.Empty;
            //OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            //string fname = "";
            //OpenFileDialog fdlg = new OpenFileDialog();
            //fdlg.Title = "Excel File Dialog";
            //fdlg.InitialDirectory = @"c:\";
            //fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            //fdlg.FilterIndex = 2;
            //fdlg.RestoreDirectory = true;
            //    //fname = fdlg.FileName;
            //    if (fdlg.ShowDialog() == DialogResult.OK)
            //    {
            //        fname = fdlg.FileName;
            //    }
            //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //    Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            //    Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            //    Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
            //    int rowCount = xlRange.Rows.Count;
            //    int colCount = xlRange.Columns.Count;
            //// dt.Column = colCount;  
            //    System.Data.DataTable d1 = new System.Data.DataTable();
            //    dataGridView1.ColumnCount = colCount;
            //    dataGridView1.RowCount = rowCount;
            //    d1.Columns.AddRange(new DataColumn[5]{ new DataColumn("Name", typeof(string)),new DataColumn("Emailid", typeof(string)),new DataColumn("Position", typeof(string)),new DataColumn("Date", typeof(DateTime)),new DataColumn("Time",typeof(DateTime)) });
            //    for (int i = 1; i <= rowCount; i++)
            //    {
            //        for (int j = 1; j <= colCount; j++)
            //        {


            //            //write the value to the Grid  


            //            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            //            {
            //                dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();

            //            }
            //            // Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");  

            //            //add useful things here!     
            //        }
            //    }

            //    //cleanup  
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();
            //    Marshal.ReleaseComObject(xlRange);
            //    Marshal.ReleaseComObject(xlWorksheet);

            //    //close and release  
            //    xlWorkbook.Close();
            //    Marshal.ReleaseComObject(xlWorkbook);

            //    //quit and release  
            //    xlApp.Quit();
            //    Marshal.ReleaseComObject(xlApp);
            string file = ""; //variable for the Excel File Location
            System.Data.DataTable dt = new System.Data.DataTable(); //container for our excel data
            DataRow row;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Check if Result == "OK".
            {
                file = openFileDialog1.FileName; //get the filename with the location of the file
                try
                {
                    //Create Object for Microsoft.Office.Interop.Excel that will be use to read excel file

                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                    Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(file);

                    Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];

                    Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

                    int rowCount = excelRange.Rows.Count; //get row count of excel data

                    int colCount = excelRange.Columns.Count; // get column count of excel data

                    //Get the first Column of excel file which is the Column Name

                    for (i = 1; i <= rowCount; i++)
                    {
                        for (int j = 1; j <= colCount; j++)
                        {
                            dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                        }
                        break;
                    }

                    //Get Row Data of Excel

                    int rowCounter; //This variable is used for row index number
                    for (int i = 2; i <= rowCount; i++) //Loop for available row of excel data
                    {
                        row = dt.NewRow(); //assign new row to DataTable
                        rowCounter = 0;
                        for (int j = 1; j <= colCount; j++) //Loop for available column of excel data
                        {
                            //check if cell is empty
                            if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                            {
                                row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                            }
                            else
                            {
                                row[i] = "";
                            }
                            rowCounter++;
                        }
                        dt.Rows.Add(row); //add row to DataTable
                    }

                    dataGridView1.DataSource = dt; //assign DataTable as Datasource for DataGridview

                    //close and clean excel process
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.ReleaseComObject(excelRange);
                    Marshal.ReleaseComObject(excelWorksheet);
                    //quit apps
                    excelWorkbook.Close();
                    Marshal.ReleaseComObject(excelWorkbook);
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                label1.Text = "Excel File Uploaded";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            MailMessage mailMsg = new MailMessage();
            mailMsg.From = new MailAddress("ajaybharath009@gmail.com", "IB IoT");
            int count = dataGridView1.Rows.Count - 1;
            for (int i = 1; i < dataGridView1.Rows.Count; i++)
            {
                mailMsg.To.Add(new MailAddress(dataGridView1.Rows[i].Cells[1].Value.ToString()));
            }
            mailMsg.Subject = "Testing Mail";
            mailMsg.Body = "Hiii";
            mailMsg.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential("ajaybharath009@gmail.com", "ajay009@1234");
            smtp.EnableSsl = true;
            smtp.Send(mailMsg);
            // mailMsg.To.Clear();
            MessageBox.Show("Mail Sent Successfully!!");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Visible = true;
        }
    }
}
