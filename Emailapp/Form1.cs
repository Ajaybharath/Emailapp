using System;
using System.Data;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Mail;
using System.Threading;

namespace Emailapp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int i;
        System.Data.DataTable dt;
        private void button1_Click(object sender, EventArgs e)
        {
            string file = ""; //variable for the Excel File Location
            //container for our excel data
            DataRow row;
            dt = new System.Data.DataTable();
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
                    for (i = 0; i < dt.Rows.Count; i++)
                    {  
                        dt.Rows[i]["Time"] = DateTime.FromOADate(Convert.ToDouble(dt.Rows[i]["Time"].ToString())).ToString("h\\:mm");
                    }
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
            progressBar1.Minimum = 0;
            progressBar1.Maximum = dt.Rows.Count;
            //progressBar1.Step = 1;
            MailMessage mailMsg = new MailMessage();
            mailMsg.From = new MailAddress("ajaybharath009@gmail.com", "IB IoT");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                mailMsg.To.Add(new MailAddress(dt.Rows[i]["Emailid"].ToString()));
                mailMsg.Subject = "Reg: Ideabytes Interview Call";
                mailMsg.IsBodyHtml = true;
                mailMsg.Body = $"Dear {dt.Rows[i]["Name"]},<br/><br/>You are shortlisted for {dt.Rows[i]["Position"]}, as we discussed your interview was scheduled on {dt.Rows[i]["Date"]} at {dt.Rows[i]["Time"]}.<br/><br/>" + @"<img src='https://adminiot.dgtrak.online/FTP_Sensorcnt/domain/IBThanks.png'/>"
                + "<br/><br/> With Regards <br/> <hr/>HR Manager<br/>Ideabytes Inc<br/>Website: www.ideabytes.com<br/><br/>" + "<img src='https://adminiot.dgtrak.online/FTP_Sensorcnt/domain/IBMailImage.png'/>"
                + "<br/><br/>Important: This email and any files transmitted with it are confidential and intended solely for the use of the individual or entity to whom they are addressed. If you have received this email in error please notify the system manager. " +
                "Please notify the sender immediately by e-mail if you have received this e-mail by mistake and delete this e-mail from your system. If you are not the intended recipient you are notified that disclosing, copying, distributing or taking any action in " +
                "reliance on the contents of this information is strictly prohibited.";
                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential("ajaybharath009@gmail.com", "ajay009@1234");
                smtp.EnableSsl = true;
                for (int j = 0; j <= i; j++)
                {
                    label2.Text = $"Dear {dt.Rows[i]["Name"]},You are shortlisted for {dt.Rows[i]["Position"]}, as we discussed your interview was scheduled on {dt.Rows[i]["Date"]} at {dt.Rows[i]["Time"]}";
                }
                smtp.Send(mailMsg);
                mailMsg.To.Clear();
                progressBar1.Value = i+1;
                Thread.Sleep(1000);
            }
            MessageBox.Show("Mail Sent Successfully!!");
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //dataGridView1.Visible = true;
        }

    }
}
