using System;
using System.Data;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.IO;
using System.Configuration;
using Newtonsoft.Json;
using System.Collections.Generic;

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
        List<string> numbers = new List<string>();
        private void button1_Click(object sender, EventArgs e)
        {
            string file = "";
            DataRow row;
            dt = new System.Data.DataTable();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            DialogResult result = openFileDialog1.ShowDialog(); 
            if (result == DialogResult.OK) 
            {
                label1.Text = "";
                richTextBox1.Text = string.Empty;
                file = openFileDialog1.FileName;
                string filename = Path.GetFileName(file);
                try
                {
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(file);
                    Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                    Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                    int rowCount = excelRange.Rows.Count; 
                    int colCount = excelRange.Columns.Count;
                    numbers = new List<string>();
                    dt.Clear();
                    for (i = 1; i <= rowCount; i++)
                    {
                        for (int j = 1; j <= colCount; j++)
                        {
                            dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                        }
                        break;
                    }
                    int rowCounter; 
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
                        numbers.Add(dt.Rows[i]["Mobile No"].ToString());
                    }
                    //for (i = 0; i < dt.Rows.Count; i++)
                    //{  
                    //  dt.Rows[i]["Date"] = DateTime.FromOADate(Convert.ToDouble(dt.Rows[i]["Date"].ToString())).ToString("dd/MM/yyyy");
                    //  dt.Rows[i]["Time"] = DateTime.FromOADate(Convert.ToDouble(dt.Rows[i]["Time"].ToString())).ToString("h\\:mm");
                    //}
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
                    label1.Text = filename;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                button1.Enabled = false;
                button2.Enabled = false;
                if (label1.Text != "")
                {
                    DialogResult dr = MessageBox.Show("Do you really want to send Mails", "Verification", MessageBoxButtons.YesNo);
                    if (dr == DialogResult.Yes)
                    {
                        progressBar1.Visible = true;
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dt.Rows.Count;
                        richTextBox1.Text = string.Empty;
                        MailMessage mailMsg = new MailMessage();
                        mailMsg.From = new MailAddress("ajaybharath009@gmail.com", "IB IoT");
                        mailMsg.Subject = "Reg: Ideabytes Interview Call";
                        for (i = 0; i < dt.Rows.Count; i++)
                        {
                            Thread.Sleep(1000);
                            mailMsg.To.Add(new MailAddress(dt.Rows[i]["Emailid"].ToString()));
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
                            smtp.Send(mailMsg);
                            mailMsg.To.Clear();
                            //mobile massages
                            msg91 m91 = new msg91();
                            m91.country = "91";
                            m91.sender = "IDESOF";
                            m91.route = "4";
                            m91.sms = new List<msg>();
                            m91.DLT_TE_ID = "1307163756009502008";
                            msg m = new msg();
                            m.to = numbers[i].Split(',');
                            m.message = $"\n\nDear {dt.Rows[i]["Name"]},You are shortlisted for {dt.Rows[i]["Position"]},as we discussed\n your interview was scheduled on {dt.Rows[i]["Date"]} at {dt.Rows[i]["Time"]} Thanks - Ideabytes";
                            m91.sms.Add(m);
                            sendingSMS(m91);
                            progressBar1.Value = i + 1;
                            richTextBox1.Text += $"\n\nDear {dt.Rows[i]["Name"]},You are shortlisted for {dt.Rows[i]["Position"]},as we discussed\n your interview was scheduled on {dt.Rows[i]["Date"]} at {dt.Rows[i]["Time"]}.";
                            richTextBox1.Update();
                        }
                        DialogResult dr1 = MessageBox.Show("Mail Sent Successfully!!", "", MessageBoxButtons.OK);
                        if (dr1 == DialogResult.OK)
                        {
                            progressBar1.Value = 0;
                            label2.Text = "";
                            label1.Text = "";
                            progressBar1.Visible = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Mails sending Cancelled!!!");
                        label1.Text = "";
                    }
                }
                else
                {
                    MessageBox.Show("Please Upload Excel File");
                }
                button1.Enabled = true;
                button2.Enabled = true;
            }
            catch (Exception ex1)
            {
                ex1= null;
                MessageBox.Show("The Mails sending was not Completed!!!!");
            }
        }
        public void sendingSMS(msg91 m91)
        {
            try
            {
                string smsKey = Convert.ToString(ConfigurationManager.AppSettings["SMSkey"]);
                string smsUrl = Convert.ToString(ConfigurationManager.AppSettings["SMSapi"]);




                if (!string.IsNullOrEmpty(m91.sms[0].message))
                {
                    string DATA = JsonConvert.SerializeObject(m91);




                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(smsUrl);
                    request.Method = "POST";



                    request.ContentType = "application/json";
                    request.Headers.Add("authkey", smsKey);



                    request.ContentLength = DATA.Length;



                    using (Stream webStream = request.GetRequestStream())
                    using (StreamWriter requestWriter = new StreamWriter(webStream, System.Text.Encoding.ASCII))
                    {
                        requestWriter.Write(DATA);
                    }



                    HttpWebResponse webResp = (HttpWebResponse)request.GetResponse();



                    if (webResp.StatusCode == HttpStatusCode.OK)
                    {
                        StreamReader responseReader = new StreamReader(webResp.GetResponseStream());
                        string responseData = string.Empty;
                        responseData += responseReader.ReadToEnd();
                        StringReader stream = null;
                        stream = new StringReader(responseData);

                    }
                }
            }
            catch (Exception ex)
            {
                ex = null;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Visible = false;
        }
    }
}
