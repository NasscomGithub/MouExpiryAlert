using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Timers;

namespace MOU_Expiry_Alert
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer1 = null;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
                timer1 = new Timer();
            this.timer1.Interval = 86000000;//15000000;//25000000;//Convert.ToInt32(ConfigurationManager.ConnectionStrings["Timer"].ToString());//80000000; //every 22 hours
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Tick);
                timer1.Enabled = true;
                Library.WriteErrorLog("MOU Alert Window Service Started");
            
        }
        private void timer1_Tick(object sender, ElapsedEventArgs e)
        {
            //Write code here to do some job depends on your requirement
            ExportDataSetToExcel();
            
            Library.WriteErrorLog("MOU Alert Timer ticked and job has been done successfully");
        }
        protected override void OnStop()
        {
            timer1.Enabled = false;
            Library.WriteErrorLog("MOU Alert Window Service Stopped");
        }
        #region SQL Connections
        public SqlConnection openDBConnection()
        {
            SqlConnection sqlConnection = null;
            try
            {
                String connectionString = ConfigurationManager.ConnectionStrings["connString"].ToString();
                sqlConnection = new SqlConnection(connectionString);
                sqlConnection.Open();
            }
            catch (Exception ex)
            {
                Library.WriteErrorLog(ex.Message);
            }
            return sqlConnection;
        }
        private void closeDBConnection(SqlConnection sqlConnection)
        {
            try
            {
                if (sqlConnection != null)
                    if (sqlConnection.State == System.Data.ConnectionState.Open)
                        sqlConnection.Close();
            }
            catch (Exception ex) { Library.WriteErrorLog(ex.Message); }
            finally { }
            try
            {
                if (sqlConnection != null)
                    sqlConnection.Dispose();
            }
            catch (Exception ex) { Library.WriteErrorLog(ex.Message); }
            finally { }
        }
        public void disposeSqlCommand(SqlCommand sqlCommand)
        {
            try
            {
                if (sqlCommand != null)
                    sqlCommand.Dispose();
            }
            catch (Exception ex) { Library.WriteErrorLog(ex.Message); }
        }
        #endregion

        #region GettingExpiryDate
        public DataSet GetLeagcyScoresFromDB()
        {
            DataSet dt = new DataSet();
            SqlConnection sqlConnection = null;
            try
            {
                sqlConnection = openDBConnection();
                SqlCommand sqlCommand = null;
                String commandStr;
                commandStr = "Sp_Get_MOUExpiryAlerts";
                sqlCommand = new SqlCommand(commandStr, sqlConnection);
                sqlCommand.CommandType = CommandType.StoredProcedure;
                sqlCommand.Parameters.Clear();
                SqlDataAdapter daTP = new SqlDataAdapter(sqlCommand);
                daTP.Fill(dt);
                disposeSqlCommand(sqlCommand);
                sqlCommand = null;
            }
            catch (Exception ex) { Library.WriteErrorLog(ex.Message); }
            closeDBConnection(sqlConnection);
            return dt;
        }

        public DataSet GetStateLeadMailIdFromDB(string state)
        {
            DataSet dt = new DataSet();
            SqlConnection sqlConnection = null;
            try
            {
                sqlConnection = openDBConnection();
                SqlCommand sqlCommand = null;
                String commandStr;
                commandStr = "Select EmailId from MOULeads where State=@State";
                sqlCommand = new SqlCommand(commandStr, sqlConnection);
                sqlCommand.CommandType = CommandType.Text;
                sqlCommand.Parameters.Clear();
                sqlCommand.Parameters.Add(new SqlParameter("@State", SqlDbType.NVarChar, 200)).Value = state;
                SqlDataAdapter daTP = new SqlDataAdapter(sqlCommand);
                daTP.Fill(dt);
                disposeSqlCommand(sqlCommand);
                sqlCommand = null;
            }
            catch (Exception ex) { Library.WriteErrorLog(ex.Message); }
            closeDBConnection(sqlConnection);
            return dt;
        }

        public DataTable FilterResult()
        {

            DataSet ds = GetLeagcyScoresFromDB();
            DataTable dt = new DataTable();
            List<MOUList> list = new List<MOUList>();
            try
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    DataRow row = ds.Tables[0].Rows[i];
                    MOUList mou = new MOUList();
                    if (Convert.ToInt32(row["ExpiryInDays"]) < 30 && Convert.ToInt32(row["ExpiryInDays"]) != -42743 && !String.IsNullOrEmpty(Convert.ToString(ds.Tables[0].Rows[i]["State"])))//&&  Convert.ToInt32(row["ExpiryInDays"]) < 60
                    {
                        mou.Project = Convert.ToString(ds.Tables[0].Rows[i]["Project"]);
                        mou.ProjectIncharge = Convert.ToString(ds.Tables[0].Rows[i]["ProjectIncharge"]);
                        mou.Vendor = Convert.ToString(ds.Tables[0].Rows[i]["Vendor"]);
                        mou.StartDate = Convert.ToString(ds.Tables[0].Rows[i]["StartDate"]);
                        mou.EndDate = Convert.ToString(ds.Tables[0].Rows[i]["EndDate"]);
                        mou.signedby = Convert.ToString(ds.Tables[0].Rows[i]["signedby"]);
                        mou.FileName = Convert.ToString(ds.Tables[0].Rows[i]["FileName"]);
                        mou.cupboardNumber = Convert.ToString(ds.Tables[0].Rows[i]["cupboardNumber"]);
                        mou.State = Convert.ToString(ds.Tables[0].Rows[i]["State"]);
                        if (Convert.ToInt32(row["ExpiryInDays"]) > 0 && Convert.ToInt32(row["ExpiryInDays"]) < 60)
                            mou.Status = "Expiring in " + Convert.ToInt32(row["ExpiryInDays"]) + " days";
                        else if (Convert.ToInt32(row["ExpiryInDays"]) < 0)
                            mou.Status = "Already Expired " + Convert.ToInt32(Convert.ToString(row["ExpiryInDays"]).Replace("-", "")) + " days before from " + Convert.ToDateTime(DateTime.Now).ToString("dd/MM/yy");


                        list.Add(mou);
                    }
                }
                dt.Columns.Add("Project");
                dt.Columns.Add("Project Incharge");
                dt.Columns.Add("Vendor");
                dt.Columns.Add("Start Date");
                dt.Columns.Add("End Date");
                dt.Columns.Add("Signed by");
                dt.Columns.Add("Cupboard Number");
                dt.Columns.Add("Status");
                dt.Columns.Add("State");

                foreach (var rowObj in list)
                {
                    DataRow row1 = null;
                    row1 = dt.NewRow();
                    dt.Rows.Add(rowObj.Project, rowObj.ProjectIncharge, rowObj.Vendor, rowObj.StartDate,
                       rowObj.EndDate, rowObj.signedby, rowObj.cupboardNumber, rowObj.Status,rowObj.State);
                }

            }
            catch (Exception ex)
            {

                Library.WriteErrorLog(ex.Message);
            }
            return dt;
        }

        private void ExportDataSetToExcel()
        {
            DataTable ds = FilterResult();
            DataTable dtState = ds.DefaultView.ToTable(true, "State");
            string LeadMail = String.Empty;
            string file = String.Empty;
            string AppLocation = String.Empty;

            try
            {
                for (int k = 0; k < dtState.Rows.Count; k++)
                {
                    var filter = ds.AsEnumerable().
                          Where(x => x.Field<string>("State") == Convert.ToString(dtState.Rows[k]["State"]));
                    DataTable row = filter.CopyToDataTable();

                    string state = Convert.ToString(dtState.Rows[k]["State"]);
                    DataSet Dt = GetStateLeadMailIdFromDB(state);
                    LeadMail = Convert.ToString(Dt.Tables[0].Rows[0]["EmailId"]);

                    AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                    AppLocation = AppLocation.Replace("file:\\", "");
                    //string file = AppLocation + "\\ExcelFiles\\MOUList.xlsx";
                    // Bind table data to Stream Writer to export data to respective folder
                    StreamWriter wr = new StreamWriter(@"C:\MOUAlertsWindowsService\ExcelFiles\MOUList.xls");
                    // Write Columns to excel file
                    for (int i = 0; i < row.Columns.Count; i++)
                    {
                        wr.Write(row.Columns[i].ToString().ToUpper() + "\t");
                    }
                    wr.WriteLine();
                    for (int i = 0; i < (row.Rows.Count); i++)
                    {
                        for (int j = 0; j < row.Columns.Count; j++)
                        {
                            if (row.Rows[i][j] != null)
                            {
                                wr.Write(Convert.ToString(row.Rows[i][j]) + "\t");
                            }
                            else
                            {
                                wr.Write("\t");
                            }
                        }
                        wr.WriteLine();
                    }
                    wr.Close();

                    file = AppLocation + "\\ExcelFiles\\ListOfExpiringMOUsOf_" + state + ".xlsx";
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add(row, "ExpiredMou(s)");
                        wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wb.Style.Font.Bold = true;
                        wb.SaveAs(file);
                        SendEmail(LeadMail, file);
                    }
                }
            }
            catch (Exception ex)
            {

                Library.WriteErrorLog(ex.Message);
            }
        }

        public void SendEmail(string LeadMail, string file)
        {
            try
            {
                string MailTo = ConfigurationManager.AppSettings["toAddress"].ToString();
                string MailSubject = ConfigurationManager.AppSettings["mailsubject"].ToString();
                string Password = ConfigurationManager.AppSettings["password"].ToString();
                string mailfrom = ConfigurationManager.AppSettings["fromAddress"].ToString();
                //string AppLocation = "";
                //AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                // AppLocation = AppLocation.Replace("file:\\", "");
                // string file = AppLocation + "\\ExcelFiles\\ListOfExpiringMOUsOf_" + state + ".xlsx";
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["Host"].ToString());
                mail.From = new MailAddress(mailfrom);
                mail.To.Add(LeadMail); // Sending MailTo  
                //List<string> li = new List<string>();
                //li.Add("saihacksoft@gmail.com");
                //li.Add("saihacksoft@gmail.com");
                //li.Add("saihacksoft@gmail.com");
                //li.Add("saihacksoft@gmail.com");
                //li.Add("saihacksoft@gmail.com");
                mail.CC.Add(MailTo); // Sending CC  
                //mail.Bcc.Add(string.Join<string>(",", li)); // Sending Bcc   
                mail.Subject = MailSubject; // Mail Subject  
                mail.Body = string.Format("Dear Team,<br/><br/> Please find attached file as MOU list which are expired or expiring soon.<br/><br/>Thanks <br/>Team InSDMS <br/><br/> *This is an automatically generated email, please do not reply*");
                System.Net.Mail.Attachment attachment;
                mail.IsBodyHtml = true;
                attachment = new System.Net.Mail.Attachment(file); //Attaching File to Mail  
                mail.Attachments.Add(attachment);
                SmtpServer.Port = 587;
                SmtpServer.EnableSsl = true;
                SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.Credentials = new NetworkCredential(mailfrom, Password);
                SmtpServer.Send(mail);
                Library.WriteErrorLog("Mail Send Successfully.");
            }
            catch (Exception ex)
            {
                Library.WriteErrorLog(ex.Message);
            }
        }
        #endregion

        #region ModelClass
        public class MOUList
        {
            public long id { get; set; }
            public string Project { get; set; }
            public string ProjectIncharge { get; set; }
            public string Vendor { get; set; }
            public string StartDate { get; set; }
            public string EndDate { get; set; }
            public string signedby { get; set; }
            public string PaperType { get; set; }
            public string SignedCopySentTo { get; set; }
            public string SPOC { get; set; }
            public string Email { get; set; }
            public string SPOCMobile { get; set; }
            public string ContactNumber { get; set; }
            public string Fax { get; set; }
            public string Address1 { get; set; }
            public string Address2 { get; set; }
            public string City { get; set; }
            public string FilePath { get; set; }
            public string FileName { get; set; }
            public string cupboardNumber { get; set; }
            public string Status { get; set; }
            public string State { get; set; }

        }
        #endregion
    }
}
