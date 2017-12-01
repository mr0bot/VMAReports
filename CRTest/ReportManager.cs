using System;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Net.Mail;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using NLog;
using System.Threading.Tasks;

namespace VMAReports
{
    public class ReportManager
    {
        Logger VMALogger;
        List<string> EmptyList { get; set; }
        bool Testing { get; set; }
        string TestEmailAddress { get; set; }
        public static ReportManager NewReportManager()
        {
            return new ReportManager();
        }

        public ReportManager()
        {
            VMALogger = LogManager.GetCurrentClassLogger();
            EmptyList = new List<string>();
            string [] test = ConfigurationManager.AppSettings["Testing"].Split(',');
            Testing = test.First() == "Y";
            TestEmailAddress = test.Last();
        }

        public void ProcessReports()
        {
            try
            {
                List<CRPT> reports = DAL.GetReportData();

                foreach (CRPT rpt in reports)
                {
                    try
                    {
                        if (Testing) { rpt.Recipients = TestEmailAddress; }
                        ConnectionInfo p_ConnectionInfo = new ConnectionInfo()
                        {
                            DatabaseName = "",
                            ServerName = System.Configuration.ConfigurationManager.AppSettings["DB_INSTANCE"],
                            UserID = System.Configuration.ConfigurationManager.AppSettings["USERNAME"],
                            Password = System.Configuration.ConfigurationManager.AppSettings["PASSWORD"]
                        };

                        ReportDocument p_ReportDocument = new ReportDocument();

                        string v11rptpath = ConfigurationManager.AppSettings["SourcePath"] + rpt.ReportSource.Split('\\').Last().Replace(".RPT", "_v11.RPT");
                        p_ReportDocument.Load(v11rptpath);
                        p_ReportDocument.SetDatabaseLogon(p_ConnectionInfo.UserID, p_ConnectionInfo.Password, p_ConnectionInfo.ServerName, "");

                        //loop through all tables and pass in the connection info
                        foreach (Table p_Table in p_ReportDocument.Database.Tables)
                        {
                            TableLogOnInfo p_TableLogOnInfo = p_Table.LogOnInfo;
                            p_TableLogOnInfo.ConnectionInfo = p_ConnectionInfo;
                            p_Table.ApplyLogOnInfo(p_TableLogOnInfo);
                        }

                        foreach (ReportDocument subReport in p_ReportDocument.Subreports)
                        {
                            foreach (Table p_Table in subReport.Database.Tables)
                            {
                                TableLogOnInfo p_TableLogOnInfo = p_Table.LogOnInfo;
                                p_TableLogOnInfo.ConnectionInfo = p_ConnectionInfo;
                                p_Table.ApplyLogOnInfo(p_TableLogOnInfo);
                            }
                        }

                        rpt.Report = p_ReportDocument;
                        DAL.SetReportParameters(rpt);
                        bool PullData = ExportReport(rpt, p_ReportDocument);
                        p_ReportDocument.Close();
                        p_ReportDocument.Dispose();
                    }
                    catch (Exception e)
                    {
                        VMALogger.Log(LogLevel.Error, Environment.NewLine + "Export: " + rpt.ReportName + " Error Message: " + e.InnerException.Message);
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                VMALogger.Log(LogLevel.Error, Environment.NewLine + "Data Load Error Message: " + ex.Message);
            }
        }

        private bool IsEmptyReport(ReportDocument report)
        {
            try
            {
                return (report.Rows.Count == 0);
            }
            catch (Exception ex)
            {
                VMALogger.Log(LogLevel.Error, Environment.NewLine + "Error on Check for no data: " + ex.Message); ;
            }

            return false;
        }

        private async void SendMail(string attachment, string recipients, string body, string subject)
        {
            try
            {
                using (SmtpClient mail = new SmtpClient())
                {
                    mail.Host = ConfigurationManager.AppSettings["Forwarder"];
                    MailMessage message = new MailMessage("NCL_Reports@ncl.com", recipients);
                    message.Body = body;
                    message.Subject = subject;
                    if (attachment != string.Empty)
                        message.Attachments.Add(new Attachment(attachment));
                    await mail.SendMailAsync(message);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// Method IsEmptyReport checks to see if report pulls data or not. This would tell us weather we needed send an email or not. 
        /// Users requested to send email with attached empty report regardless of the existence of data.
        /// </summary>
        /// <param name="report"></param>
        /// <param name="p_ReportDocument"></param>
        /// <returns></returns>
        private bool ExportReport(CRPT report, ReportDocument p_ReportDocument)
        {
            try
            {
                string outputDir = ConfigurationManager.AppSettings["OutputPath"].ToString() + report.OutputPath + "\\" + DateTime.Now.ToString("yyyy_MM_dd");
                string fname = GetGroupLevelFileName(report); 
                string strReportExportNamePath = outputDir + "\\" + fname;

                bool IsRFPT = SetParameterValues(report, p_ReportDocument);

                if (IsRFPT)
                {
                   // if (IsEmptyReport(p_ReportDocument)) return false;
                    try
                    {
                        if (!Directory.Exists(outputDir))
                            Directory.CreateDirectory(outputDir);

                        if (report.IsPdf())
                            p_ReportDocument.ExportToDisk(ExportFormatType.PortableDocFormat, strReportExportNamePath);
                        else
                            p_ReportDocument.ExportToDisk(ExportFormatType.Excel, strReportExportNamePath);

                        SendMail(strReportExportNamePath, report.Recipients, "Generated Report Attached.", "FreestyleConnect Report");
                    }
                    catch (Exception e)
                    {
                        VMALogger.Log(LogLevel.Error, Environment.NewLine + "Report Name: " + report.ReportName + "Export error message: " + e.Message);
                        return false;
                    }
                }
                else
                {
                    VMALogger.Log(LogLevel.Error, Environment.NewLine + "Report Name: " + report.ReportName + "Export Parameters not found" );
                    return false;
                }
            }
            catch (Exception ex)
            {
                VMALogger.Log(LogLevel.Error, Environment.NewLine + "Set Parameter Error: " + report.ReportName + " Error Message: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool SetParameterValues(CRPT report, ReportDocument p_ReportDocument)
        {
            report.SetParameterValues();

            if (report.MainParmDic.Count + report.SubParmDic.Count == 0) return false;

            SetReportDataValues(p_ReportDocument, report.MainParmDic);

            foreach (ReportDocument sub in p_ReportDocument.Subreports)
                SetReportDataValues(sub, report.SubParmDic);

            return true;
        }

        private void SetReportDataValues(ReportDocument report, Dictionary<string, string> parmdic)
        {

            foreach (KeyValuePair<string, string> kvp in parmdic)
            {

                if (report.IsSubreport)
                    if (IsNumeric(kvp.Value))
                        report.DataDefinition.ParameterFields[kvp.Key].CurrentValues.AddValue(Convert.ToInt64(kvp.Value));
                    else
                        report.DataDefinition.ParameterFields[kvp.Key].CurrentValues.AddValue(kvp.Value);
                else
                {
                    if (IsNumeric(kvp.Value))
                        report.SetParameterValue(kvp.Key, Convert.ToInt64(kvp.Value));
                    else
                        report.SetParameterValue(kvp.Key, kvp.Value);
                }
            }
        }

        private bool IsNumeric(String value)
        {
            int result = 0;
            return int.TryParse(value, out result);
        }

        private string GetGroupLevelFileName(CRPT report)
        {
            if (report.GroupLevel <= 1) return report.Filename;
            string[] fname = report.Filename.Split('.');
            return fname.First() + "_00" + report.GroupLevel.ToString() + "." + fname.Last();
        }
    }
}
