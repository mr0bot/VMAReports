using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VMAReports
{
    public class DAL
    {
        private static string GetConnectionString()
        {
            return "Data Source = (DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = " + ConfigurationManager.AppSettings["HOST"] + ")(PORT = " + ConfigurationManager.AppSettings["PORT"] 
                + ")))(CONNECT_DATA = (SERVICE_NAME = " + ConfigurationManager.AppSettings["DB_INSTANCE"] 
                + "))); User ID = " + ConfigurationManager.AppSettings["USERNAME"] 
                + "; Password = " + ConfigurationManager.AppSettings["PASSWORD"] + ";";
        }


        public static List<CRPT> GetReportData()
        {
            DataTable dt = new DataTable();
            List<CRPT> rptlist = new List<CRPT>();
            //string sql = SqlString();

            try
            {
                using (OracleConnection oc = new OracleConnection(GetConnectionString()))
                {
                    using (OracleCommand cmd = new OracleCommand("NCLSEA.NCL_VMA_RPT_PKG.GET_RPT_ROUTING_INFO", oc))
                    {
                        oc.Open();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("OUT_CUR", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                        OracleDataAdapter adapter = new OracleDataAdapter(cmd);
                        adapter.Fill(dt);
                    }

                    foreach (DataRow row in dt.Rows)
                    {
                        CRPT report = new CRPT()
                        {
                            ReportName = row["RPTNAME"].ToString(),
                            GroupLevel = Convert.ToInt32(row["GROUP_LEVEL"]),
                            OutputPath = row["OUTPATH"].ToString(),
                            ReportSource = row["RPT_SOURCE"].ToString(),
                            Filename = row["FILENAME"].ToString(),
                            Recipients = row["RECIPIENTS"].ToString().Replace(";", ",")
                        };
                        rptlist.Add(report);
                    }
                }
                return rptlist;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void SetReportParameters(CRPT rpt)
        {
            DataTable dt = new DataTable();

            try
            {
                using (OracleConnection oc = new OracleConnection(GetConnectionString()))
                {
                    using (OracleCommand cmd = new OracleCommand("NCLSEA.NCL_VMA_RPT_PKG.GET_RPT_PARM_VALUES", oc))
                    {
                        oc.Open();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("OUT_CUR", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                        cmd.Parameters.Add("IN_RPT_NAME", OracleDbType.Varchar2).Direction = ParameterDirection.Input;
                        cmd.Parameters["IN_RPT_NAME"].Value = rpt.ReportName;
                        cmd.Parameters.Add("IN_GROUP_LEVEL", OracleDbType.Int32).Direction = ParameterDirection.Input;
                        cmd.Parameters["IN_GROUP_LEVEL"].Value = rpt.GroupLevel;
                        OracleDataAdapter adapter = new OracleDataAdapter(cmd);
                        adapter.Fill(dt);
                    }

                    foreach (DataRow row in dt.Rows)
                    {
                        rpt.Parameters.Add(row["PARAM_NAME"].ToString(), row["PARAM_VALUE"].ToString());
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
