using CrystalDecisions.CrystalReports.Engine;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace VMAReports
{
    public class CRPT
    {
        public ReportDocument Report { get; set; }
        public int GroupLevel { get; set; }
        public string ReportName { get; set; }
        public string OutputPath { get; set; }
        public string ReportSource { get; set; }
        public string Filename { get; set; }
        
        public string Recipients { get; set; }
        public Dictionary<string, string> Parameters { get; set; }
        public Dictionary<string, string> MainParmDic { get; set; }
        public Dictionary<string, string> SubParmDic { get; set; }


        public bool IsPdf()
        {
            return (Filename.Split('.').Last().ToUpper() == "PDF");
        }

        public CRPT()
        {
            Parameters = new Dictionary<string, string>();
            MainParmDic = new Dictionary<string, string>();
            SubParmDic = new Dictionary<string, string>();
        }

        public void SetParameterValues()
        {
            foreach (ParameterFieldDefinition parm in Report.DataDefinition.ParameterFields)
            {
                if (parm.ReportName == string.Empty)
                {
                    if (Parameters.ContainsKey(parm.Name.ToUpper()))
                        MainParmDic.Add(parm.Name.ToUpper(), Parameters[parm.Name.ToUpper()]);
                }
                else
                if (Parameters.ContainsKey(parm.Name.ToUpper()))
                    SubParmDic.Add(parm.Name.ToUpper(), Parameters[parm.Name.ToUpper()]);
            }
        }
    }
}

