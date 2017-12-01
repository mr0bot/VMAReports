using System;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Net.Mail;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace VMAReports
{
    public class Program
    {

        static void Main(string[] args)
        {
            ReportManager rm = ReportManager.NewReportManager();
            rm.ProcessReports();
        }

    }

}