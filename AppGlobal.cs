using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using SAPbobsCOM;


namespace ACM_CSV_Import
{
    class AppGlobal
    {
        public static string AppVersion = "2024.09.24.01";
        public static string SaUser = ConfigurationManager.AppSettings["SaUser"];
        public static string SaPsw = ConfigurationManager.AppSettings["SaPsw"];
        public static Company oCompany;

        public static string SAPServer = ConfigurationManager.AppSettings.Get("SAPServer").ToString();

        public static string SAPDBName = ConfigurationManager.AppSettings["SAPDBName"].ToString();
        public static string SAPUser = ConfigurationManager.AppSettings.Get("SAPUser").ToString();
        public static string SAPUserPsw = ConfigurationManager.AppSettings.Get("SAPUserPsw").ToString();

    }
}
