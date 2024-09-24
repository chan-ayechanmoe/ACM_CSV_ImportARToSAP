using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.Sql;
using System.Data.SqlTypes;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace ACM_CSV_Import
{
    class CommonFunction
    {
       

        SqlConnection Conn = new SqlConnection(ConfigurationManager.ConnectionStrings["SAPDBConnectionString"].ConnectionString);
        SqlDataAdapter SQLAdapter = new SqlDataAdapter();
        SqlCommand Cmd = new SqlCommand();

        DataTable mDT = new DataTable();
        DataTable mDMSdt = new DataTable();

        
        #region " DB Related "

        public  string ConnectedtoDB()
        {
            string rs = "";
            try
            {
                int err = 0;

                if (AppGlobal.oCompany == null)
                {
                    AppGlobal.oCompany = new SAPbobsCOM.Company();
                }

                if (AppGlobal.oCompany.Connected)
                {
                    if (AppGlobal.oCompany.CompanyDB.ToString() == ConfigurationManager.AppSettings["SAPDBName"])
                    {
                        //already connected to the target company
                        rs = "";
                    }
                    else
                    {
                        AppGlobal.oCompany.Disconnect();
                    }
                }
                switch (ConfigurationManager.AppSettings["SAPServerType"])
                {
                    case "MSSQL2016":
                        AppGlobal.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                        break;
                    case "MSSQL2017":
                        AppGlobal.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
                        break;
                }

                AppGlobal.oCompany.Server = ConfigurationManager.AppSettings["SAPServer"];
                AppGlobal.oCompany.CompanyDB = ConfigurationManager.AppSettings["SAPDBName"];
                AppGlobal.oCompany.UserName = ConfigurationManager.AppSettings["SAPUser"];
                AppGlobal.oCompany.Password = ConfigurationManager.AppSettings["SAPUserPsw"];
                AppGlobal.oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                err = AppGlobal.oCompany.Connect();

                if (err != 0)
                {

                    throw new Exception(AppGlobal.oCompany.GetLastErrorCode() + " | " + AppGlobal.oCompany.GetLastErrorDescription());
                }
                else
                {
                    return "";
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return rs;
        }

        public static void DisConnectToDB()
        {
            try
            {
                if (AppGlobal.oCompany != null)
                {
                    if (AppGlobal.oCompany.Connected)
                    {
                        if (AppGlobal.oCompany.CompanyDB.ToString() == ConfigurationManager.AppSettings["SAPDBName"])
                        {
                            //already connected to the target company
                            AppGlobal.oCompany.Disconnect();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region " Execute Query "

        public DataTable SAP_Local_RunQuery(string querystr)
        {

            mDT = new DataTable();
            try
            {
                if (Conn.State == ConnectionState.Closed)
                {
                    Conn.Open();
                }

                Cmd.CommandType = CommandType.Text;
                Cmd.Connection = Conn;
                Cmd.CommandText = querystr;
                Cmd.CommandTimeout = 0;

                SQLAdapter.SelectCommand = Cmd;


                mDT.Clear();
                SQLAdapter.Fill(mDT);

                CloseConn();


            }
            catch (SqlException sqlEx)
            {
                return new DataTable();
            }
            catch (Exception ex)
            {
                return new DataTable();
            }
            finally
            {
                if ((Conn != null) & Conn == null)
                {
                    Conn.Close();
                }
            }
            return mDT;
        }

        public String SAP_Local_RunQuery_NoResult(string querystr)
        {
            int result = 0;
            try
            {
                if (Conn.State == ConnectionState.Closed)
                {
                    Conn.Open();
                }

                Cmd.CommandType = CommandType.Text;
                Cmd.Connection = Conn;
                Cmd.CommandText = querystr;
                Cmd.CommandTimeout = 0;
                //SQLAdapter.SelectCommand = Cmd;

                result = Convert.ToInt32(Cmd.ExecuteNonQuery());

                return "";

            }
            catch (SqlException sqlEx)
            {

                return sqlEx.Message;
            }
            catch (Exception ex)
            {

                return ex.Message;
            }

            finally
            {
                if ((Conn != null) & Conn == null)
                {
                    Conn.Close();
                }
            }

        }

        public void CloseConn()
        {
            if (Conn.State == ConnectionState.Open)
            {
                Conn.Close();
            }
        }

        #endregion

        

        #region " Write Error File "
        public void WriteErrorToFile(String text, string aFileName)
        {
            try
            {
                string destinationFiles = ConfigurationManager.AppSettings["LogFolder"];

                string Todaysdate = DateTime.Now.ToString("dd-MMM-yyyy");//get now date  
                destinationFiles = Path.Combine(destinationFiles, Todaysdate);
                if (!Directory.Exists(destinationFiles))
                    Directory.CreateDirectory(destinationFiles);//create folder
                String path = Path.Combine(destinationFiles, String.Format("{0}_{1}", aFileName, "ErrorLog.txt"));
                if (!File.Exists(path))
                {
                    using (StreamWriter writer = File.CreateText(path))
                    {
                        //writer.WriteLine(string.Format(text));
                        writer.WriteLine(string.Format("{0}_{1:yyyy-MM-dd HH-mm-ss}", text, DateTime.Now.ToString("hh:mm")));
                        writer.WriteLine();
                        writer.Close();
                    }
                }
                else
                {
                    using (StreamWriter writer = File.AppendText(path))
                    {
                        //writer.WriteLine(string.Format(text));
                        writer.WriteLine(string.Format("{0}_{1:yyyy-MM-dd HH-mm-ss}", text, DateTime.Now.ToString("hh:mm")));
                        writer.WriteLine();
                        writer.Close();
                    }
                }
            }
            catch
            { }
        }
        #endregion

    }
}
