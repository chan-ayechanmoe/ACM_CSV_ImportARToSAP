using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using SAPbobsCOM;
using System.Configuration;
using System.Globalization;

namespace ACM_CSV_Import
{
    public partial class FrmImport : Form
    {
        CommonFunction func = new CommonFunction();
        string filename = "";

        public FrmImport()
        {
            this.Text = "CSV Import - Version " + AppGlobal.AppVersion;
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ShowFolderBrowser);
            backgroundWorker.RunWorkerAsync();

        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }



        #region " Browse File "
        public void ShowFolderBrowser(object sender, RunWorkerCompletedEventArgs e)
        {
            OpenFileDialog Op = new OpenFileDialog();
            DialogResult result = Op.ShowDialog();
            Op.Title = "Open AR Invoice Import File";
            Op.Filter = "CSV Files (*.csv)|*.csv";
            Op.InitialDirectory = "C:\\";
            Op.FilterIndex = 1;
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                filename = Op.FileName;
                this.txtBrowse.Text = filename;
            }
        }
        #endregion

        private void btnImport_Click(object sender, EventArgs e)
        {
            string SapConn = "";
            try
            {
                int success = 0;
                int skip = 0;

                string filePath = txtBrowse.Text.ToString();
                if (filePath != "")
                {
                    DataTable dt = new DataTable();
                    using (TextFieldParser csvReader = new TextFieldParser(filePath))
                    {
                        csvReader.SetDelimiters(new string[] { "," });
                        csvReader.HasFieldsEnclosedInQuotes = true;
                        //Get Column Names
                        string[] colFields = csvReader.ReadFields();
                        if (colFields != null)
                        {
                            foreach (string _column in colFields)
                            {
                                DataColumn l_DataColumn = new DataColumn(_column);
                                l_DataColumn.AllowDBNull = true;
                                dt.Columns.Add(l_DataColumn);
                            }
                            while (!csvReader.EndOfData)
                            {
                                string[] l_FieldData = csvReader.ReadFields();

                                for (int i = 0; i < l_FieldData.Length; i++)
                                {
                                    if (l_FieldData[i] == "")
                                    {
                                        l_FieldData[i] = null;
                                    }
                                }
                                dt.Rows.Add(l_FieldData);
                            }
                        }

                    }

                    #region "AR Import "

                    #region " Correct excel format "
                    string col1 = "", col2 = "", col3 = "", col4 = "", col5 = "", col6 = "", col7 = "", col8 = "", col9="";
                                        foreach (DataColumn col in dt.Columns)
                    {
                        if (col1 == "")
                            col1 = col.ColumnName;
                        else if (col2 == "")
                            col2 = col.ColumnName;
                        else if (col3 == "")
                            col3 = col.ColumnName;
                        else if (col4 == "")
                            col4 = col.ColumnName;
                        else if (col5 == "")
                            col5 = col.ColumnName;
                        else if (col6 == "")
                            col6 = col.ColumnName;
                        else if (col7 == "")
                            col7 = col.ColumnName;
                        else if (col8 == "")
                            col8 = col.ColumnName;
                        else if (col9 == "")
                            col9 = col.ColumnName;
                    }
                    #endregion

                   
                    try
                    {
                      if (col1.ToUpper().ToString() == "INVOICE NO." && col2.ToUpper().ToString() == "POSTING DATE" && col3.ToUpper().ToString() == "CUSCODE" && col4.ToUpper().ToString() == "CUSNAME" && col5.ToUpper().ToString() == "ITEMCODE"
                               && col6.ToUpper().ToString() == "WHSCODE" && col7.ToUpper().ToString() == "QTY" && col8.ToUpper().ToString() == "UOM" && col9.ToUpper().ToString() == "PRICE") // Correct excel format
                      
                        {
                            if (dt.Rows.Count > 0)
                            {
                                SapConn = func.ConnectedtoDB();

                                if (SapConn == "")
                                {
                                  

                                    SAPbobsCOM.Recordset l_RS = (SAPbobsCOM.Recordset)AppGlobal.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string l_Orderno = "";
                                    SAPbobsCOM.BusinessPartners BP = (SAPbobsCOM.BusinessPartners)AppGlobal.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                                    DataTable l_Tbl = dt.DefaultView.ToTable(true, new string[] { "Invoice No." });
                                    for (int i = 0; i < l_Tbl.Rows.Count; i++)
                                    {
                                         


                                            if (l_Tbl.Rows[i]["Invoice No."].ToString() != "")
                                            {
                                                DataRow[] l_Rows = dt.Select("[Invoice No.]='" + (l_Tbl.Rows[i]["Invoice No."].ToString()) + "'");
                                                SAPbobsCOM.Documents AR = (SAPbobsCOM.Documents)AppGlobal.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                                bool IsFirst = true;



                                                SAPbobsCOM.Recordset l_SlpRS = (SAPbobsCOM.Recordset)AppGlobal.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                                for (int j = 0; j < l_Rows.Count(); j++)
                                                {

                                                    #region " StartTransaction "


                                                    try
                                                    {
                                                        if (AppGlobal.oCompany.InTransaction)
                                                        {
                                                            AppGlobal.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        throw ex;
                                                    }
                                                    finally
                                                    {
                                                        AppGlobal.oCompany.StartTransaction();
                                                    }

                                                    #endregion

                                                    if (IsFirst)
                                                    {
                                                        AR.DocDate = DateTime.ParseExact(l_Rows[j]["Posting Date"].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                                       
                                                        AR.CardName = l_Rows[0]["CusName"].ToString();

                                                        if (l_Rows[0]["CusCode"].ToString() == "")
                                                        {
                                                            AR.CardCode = "C00208900";
                                                        }
                                                        else
                                                        {
                                                            AR.CardCode = l_Rows[0]["CusCode"].ToString();
                                                        }


                                                    }

                                                    if (l_Rows[j]["ItemCode"].ToString() == "")
                                                    {
                                                        AR.Lines.ItemCode = "0100000001";
                                                    }
                                                    else
                                                    {
                                                        AR.Lines.ItemCode = l_Rows[j]["ItemCode"].ToString();
                                                    }
                                                   
                                                    AR.Lines.Quantity = Convert.ToDouble(l_Rows[j]["Qty"]);
                                                    AR.Lines.WarehouseCode = l_Rows[j]["WhsCode"].ToString();
                                                    AR.Lines.UnitPrice = Convert.ToDouble(l_Rows[j]["Price"]);
                                                    AR.Lines.Add();
                                                }

                                                if (AR.Add() != 0)
                                                {
                                                    if(AppGlobal.oCompany.InTransaction)
                                                        AppGlobal.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                                                    skip++;
                                                    int k;
                                                    string errmsg;
                                                    AppGlobal.oCompany.GetLastError(out k, out errmsg);
                                                    func.WriteErrorToFile(errmsg, "Invoice Import Error [" + DateTime.Now.ToString("dd-MM-yyyy") + "]");



                                                }
                                                else
                                                {
                                                    success++;
                                                    if (AppGlobal.oCompany.InTransaction)
                                                        AppGlobal.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                                                }
                                            }
                                        }
                                    AppGlobal.oCompany.Disconnect();
                                    MessageBox.Show("Import Successful! \nTotal Imported Count = " + success + "\nTotal Skip Count = " + skip, "IMPORTING", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                              
                                    }
                                else
                                {
                                    MessageBox.Show("Can't connect to SAP");
                                }


                                   

                                }
                            else
                            {
                                MessageBox.Show("Please choose your file to import!", "IMPORTING", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                            }
									 


                                
                            }
                      else
                        { MessageBox.Show("Invalid Excel Format!", "IMPORTING AR", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                        }

                        }
                    catch (Exception ex)
                    {
                        if (ex.Message != "")
                        {
                            MessageBox.Show(ex.Message, "IMPORTING AR", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                        }
                    }
                       
                    }
                    
                   
                    #endregion

                    #region " Clear File "

                    txtBrowse.Text = "";

                    #endregion
                }
            catch //(Exception ex)
            {
                #region " Clear File "
                txtBrowse.Text = "";
                #endregion

               MessageBox.Show("Import is not Success!", "IMPORTING AR", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
                
            }

            
        }
    

    
}
