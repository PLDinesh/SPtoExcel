using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using XL = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Configuration;
using System.Data.SqlClient;

namespace SPtoExcel
{
    public partial class frmMain : Form
    {
        public SqlConnection DBConnection { get; set; }
        public string ConnectionString { get; set; }
        public SqlCommand DBCommand { get; set; }
        public DataTable SQLResultsTable { get; set; }
        private DataTableCollection CompleteResult { get; set; }

        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {
                EnableDisableSQLTxt(false);
                LoadSettings();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } 

        }

        private void InitializeDBConnection()
        {
            try
            {
                if (DBConnection != null)
                {
                    DBConnection.Dispose();
                    DBConnection = null;
                }

                DBConnection = new SqlConnection();
                ConnectionString = "Server=" + txtServerName.Text + ";Initial Catalog=" +txtDatabaseName.Text + ";";
                if (rbnWindowsAuthentication.Checked)
                {
                    ConnectionString = ConnectionString + "Integrated Security=SSPI";
                }
                else
                {
                    ConnectionString = ConnectionString + "Integrated Security=false;User ID=" +txtSQLUserID.Text+ ";Password=" +txtSQLPassword.Text;
                }

                DBConnection.ConnectionString = ConnectionString;
                DBConnection.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } 
        }
        public void ExportToExcel(DataTableCollection dtList)
        {
            XL.Application oXL;
            XL._Workbook oWB;
            XL._Worksheet oSheet;
            XL.Range oRng;

            try
            {
                oXL = new XL.Application();
                Application.DoEvents();
                oXL.Visible = false;
                //Get a new workbook.
                oWB = (XL._Workbook)(oXL.Workbooks.Add(Missing.Value));
                string[] strWorkSheetNames = textBox1.Text.Split('\n');
                int iCurrentSheet = 0;

                foreach (DataTable dt in dtList)
                {

                    oSheet = (XL._Worksheet)oWB.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                    if(strWorkSheetNames.Length >= iCurrentSheet )
                    {
                        if (strWorkSheetNames[iCurrentSheet] != null && !strWorkSheetNames[iCurrentSheet].Equals(string.Empty))
                            oSheet.Name = strWorkSheetNames[iCurrentSheet].Replace("\r", "").Replace(" ",string.Empty);
                    }
                    //System.Data.DataTable dtGridData=ds.Tables[0];
                    //System.Data.DataTable dtGridData=ds.Tables[0];

                    // Copy the DataTable to an object array
                    object[,] rawData = new object[dt.Rows.Count + 1, dt.Columns.Count];

                    // Copy the column names to the first row of the object array
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        rawData[0, col] = dt.Columns[col].ColumnName;
                    }

                    // Copy the values to the object array
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        for (int row = 0; row < dt.Rows.Count; row++)
                        {
                            rawData[row + 1, col] = dt.Rows[row].ItemArray[col].ToString();
                        }
                    }

                    // Calculate the final column letter
                    string finalColLetter = string.Empty;
                    string colCharset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                    int colCharsetLen = colCharset.Length;

                    if (dt.Columns.Count > colCharsetLen)
                    {
                        finalColLetter = colCharset.Substring(
                            (dt.Columns.Count - 1) / colCharsetLen - 1, 1);
                    }

                    finalColLetter += colCharset.Substring(
                            (dt.Columns.Count - 1) % colCharsetLen, 1);

                    // Fast data export to Excel
                    string excelRange = string.Format("A1:{0}{1}",
                        finalColLetter, dt.Rows.Count + 1);
                    
                    //oSheet.get_Range(excelRange, Type.Missing).Style.NumberFormat = "@";
                    oSheet.get_Range(excelRange, Type.Missing).Value2 = rawData;
                    XL.Range oRange = oSheet.get_Range(excelRange, Type.Missing);
                    XL.Style oStyle = (XL.Style)oRange.Style;
                    oStyle.NumberFormat = "@";


                    /*
                    int iRow = 2;
                    if (dt.Rows.Count > 0)
                    {

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            oSheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                        }
                        // For each row, print the values of each column.
                        for (int rowNo = 0; rowNo < dt.Rows.Count; rowNo++)
                        {
                            for (int colNo = 0; colNo < dt.Columns.Count; colNo++)
                            {
                                oSheet.Cells[iRow, colNo + 1] = dt.Rows[rowNo][colNo].ToString();
                            }
                            iRow++;
                        }
                        iRow++;
                    }
                    */
                    XL.Range oRng2 = oSheet.get_Range("A1", "IV1");
                    oRng2.EntireColumn.AutoFit();

                    /*
                    int iRow = 2;
                    if (dt.Rows.Count > 0)
                    {

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            oSheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                        }
                        // For each row, print the values of each column.
                        for (int rowNo = 0; rowNo < dt.Rows.Count; rowNo++)
                        {
                            for (int colNo = 0; colNo < dt.Columns.Count; colNo++)
                            {
                                oSheet.Cells[iRow, colNo + 1] = dt.Rows[rowNo][colNo].ToString();
                            }
                            iRow++;
                        }
                        iRow++;
                    }
                    oRng = oSheet.get_Range("A1", "IV1");
                    oRng.EntireColumn.AutoFit();*/
                    iCurrentSheet++;
                }
                oXL.Visible = true;
            }
            catch (Exception theException)
            {
                throw theException;
            }
            finally
            {
                oXL = null;
                oWB = null;
                oSheet = null;
                oRng = null;
            }

        }
        public void ExportToExcel(DataTable[] dtList)
        {
            XL.Application oXL;
            XL._Workbook oWB;
            XL._Worksheet oSheet;
            XL.Range oRng;

            try
            {
                oXL = new XL.Application();
                Application.DoEvents();
                oXL.Visible = false;
                //Get a new workbook.
                oWB = (XL._Workbook)(oXL.Workbooks.Add(Missing.Value));

                foreach (DataTable dt in dtList)
                {

                    oSheet = (XL._Worksheet)oWB.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                    //System.Data.DataTable dtGridData=ds.Tables[0];

                    //System.Data.DataTable dtGridData=ds.Tables[0];

                    // Copy the DataTable to an object array
                    object[,] rawData = new object[dt.Rows.Count + 1, dt.Columns.Count];

                    // Copy the column names to the first row of the object array
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        rawData[0, col] = dt.Columns[col].ColumnName;
                    }

                    // Copy the values to the object array
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        for (int row = 0; row < dt.Rows.Count; row++)
                        {
                            rawData[row + 1, col] = dt.Rows[row].ItemArray[col].ToString();
                        }
                    }

                    // Calculate the final column letter
                    string finalColLetter = string.Empty;
                    string colCharset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                    int colCharsetLen = colCharset.Length;

                    if (dt.Columns.Count > colCharsetLen)
                    {
                        finalColLetter = colCharset.Substring(
                            (dt.Columns.Count - 1) / colCharsetLen - 1, 1);
                    }

                    finalColLetter += colCharset.Substring(
                            (dt.Columns.Count - 1) % colCharsetLen, 1);

                    // Fast data export to Excel
                    string excelRange = string.Format("A1:{0}{1}",
                        finalColLetter, dt.Rows.Count + 1);

                    oSheet.get_Range(excelRange, Type.Missing).Value2 = rawData;
                    /*
                    int iRow = 2;
                    if (dt.Rows.Count > 0)
                    {

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            oSheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                        }
                        // For each row, print the values of each column.
                        for (int rowNo = 0; rowNo < dt.Rows.Count; rowNo++)
                        {
                            for (int colNo = 0; colNo < dt.Columns.Count; colNo++)
                            {
                                oSheet.Cells[iRow, colNo + 1] = dt.Rows[rowNo][colNo].ToString();
                            }
                            iRow++;
                        }
                        iRow++;
                    }
                     * */
                    oRng = oSheet.get_Range("A1", "IV1");
                    oRng.EntireColumn.AutoFit();
                    /*
                    int iRow = 2;
                    if (dt.Rows.Count > 0)
                    {

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            oSheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                        }
                        // For each row, print the values of each column.
                        for (int rowNo = 0; rowNo < dt.Rows.Count; rowNo++)
                        {
                            for (int colNo = 0; colNo < dt.Columns.Count; colNo++)
                            {
                                oSheet.Cells[iRow, colNo + 1] = dt.Rows[rowNo][colNo].ToString();
                            }
                            iRow++;
                        }
                        iRow++;
                    }
                    oRng = oSheet.get_Range("A1", "IV1");
                    oRng.EntireColumn.AutoFit();*/
                }
                oXL.Visible = true;
            }
            catch (Exception theException)
            {
                throw theException;
            }
            finally
            {
                oXL = null;
                oWB = null;
                oSheet = null;
                oRng = null;
            }

        }
        public void ExportToExcel(DataTable dt)
        {
            XL.Application oXL;
            XL._Workbook oWB;
            XL._Worksheet oSheet;
            XL.Range oRng;

            try
            {
                oXL = new XL.Application();
                Application.DoEvents();
                oXL.Visible = false;
                //Get a new workbook.
                oWB = (XL._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (XL._Worksheet)oWB.ActiveSheet;
                //System.Data.DataTable dtGridData=ds.Tables[0];

                // Copy the DataTable to an object array
                object[,] rawData = new object[dt.Rows.Count + 1, dt.Columns.Count];

                // Copy the column names to the first row of the object array
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    rawData[0, col] = dt.Columns[col].ColumnName;
                }

                // Copy the values to the object array
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        rawData[row + 1, col] = dt.Rows[row].ItemArray[col].ToString();
                    }
                }

                // Calculate the final column letter
                string finalColLetter = string.Empty;
                string colCharset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                int colCharsetLen = colCharset.Length;

                if (dt.Columns.Count > colCharsetLen)
                {
                    finalColLetter = colCharset.Substring(
                        (dt.Columns.Count - 1) / colCharsetLen - 1, 1);
                }

                finalColLetter += colCharset.Substring(
                        (dt.Columns.Count - 1) % colCharsetLen, 1);

                // Fast data export to Excel
                string excelRange = string.Format("A1:{0}{1}",
                    finalColLetter, dt.Rows.Count + 1);

                oSheet.get_Range(excelRange, Type.Missing).Value2 = rawData;

                XL.Range oRange = oSheet.get_Range(excelRange, Type.Missing);
                XL.Style oStyle = (XL.Style)oRange.Style;
               // oStyle.NumberFormat = "@";
                /*
                int iRow = 2;
                if (dt.Rows.Count > 0)
                {

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        oSheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                    }
                    // For each row, print the values of each column.
                    for (int rowNo = 0; rowNo < dt.Rows.Count; rowNo++)
                    {
                        for (int colNo = 0; colNo < dt.Columns.Count; colNo++)
                        {
                            oSheet.Cells[iRow, colNo + 1] = dt.Rows[rowNo][colNo].ToString();
                        }
                        iRow++;
                    }
                    iRow++;
                }
                 * */
                oRng = oSheet.get_Range("A1", "IV1");
                oRng.EntireColumn.AutoFit();
                oXL.Visible = true;
            }
            catch (Exception theException)
            {
                throw theException;
            }
            finally
            {
                oXL = null;
                oWB = null;
                oSheet = null;
                oRng = null;
            }

        }


        /*Import from Excel to datatable */

        public DataTable ImportFromExcel(string strPath)
        {
            
                DataTable dtTable = new DataTable();
                DataColumn col = new DataColumn("Rfid");
                dtTable.Columns.Add(col);
                DataRow drRow;

                    Microsoft.Office.Interop.Excel.Application ExcelObj =
                        new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook theWorkbook =
                            ExcelObj.Workbooks.Open(strPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;
                    try
                    {
                        for (int sht = 1; sht <= sheets.Count; sht++)
                       {
                            Microsoft.Office.Interop.Excel.Worksheet worksheet =
                                    (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(sht);

                            for (int i = 2; i <= worksheet.UsedRange.Rows.Count; i++)
                            {
                                Microsoft.Office.Interop.Excel.Range range = worksheet.get_Range("A" + i.ToString(), "B" + i.ToString());
                                System.Array myvalues = (System.Array)range.Cells.Value2;
                                String name = Convert.ToString(myvalues.GetValue(1, 1));
                                if (string.IsNullOrEmpty(name) == false)
                                {
                                    drRow = dtTable.NewRow();
                                    drRow["Rfid"] = name;
                                    dtTable.Rows.Add(drRow);
                                }
                            }
                            Marshal.ReleaseComObject(worksheet);
                            worksheet = null;
                        }
                    return dtTable;

                }
                catch
                {
                    throw;
                }
                finally
                {
                   // Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(sheets);
                    Marshal.ReleaseComObject(theWorkbook);
                    Marshal.ReleaseComObject(ExcelObj);
                    //worksheet = null;
                    sheets = null;
                    theWorkbook = null;
                    ExcelObj = null;
                }

        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            //Save the settings here..
            SaveSettings();
          
        }

        private void rbnWindowsAuthentication_CheckedChanged(object sender, EventArgs e)
        {
            //Disable the controls
            EnableDisableSQLTxt(false);
        }

        private void rbnSqlAuthentication_CheckedChanged(object sender, EventArgs e)
        {
            //Enable SQL user id and password
            EnableDisableSQLTxt(true);
        }

        private void EnableDisableSQLTxt(bool isEnable)
        {
            try
            {
                if (isEnable)
                {
                    txtSQLPassword.Enabled = true;
                    txtSQLUserID.Enabled = true;
                }
                else
                {
                    txtSQLPassword.Enabled = false;
                    txtSQLUserID.Enabled = false;

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void SaveSettings()
        {
            try
            {
                Properties.Settings MySettings = new SPtoExcel.Properties.Settings();

                MySettings.Server = txtServerName.Text;
                MySettings.Database = txtDatabaseName.Text;
                MySettings.IntegratedSecurity = rbnWindowsAuthentication.Checked;
                MySettings.SQLUserID = txtSQLUserID.Text;
                MySettings.SQLPassword = txtSQLPassword.Text;
                MySettings.SQLQuery = txtSQLQuery.Text;
                MySettings.SheetNames = textBox1.Text;
                MySettings.Save();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadSettings()
        {
            //Load the settings here..
            try
            {
                Properties.Settings MySettings = new SPtoExcel.Properties.Settings();

                MySettings.Reload();
                txtServerName.Text = MySettings.Server;
                txtDatabaseName.Text= MySettings.Database;
                rbnWindowsAuthentication.Checked=MySettings.IntegratedSecurity;
                txtSQLUserID.Text=MySettings.SQLUserID;
                txtSQLPassword.Text=MySettings.SQLPassword;
                txtSQLQuery.Text=MySettings.SQLQuery;
                textBox1.Text = MySettings.SheetNames;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {

                InitializeDBConnection();
                MessageBox.Show("Connection Test success!", "Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                ExecuteQuery();
                grdResultsGrid.DataSource = SQLResultsTable;
                grdResultsGrid.Refresh();
                MessageBox.Show("Query Execution completed!", "Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExecuteQuery()
        {
            if (DBConnection == null || DBConnection.State != ConnectionState.Open)
            {
                InitializeDBConnection();
            }
            DBCommand = new SqlCommand();
            DBCommand.CommandText = txtSQLQuery.Text;
            DBCommand.Connection = DBConnection;

            SqlDataAdapter oAdap = new SqlDataAdapter(DBCommand);
            DataSet oDS = new DataSet();
            oAdap.Fill(oDS);
            SQLResultsTable = oDS.Tables[0];
        }

        private void ExecuteMultiResultSetQuery()
        {
            if (DBConnection == null || DBConnection.State != ConnectionState.Open)
            {
                InitializeDBConnection();
            }
            DBCommand = new SqlCommand();
            DBCommand.CommandText = txtSQLQuery.Text;
            DBCommand.Connection = DBConnection;

            SqlDataAdapter oAdap = new SqlDataAdapter(DBCommand);
            DataSet oDS = new DataSet();
            oAdap.Fill(oDS);
            SQLResultsTable = oDS.Tables[0];
            CompleteResult = oDS.Tables;
        }

        private void btnSaveExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (SQLResultsTable == null)
                {
                    ExecuteQuery();
                }
                ExportToExcel(SQLResultsTable);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                ExecuteMultiResultSetQuery();

                ExportToExcel(CompleteResult);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

 }

