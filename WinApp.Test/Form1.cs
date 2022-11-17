using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Data.OleDb;
using DocumentFormat.OpenXml.Office;
using OfficeOpenXml;
using System.Runtime.InteropServices;

namespace WinApp.Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Excel 97 - 2003|*.xls|Excel 2007|*.xlsx";
            if (op.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (File.Exists(op.FileName))
                {
                    string[] Arr = null;
                    Arr = op.FileName.Split('.');
                    if (Arr.Length > 0)
                    {
                        if (Arr[Arr.Length - 1] == "xls")
                        {
                            sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + op.FileName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
                        }
                        else if (Arr[Arr.Length - 1] == "xlsx")
                        {
                            sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + op.FileName + ";Extended Properties='Excel 12.0 Xml;HDR=YES';";
                        }
                    }
                    foreach (string s in GetExcelSheetNames(op.FileName))
                    {
                        cboSheetName.Items.Add(s);
                        sheetNameText = cboSheetName.Text;
                        //cboSheetName.Text = cboSheetName.Items[1].ToString();
                    }
                    FillData(cboSheetName.Text);
                }
            }
        }
        public string sheetNameText;
        public string sConnectionString;
        private void FillData(string sheetName)
        {
            if (sConnectionString.Length > 0)
            {
                OleDbConnection cn = new OleDbConnection(sConnectionString);
                try
                {
                    cn.Open();
                    DataTable dt = new DataTable();
                    OleDbDataAdapter Adpt = new OleDbDataAdapter("select * from [" + sheetName + "]", cn);
                    Adpt.Fill(dt);
                    DataGridView1.DataSource = dt;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private String[] GetExcelSheetNames(string excelFile)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {
                // Connection String. Change the excel file to the file you will search.
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                    "Data Source=" + excelFile + ";" +
                                    "Extended Properties='Excel 12.0 XML;" +
                                    "HDR=YES;';";

                // Create connection object by using the preceding connection string.
                objConn = new OleDbConnection(sConnectionString);

                // Open connection with the database.
                objConn.Open();

                // Get the data table containg the schema guid.
                //dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    // Query each excel sheet.
                }

                return excelSheets;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        private void loadData_Click(object sender, EventArgs e)
        {
            FillData(sheetNameText);
        }

        private void cboSheetName_SelectedIndexChanged(object sender, EventArgs e)
        {
            sheetNameText = cboSheetName.Text;
        }
    }
}