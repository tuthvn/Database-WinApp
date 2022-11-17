using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;

namespace WinApp.Data
{
    public partial class ToolApp : Form
    {
        public ToolApp()
        {
            InitializeComponent();
        }
        public string sheetNameText;
        public string sConnectionString;
        private void loadButton_Click(object sender, EventArgs e)
        {
            int size = -1;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                try
                {
                    string text = File.ReadAllText(file);
                    size = text.Length;
                }
                catch (IOException)
                {
                }
                linkFile.Text = file;

                foreach (string s in GetExcelSheetNames(file))
                {
                    comboBoxSheet.Items.Add(s);
                }
                if (comboBoxSheet.Items[0].ToString() != null)
                {
                    comboBoxSheet.Text = comboBoxSheet.Items[0].ToString();
                    sheetNameText = comboBoxSheet.Text;
                    FillData(sheetNameText);
                }
               
            }
            
        }
        private String[] GetExcelSheetNames(string excelFile)
        {
            OleDbConnection objConn = null!;
            System.Data.DataTable dt = null!;

            try
            {
                // Connection String. Change the excel file to the file you will search.
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                    "Data Source=" + excelFile + ";" +
                                    "Extended Properties='Excel 12.0 XML;" +
                                    "HDR=YES;';";
                sConnectionString = connString;
                // Create connection object by using the preceding connection string.
                objConn = new OleDbConnection(connString);

                // Open connection with the database.
                objConn.Open();

                // Get the data table containg the schema guid.
                //dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null!;
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
                MessageBox.Show(ex.Message);
                return null!;
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
                    dataGridView.DataSource = dt;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void comboBoxSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillData(comboBoxSheet.Text);
        }

        private void copyButton_Click(object sender, EventArgs e)
        {
            string query = "SELECT * FROM ["+ comboBoxSheet.Text +"]";
            using (OleDbConnection connection = new(sConnectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);
                connection.Open();

               

                DataTable dt = new DataTable();
                OleDbDataAdapter Adpt = new OleDbDataAdapter("select * from [" + comboBoxSheet.Text + "]", connection);
                Adpt.Fill(dt);
                dataGridView.DataSource = dt;

                var sqlQuery = "CREATE TABLE " + dataGridView.Columns[0].HeaderText.ToString() + "(\n";

                for (int i = 1; i < dataGridView.RowCount; i++)
                {
                    for (int j = 1; j < dataGridView.ColumnCount; j++)
                    {
                        if (dataGridView.Rows[i].Cells[j].Value.ToString() != null || dataGridView.Rows[i].Cells[j].Value.ToString() != String.Empty)
                        {
                            if (j == 1)
                            {
                                sqlQuery += (dataGridView.Rows[i].Cells[j].Value.ToString() ?? "") + " ";
                            }
                            if (j == 2)
                            {
                                sqlQuery += (dataGridView.Rows[i].Cells[j].Value.ToString() ?? "") + " ";
                            }
                            if (j == 4)
                            {
                                sqlQuery += "IDENTITY(1,1) PRIMARY KEY,\n ";
                            }
                        }
                    }
                }

                OleDbDataReader reader = command.ExecuteReader();
                var lines = new List<string>();
                while (reader.Read())
                {
                    var fieldCount = reader.FieldCount;
                    var fieldIncrementor = 1;
                    var fields = new List<string>();
                    while (fieldCount >= fieldIncrementor)
                    {
                        fields.Add(reader[fieldIncrementor - 1].ToString());
                        fieldIncrementor++;
                    }

                    lines.Add(string.Join("\t", fields));
                }
                
                
                reader.Close();
            }
        }
    }
}