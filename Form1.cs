using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SqlConnection con = new SqlConnection("Data Source=DESKTOP-8J0J8M8\\sqlexpress;Initial Catalog=master;Integrated Security=True");

        OpenFileDialog openFileDialog1 = new OpenFileDialog();

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            DialogResult result = openFileDialog1.ShowDialog();

            try
            {
                // Create OpenFileDialog

                // Set filter for file extension and default file extension
                openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog1.FilterIndex = 1;

                // Display OpenFileDialog by calling ShowDialog method
                // Check if the user selected a file
                if (result == DialogResult.OK)
                {
                    string filename = openFileDialog1.FileName;
                    if (filename == null)
                    {
                        MessageBox.Show("Please select a valid document.");
                    }
                    else
                    {
                        try
                        {
                            // Open connection
                            con.Open();

                            // Excel Application
                            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                            // Open the Workbook
                            Workbook workbook = excelApp.Workbooks.Open(filename);

                            // Select the first worksheet
                            Worksheet worksheet = (Worksheet)workbook.Sheets[1];

                            // Range of used cells
                            Range range = worksheet.UsedRange;

                            // Row count and column count
                            int rowCount = range.Rows.Count;
                            int columnCount = range.Columns.Count;

                            // Loop through the rows in the Excel sheet
                            for (int row = 2; row <= rowCount; row++) // Assuming the first row contains headers
                            {
                                // Insert command with parameters
                                string insertQuery = "INSERT INTO doc (SID, NAME, AGE, LOCATION) VALUES (@SID, @NAME, @AGE, @LOCATION)";
                                SqlCommand cmd = new SqlCommand(insertQuery, con);

                                // Add parameter values from Excel
                                cmd.Parameters.AddWithValue("@SID", ((Range)range.Cells[row, 1]).Value2);
                                cmd.Parameters.AddWithValue("@NAME", ((Range)range.Cells[row, 2]).Value2);
                                cmd.Parameters.AddWithValue("@AGE", ((Range)range.Cells[row, 3]).Value2);
                                cmd.Parameters.AddWithValue("@LOCATION", ((Range)range.Cells[row, 4]).Value2);

                                // Execute the command
                                cmd.ExecuteNonQuery();
                            }

                            MessageBox.Show("Data uploaded successfully.");

                            // Close the workbook and quit Excel
                            workbook.Close();
                            excelApp.Quit();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Data Already uploaded ");
                        }
                        finally
                        {
                            // Close connection
                            con.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {

            try
            {
                // Open the connection to your database
                con.Open();

                // SQL query to select data from the "doc" table
                string query = "SELECT * FROM doc";

                // Create a SqlCommand object
                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    // Create a SqlDataAdapter to fetch data
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);

                    // Create a DataTable to store the data
                    System.Data.DataTable table = new System.Data.DataTable();

                    // Fill the DataTable with data from the adapter
                    adapter.Fill(table);

                    // Display data in the DataGridView
                    dataGridView1.DataSource = table;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Close the connection to your database
                con.Close();
            }

        }
        // Excel Application object
        System.Data.DataTable excelApp = new System.Data.DataTable();
        // List to store sheet names
        List<string> sheetNames = new List<string>();




        private void button4_Click(object sender, EventArgs e)
        {

            try
            {
                // Create OpenFileDialog
                OpenFileDialog openFileDialog = new OpenFileDialog();

                // Set filter for file extension and default file extension
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.FilterIndex = 1;

                // Display OpenFileDialog by calling ShowDialog method
                DialogResult result = openFileDialog.ShowDialog();

                // Check if the user selected a file
                if (result == DialogResult.OK)
                {
                    string filename = openFileDialog.FileName;
                    if (filename == null)
                    {
                        MessageBox.Show("Please select a valid document.");
                    }
                    else
                    {
                        try
                        {
                            // Open connection
                            con.Open();

                            // Excel Application
                            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                            // Open the Workbook
                            Workbook workbook = excelApp.Workbooks.Open(filename);

                            // Select the first worksheet
                            Worksheet worksheet = (Worksheet)workbook.Sheets[1];

                            // Range of used cells
                            Range range = worksheet.UsedRange;

                            // Row count and column count
                            int rowCount = range.Rows.Count;
                            int columnCount = range.Columns.Count;

                            // Loop through the rows in the Excel sheet
                            for (int row = 2; row <= rowCount; row++) // Assuming the first row contains headers
                            {
                                // Insert command with parameters
                                string insertQuery = "INSERT INTO doc (SID, NAME, AGE, LOCATION) VALUES (@SID, @NAME, @AGE, @LOCATION)";
                                SqlCommand cmd = new SqlCommand(insertQuery, con);

                                // Add parameter values from Excel
                                cmd.Parameters.AddWithValue("@SID", ((Range)range.Cells[row, 1]).Value2);
                                cmd.Parameters.AddWithValue("@NAME", ((Range)range.Cells[row, 2]).Value2);
                                cmd.Parameters.AddWithValue("@AGE", ((Range)range.Cells[row, 3]).Value2);
                                cmd.Parameters.AddWithValue("@LOCATION", ((Range)range.Cells[row, 4]).Value2);

                                // Execute the command
                                cmd.ExecuteNonQuery();
                            }

                            MessageBox.Show("Data Updated successfully.");

                            // Close the workbook and quit Excel
                            workbook.Close();
                            excelApp.Quit();
                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("Data Updated uploaded ");
                        }
                        finally
                        {
                            // Close connection
                            con.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                // Open the SqlConnection
                con.Open();

                // SQL query to delete uploaded data
                string query = "DELETE FROM doc";

                // Create a SqlCommand object
                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    // Execute the command
                    int rowsAffected = cmd.ExecuteNonQuery();

                    // Show a message box indicating the number of rows deleted
                    MessageBox.Show(rowsAffected + " row(s) deleted successfully.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Close the SqlConnection
                con.Close();
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Open the database connection
                con.Open();

                // Get the search conditions from the TextBoxes
                string ageCondition = textBox1.Text.Trim();
                string nameCondition = textBox2.Text.Trim();
                string locationCondition = textBox3.Text.Trim();

                // SQL query to select data based on all the conditions
                string query = "SELECT * FROM doc WHERE 1 = 1";

                if (!string.IsNullOrEmpty(ageCondition))
                    query += $" AND AGE = {ageCondition}";

                if (!string.IsNullOrEmpty(nameCondition))
                    query += $" AND NAME LIKE '%{nameCondition}%'";

                if (!string.IsNullOrEmpty(locationCondition))
                    query += $" AND LOCATION LIKE '%{locationCondition}%'";

                // Create a SqlCommand object
                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    // Create a DataTable to store the retrieved data
                    System.Data.DataTable dataTable = new System.Data.DataTable();

                    // Create a SqlDataAdapter to fill the DataTable
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        // Fill the DataTable with the data from the SqlDataAdapter
                        adapter.Fill(dataTable);
                    }

                    // Display the data in the DataGridView
                    if (dataTable.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dataTable;
                    }
                    else
                    {
                        // Display a message indicating no data found
                        MessageBox.Show("No data found for the specified criteria.", "Data Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Close the database connection
                con.Close();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}



    

