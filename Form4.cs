using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel; // Add this namespace for ClosedXML
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Windows.Forms.DataVisualization.Charting;
using MySql.Data.MySqlClient;
namespace MySystem
{
    public partial class Form4 : Form
    {
        // Declare MySqlConnection as a member variable
        MySqlConnection conn;
        public Form4()
        {
            InitializeComponent();
            // Initialize MySqlConnection in the constructor
            string myConnectionString = "server=127.0.0.1;uid=root;pwd=09092902988;database=barangay_info_system";
            conn = new MySqlConnection(myConnectionString);
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }
      
        private void Form4_Load(object sender, EventArgs e)
        {

        }
    
        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void label9_Click(object sender, EventArgs e)
        {
            this.Hide();
            var myform = new AboutBox1();
            myform.Show();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {
            this.Hide();
            var myform = new Form1();
            myform.Show();
        }

        private void label6_Click(object sender, EventArgs e)
        {
            this.Hide();
            var myform = new Form2();
            myform.Show();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            this.Hide();
            var myform = new Form2();
            myform.Show();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }



        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
           
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
     
        }


        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void chart1_Click_1(object sender, EventArgs e)
        {

        }

        private void btn1_Click(object sender, EventArgs e)
        {

        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // Open connection
                conn.Open();

                // Query to select all data from the BarangayOfficials table
                string query = "SELECT * FROM BarangayOfficials";

                // Create a DataTable to hold the result of the query
                DataTable dt = new DataTable();

                // Create a MySqlDataAdapter to execute the query and fill the DataTable
                using (MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn))
                {
                    adapter.Fill(dt);
                }

                // Bind the DataTable to the DataGridView
                dgv1.DataSource = dt;
            }
            catch (MySqlException ex)
            {
                // Handle exception
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Close connection
                conn.Close();
            }
        }

        private void dgv1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn1_Click_1(object sender, EventArgs e)
        {
            // Database Connection
            MySqlConnection conn;
            string myConnectionString;
            myConnectionString = "server=127.0.0.1;uid=root;" +
                                   "pwd=09092902988;database=barangay_info_system";

            conn = new MySqlConnection(myConnectionString);

            try
            {
                conn.Open();

                // Construct the SELECT SQL query to fetch data from barangayofficials table
                string selectQuery = "SELECT * FROM barangayofficials";
                MySqlCommand selectCmd = new MySqlCommand(selectQuery, conn);

                DataTable dataTable = new DataTable();

                // Create a MySqlDataAdapter to fill the DataTable
                using (MySqlDataAdapter adapter = new MySqlDataAdapter(selectCmd))
                {
                    // Fill the DataTable with the results from the select query
                    adapter.Fill(dataTable);
                }

                // Bind the DataTable to the dgv1 DataGridView
                dgv1.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error retrieving records: " + ex.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }


        private void btn2_Click(object sender, EventArgs e)
        {
            // Check if there is any data in the DataGridView
            if (dgv1.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Create a new Excel workbook
                var workbook = new XLWorkbook();

                // Sheet 1: Data from DataGridView
                var worksheet1 = workbook.Worksheets.Add("BarangayOfficials_Data");

                // Write column headers to the Excel worksheet
                for (int i = 0; i < dgv1.Columns.Count; i++)
                {
                    worksheet1.Cell(1, i + 1).Value = dgv1.Columns[i].HeaderText;
                }

                // Write data rows to the Excel worksheet
                for (int i = 0; i < dgv1.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv1.Columns.Count; j++)
                    {
                        // Handle null values using ?.ToString() syntax
                        worksheet1.Cell(i + 2, j + 1).Value = dgv1.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // Sheet 2: Add a bar graph
                var worksheet2 = workbook.Worksheets.Add("BarangayOfficials_Graph");

                // Create a bitmap to draw the chart
                Bitmap bmp = new Bitmap(600, 400);

                // Create a chart control
                Chart chart = new Chart();
                chart.Size = new Size(600, 400);

                // Create a chart area and series
                ChartArea chartArea = new ChartArea();
                chartArea.AxisX.MajorGrid.Enabled = false;
                chartArea.AxisY.MajorGrid.Enabled = false;
                chart.ChartAreas.Add(chartArea);

                Series series = new Series();
                series.ChartType = SeriesChartType.Bar;

                // Add data to the series from the DataGridView
                for (int i = 0; i < dgv1.Rows.Count; i++)
                {
                    if (dgv1.Rows[i].Cells[0].Value != null && dgv1.Rows[i].Cells[1].Value != null)
                    {
                        series.Points.AddXY(dgv1.Rows[i].Cells[0].Value.ToString(), dgv1.Rows[i].Cells[1].Value);
                    }
                }

                chart.Series.Add(series);

                // Draw the chart onto the bitmap
                chart.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));

                // Convert the bitmap to a MemoryStream
                MemoryStream stream = new MemoryStream();
                bmp.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                stream.Position = 0;

                // Add the MemoryStream as a picture to the worksheet
                worksheet2.Pictures.Add(stream).MoveTo(worksheet2.Cell("A1"));

                // Save the Excel file to a local path
                string fileName = "BarangayOfficials_Data_With_Graph.xlsx";
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
                workbook.SaveAs(filePath);

                // Open the Excel file
                System.Diagnostics.Process.Start(filePath);

                MessageBox.Show("Excel file with graph created and opened successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating Excel file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn3_Click(object sender, EventArgs e)
        {
            // Database Connection
            MySqlConnection conn;
            string myConnectionString;
            myConnectionString = "server=127.0.0.1;uid=root;" +
                                   "pwd=09092902988;database=barangay_info_system";

            conn = new MySqlConnection(myConnectionString);

            try
            {
                conn.Open();

                // Construct the SELECT SQL query to fetch data from barangayprojects table
                string selectQuery = "SELECT * FROM barangayprojects";
                MySqlCommand selectCmd = new MySqlCommand(selectQuery, conn);

                DataTable dataTable = new DataTable();

                // Create a MySqlDataAdapter to fill the DataTable
                using (MySqlDataAdapter adapter = new MySqlDataAdapter(selectCmd))
                {
                    // Fill the DataTable with the results from the select query
                    adapter.Fill(dataTable);
                }

                // Bind the DataTable to the dgv2 DataGridView
                dgv2.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error retrieving records: " + ex.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        private void btn4_Click(object sender, EventArgs e)
        {
            // Check if there is any data in the DataGridView
            if (dgv2.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Create a new Excel workbook
                var workbook = new XLWorkbook();

                // Sheet 1: Data from DataGridView
                var worksheet1 = workbook.Worksheets.Add("BarangayProjects_Data");

                // Write column headers to the Excel worksheet
                for (int i = 0; i < dgv2.Columns.Count; i++)
                {
                    worksheet1.Cell(1, i + 1).Value = dgv2.Columns[i].HeaderText;
                }

                // Write data rows to the Excel worksheet
                for (int i = 0; i < dgv2.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv2.Columns.Count; j++)
                    {
                        // Handle null values using ?.ToString() syntax
                        worksheet1.Cell(i + 2, j + 1).Value = dgv2.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // Sheet 2: Add a bar graph
                var worksheet2 = workbook.Worksheets.Add("BarangayProjects_Graph");

                // Create a bitmap to draw the chart
                Bitmap bmp = new Bitmap(600, 400);

                // Create a chart control
                Chart chart = new Chart();
                chart.Size = new Size(600, 400);

                // Create a chart area and series
                ChartArea chartArea = new ChartArea();
                chartArea.AxisX.MajorGrid.Enabled = false;
                chartArea.AxisY.MajorGrid.Enabled = false;
                chart.ChartAreas.Add(chartArea);

                Series series = new Series();
                series.ChartType = SeriesChartType.Bar;

                // Add data to the series from the DataGridView
                for (int i = 0; i < dgv2.Rows.Count; i++)
                {
                    if (dgv2.Rows[i].Cells[0].Value != null && dgv2.Rows[i].Cells[1].Value != null)
                    {
                        series.Points.AddXY(dgv2.Rows[i].Cells[0].Value.ToString(), dgv2.Rows[i].Cells[1].Value);
                    }
                }

                chart.Series.Add(series);

                // Draw the chart onto the bitmap
                chart.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));

                // Convert the bitmap to a MemoryStream
                MemoryStream stream = new MemoryStream();
                bmp.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                stream.Position = 0;

                // Add the MemoryStream as a picture to the worksheet
                worksheet2.Pictures.Add(stream).MoveTo(worksheet2.Cell("A1"));

                // Save the Excel file to a local path
                string fileName = "BarangayProjects_Data_With_Graph.xlsx";
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
                workbook.SaveAs(filePath);

                // Open the Excel file
                System.Diagnostics.Process.Start(filePath);

                MessageBox.Show("Excel file with graph created and opened successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating Excel file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn5_Click(object sender, EventArgs e)
        {
            // Database Connection
            MySqlConnection conn;
            string myConnectionString;
            myConnectionString = "server=127.0.0.1;uid=root;" +
                                   "pwd=09092902988;database=barangay_info_system";

            conn = new MySqlConnection(myConnectionString);

            try
            {
                conn.Open();

                // Construct the SELECT SQL query to fetch data from barangayevents table
                string selectQuery = "SELECT * FROM barangayevents";
                MySqlCommand selectCmd = new MySqlCommand(selectQuery, conn);

                DataTable dataTable = new DataTable();

                // Create a MySqlDataAdapter to fill the DataTable
                using (MySqlDataAdapter adapter = new MySqlDataAdapter(selectCmd))
                {
                    // Fill the DataTable with the results from the select query
                    adapter.Fill(dataTable);
                }

                // Bind the DataTable to the dgv3 DataGridView
                dgv3.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error retrieving records: " + ex.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        private void btn6_Click(object sender, EventArgs e)
        {
            // Check if there is any data in the DataGridView
            if (dgv3.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Create a new Excel workbook
                var workbook = new XLWorkbook();

                // Sheet 1: Data from DataGridView
                var worksheet1 = workbook.Worksheets.Add("BarangayEvents_Data");

                // Write column headers to the Excel worksheet
                for (int i = 0; i < dgv3.Columns.Count; i++)
                {
                    worksheet1.Cell(1, i + 1).Value = dgv3.Columns[i].HeaderText;
                }

                // Write data rows to the Excel worksheet
                for (int i = 0; i < dgv3.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv3.Columns.Count; j++)
                    {
                        // Handle null values using ?.ToString() syntax
                        worksheet1.Cell(i + 2, j + 1).Value = dgv3.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // Sheet 2: Add a bar graph
                var worksheet2 = workbook.Worksheets.Add("BarangayEvents_Graph");

                // Create a bitmap to draw the chart
                Bitmap bmp = new Bitmap(600, 400);

                // Create a chart control
                Chart chart = new Chart();
                chart.Size = new Size(600, 400);

                // Create a chart area and series
                ChartArea chartArea = new ChartArea();
                chartArea.AxisX.MajorGrid.Enabled = false;
                chartArea.AxisY.MajorGrid.Enabled = false;
                chart.ChartAreas.Add(chartArea);

                Series series = new Series();
                series.ChartType = SeriesChartType.Bar;

                // Add data to the series from the DataGridView
                for (int i = 0; i < dgv3.Rows.Count; i++)
                {
                    if (dgv3.Rows[i].Cells[0].Value != null && dgv3.Rows[i].Cells[1].Value != null)
                    {
                        series.Points.AddXY(dgv3.Rows[i].Cells[0].Value.ToString(), dgv3.Rows[i].Cells[1].Value);
                    }
                }
                
                chart.Series.Add(series);

                // Draw the chart onto the bitmap
                chart.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));

                // Convert the bitmap to a MemoryStream
                MemoryStream stream = new MemoryStream();
                bmp.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                stream.Position = 0;

                // Add the MemoryStream as a picture to the worksheet
                worksheet2.Pictures.Add(stream).MoveTo(worksheet2.Cell("A1"));

                // Save the Excel file to a local path
                string fileName = "BarangayEvents_Data_With_Graph.xlsx";
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
                workbook.SaveAs(filePath);

                // Open the Excel file
                System.Diagnostics.Process.Start(filePath);

                MessageBox.Show("Excel file with graph created and opened successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating Excel file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
