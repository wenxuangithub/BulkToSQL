using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using ClosedXML.Excel;
using System.Configuration;


namespace SQLBulkCopy
{
    public partial class initialSQLBulkCopy : Form
    {

        public initialSQLBulkCopy()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xlsx";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                var dt = ReadExcelFile(ofd.FileName);
                BulkInsertToSql(dt);
                MessageBox.Show("Import completed.");
            }

        }

        private DataTable ReadExcelFile(string filePath)
        {
            var dt = new DataTable();
            dt.Columns.Add("FirstName");
            dt.Columns.Add("LastName");
            dt.Columns.Add("Email");
            dt.Columns.Add("Age", typeof(int));

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // First sheet
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

                foreach (var row in rows)
                {
                    dt.Rows.Add(
                        row.Cell(1).GetString(), // FirstName
                        row.Cell(2).GetString(), // LastName
                        row.Cell(3).GetString(), // Email
                        row.Cell(4).GetValue<int>() // Age
                    );
                }
            }

            return dt;
        }

        private void BulkInsertToSql(DataTable dataTable)
        {
            string connString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            using (var connection = new SqlConnection(connString))
            {
                connection.Open();
                using (var bulkCopy = new SqlBulkCopy(connection))
                {
                    bulkCopy.DestinationTableName = "DemoUsers";

                    // Map columns
                    bulkCopy.ColumnMappings.Add("FirstName", "FirstName");
                    bulkCopy.ColumnMappings.Add("LastName", "LastName");
                    bulkCopy.ColumnMappings.Add("Email", "Email");
                    bulkCopy.ColumnMappings.Add("Age", "Age");

                    bulkCopy.WriteToServer(dataTable);
                }
            }
        }
    }
}
