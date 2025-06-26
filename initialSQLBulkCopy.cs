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
        private DataTable _excelTable;
        private DataTable _sqlSchema;
        private Dictionary<string, string> _columnMappings = new();

        public initialSQLBulkCopy()
        {
            InitializeComponent();
            btnImport.Enabled = false;
            btnGetSqlSchema.Enabled = false;

        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                // Rename columns in Excel table to match SQL before import
                foreach (var map in _columnMappings)
                {
                    if (_excelTable.Columns.Contains(map.Key))
                    {
                        _excelTable.Columns[map.Key].ColumnName = map.Value;
                    }
                }

                BulkInsertToSql(_excelTable);
                MessageBox.Show("Import completed.");
                lblStatus.Text = "Import completed ✔";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Import failed: {ex.Message}");
                lblStatus.Text = "Import failed ✖";
            }
        }


        private static DataTable ReadExcelFile(string filePath)
        {
            var dt = new DataTable();

            using var workbook = new XLWorkbook(filePath);
            var ws = workbook.Worksheet(1);
            var firstRow = ws.FirstRowUsed();
            var headers = firstRow.Cells().Select(c => c.GetString().Trim()).ToList();

            foreach (var h in headers)
                dt.Columns.Add(h);

            foreach (var row in ws.RowsUsed().Skip(1))
            {
                var newRow = dt.NewRow();
                for (int i = 0; i < headers.Count; i++)
                    newRow[i] = row.Cell(i + 1).GetValue<string>();
                dt.Rows.Add(newRow);
            }
            return dt;
        }

        private static DataTable GetTableSchema(string tableName)
        {
            string connString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            using var con = new SqlConnection(connString);
            using var cmd = new SqlCommand($"SELECT TOP 0 * FROM {tableName}", con);
            con.Open();
            using var rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
            return rdr.GetSchemaTable();
        }

        private void ShowColumnMappingDialog(List<string> excelHeaders, List<string> sqlHeaders)
        {
            var mappingForm = new Form { Text = "Map Excel Columns to SQL Columns", Width = 500, Height = 400 };
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false
            };

            dgv.Columns.Clear();

            var excelCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Excel Header",
                Name = "ExcelHeader",
                DataPropertyName = "ExcelHeader",
                ReadOnly = true
            };

            var sqlCol = new DataGridViewComboBoxColumn
            {
                HeaderText = "SQL Column",
                Name = "SqlColumn",
                DataSource = sqlHeaders,
                DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            };

            dgv.Columns.Add(excelCol);
            dgv.Columns.Add(sqlCol);

            foreach (var excel in excelHeaders)
            {
                dgv.Rows.Add(excel);
            }

            var btnSave = new Button { Text = "Save Mapping", Dock = DockStyle.Bottom };
            btnSave.Click += (s, e) =>
            {
                _columnMappings.Clear();
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.IsNewRow) continue;
                    string excelHeader = row.Cells["ExcelHeader"].Value?.ToString();
                    string sqlHeader = row.Cells["SqlColumn"].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(excelHeader) && !string.IsNullOrWhiteSpace(sqlHeader))
                    {
                        _columnMappings[excelHeader] = sqlHeader;
                    }
                }

                mappingForm.DialogResult = DialogResult.OK;
                mappingForm.Close();
            };

            mappingForm.Controls.Add(dgv);
            mappingForm.Controls.Add(btnSave);

            if (mappingForm.ShowDialog() == DialogResult.OK && ValidateMappedHeadersMatch(sqlHeaders))
            {
                btnImport.Enabled = true;
                lblStatus.Text = "Mapped headers matched – ready to import.";
            }
            else
            {
                btnImport.Enabled = false;
                lblStatus.Text = "Header mapping incomplete.";
            }
        }

        //private static DataTable CreateMappingTable(List<string> excelHeaders)
        //{
        //    var dt = new DataTable();
        //    dt.Columns.Add("ExcelHeader");
        //    dt.Columns.Add("SqlColumn");

        //    foreach (var header in excelHeaders)
        //        dt.Rows.Add(header, "");

        //    return dt;
        //}

        private bool ValidateMappedHeadersMatch(List<string> sqlHeaders)
        {
            var mappedSql = _columnMappings.Values.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var sqlSet = sqlHeaders.ToHashSet(StringComparer.OrdinalIgnoreCase);
            return mappedSql.IsSubsetOf(sqlSet);
        }


        private void BulkInsertToSql(DataTable dt)
        {
            string connString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            using var con = new SqlConnection(connString);
            con.Open();

            using var bc = new SqlBulkCopy(con, SqlBulkCopyOptions.TableLock, null)
            {
                DestinationTableName = "DemoUsers",
                BatchSize = 5000,
                BulkCopyTimeout = 0
            };

            foreach (DataColumn col in dt.Columns)
            {
                bc.ColumnMappings.Add(col.ColumnName, col.ColumnName);
            }

            bc.WriteToServer(dt);
        }

        private void btnGetSqlSchema_Click(object sender, EventArgs e)
        {
            string sqltable = ConfigurationManager.ConnectionStrings["SqlTable"].ConnectionString;
            _sqlSchema = GetTableSchema(sqltable);
            var sqlHeaders = _sqlSchema.Rows.Cast<DataRow>().Select(r => r["ColumnName"].ToString()).ToList();
            lstSqlHeaders.DataSource = sqlHeaders;

            var excelHeaders = _excelTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();

            ShowColumnMappingDialog(excelHeaders, sqlHeaders);
        }

        private void btnLoadExcel_Click(object sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog { Filter = "Excel Files|*.xlsx" };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            _excelTable = ReadExcelFile(ofd.FileName);
            lstExcelHeaders.DataSource = _excelTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();

            btnGetSqlSchema.Enabled = true;
            lblStatus.Text = "Excel loaded – fetch SQL schema next.";
        }

        private void lblStatus_Click(object sender, EventArgs e)
        {

        }

    }
}