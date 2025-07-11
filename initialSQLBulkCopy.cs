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
using ExcelDataReader;
using System.Configuration;


namespace SQLBulkCopy
{
    public partial class initialSQLBulkCopy : Form
    {
        private DataTable _excelTable;
        private DataSet _excelDataSet;
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
                foreach (var map in _columnMappings)
                {
                    if (_excelTable.Columns.Contains(map.Key))
                    {
                        _excelTable.Columns[map.Key].ColumnName = map.Value;
                    }
                }

                foreach (var map in _columnMappings)
                {
                    var columnName = map.Value;

                    if (_sqlSchema != null)
                    {
                        var row = _sqlSchema.Rows
                            .Cast<DataRow>()
                            .FirstOrDefault(r => r["ColumnName"].ToString() == columnName);

                        if (row != null && row["DataType"] is Type type)
                        {
                            int maxLength = -1;

                            if (type == typeof(string) && _sqlSchema.Columns.Contains("ColumnSize"))
                            {
                                var sizeObj = row["ColumnSize"];
                                if (sizeObj != DBNull.Value)
                                    maxLength = Convert.ToInt32(sizeObj);
                            }

                            foreach (DataRow dataRow in _excelTable.Rows)
                            {
                                var val = dataRow[columnName]?.ToString().Trim();

                                // Handle BIT / BOOL
                                if (type == typeof(bool))
                                {
                                    if (val?.ToLower() == "true" || val == "1")
                                        dataRow[columnName] = true;
                                    else if (val?.ToLower() == "false" || val == "0")
                                        dataRow[columnName] = false;
                                    else
                                        dataRow[columnName] = DBNull.Value;
                                }

                                // Handle DATETIME
                                else if (type == typeof(DateTime))
                                {
                                    if (DateTime.TryParse(val, out var parsedDate))
                                        dataRow[columnName] = parsedDate;
                                    else
                                        dataRow[columnName] = DBNull.Value;
                                }

                                else if (type == typeof(string))
                                {
                                    if (string.IsNullOrWhiteSpace(val))
                                    {
                                        dataRow[columnName] = DBNull.Value;
                                    }
                                    else
                                    {
                                        val = val.Trim();

                                        // Optional: truncate if longer than column size
                                        if (maxLength > 0 && val.Length > maxLength)
                                            val = val.Substring(0, maxLength);

                                        dataRow[columnName] = val;
                                    }
                                }
                            }
                        }
                    }
                }



                BulkInsertToSql(_excelTable);
                MessageBox.Show("Import completed.");
                lblStatus.Text = "Import completed ✔";
            }
            catch (Exception ex)
            {
                ShowError("Import failed", ex);
                lblStatus.Text = "Import failed ✖";
            }
        }


        private static DataTable ReadExcelFile(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });

            return result.Tables[0];
        }

        private static DataTable ReadExcelSheet(string filePath, string sheetName)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });

            return result.Tables.Cast<DataTable>().FirstOrDefault(t => t.TableName == sheetName)
                   ?? throw new ApplicationException($"Sheet '{sheetName}' not found in the file.");
        }


        private static DataTable GetTableSchema(string tableName)
        {
            string connString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            try
            {
                using var con = new SqlConnection(connString);
                using var cmd = new SqlCommand($"SELECT TOP 0 * FROM {tableName}", con);
                con.Open();
                using var rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
                return rdr.GetSchemaTable();
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Failed to get schema for table: {tableName}", ex);
            }
        }


        private void ShowColumnMappingDialog(List<string> excelHeaders, List<string> sqlHeaders)
        {
            _columnMappings.Clear();

            var mappingForm = new Form { Text = "Map Excel Columns to SQL Columns", Width = 600, Height = 450 };
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false
            };

            var excelCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Excel Header",
                Name = "ExcelHeader",
                ReadOnly = true
            };

            var sqlCol = new DataGridViewComboBoxColumn
            {
                HeaderText = "SQL Column",
                Name = "SqlColumn",
                DataSource = sqlHeaders,
                DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            };

            var ignoreCol = new DataGridViewCheckBoxColumn
            {
                HeaderText = "Ignore",
                Name = "IgnoreColumn"
            };

            dgv.Columns.Add(excelCol);
            dgv.Columns.Add(sqlCol);
            dgv.Columns.Add(ignoreCol);

            foreach (var excel in excelHeaders)
            {
                var matchedSql = sqlHeaders.FirstOrDefault(sqlColName =>
                    string.Equals(sqlColName, excel, StringComparison.OrdinalIgnoreCase));

                var index = dgv.Rows.Add(excel, matchedSql ?? null, false);
            }

            var btnSave = new Button { Text = "Save Mapping", Dock = DockStyle.Bottom };
            btnSave.Click += (s, e) =>
            {
                _columnMappings.Clear();
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    string excelHeader = row.Cells["ExcelHeader"].Value?.ToString();
                    string sqlHeader = row.Cells["SqlColumn"].Value?.ToString();
                    bool ignore = Convert.ToBoolean(row.Cells["IgnoreColumn"].Value ?? false);

                    if (!string.IsNullOrWhiteSpace(excelHeader) &&
                        !string.IsNullOrWhiteSpace(sqlHeader) &&
                        !ignore)
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
                lblStatus.Text = "Mapped headers ready to import.";
            }
            else
            {
                btnImport.Enabled = false;
                lblStatus.Text = "Header mapping incomplete or ignored.";
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
            string tableName = ConfigurationManager.AppSettings["SqlTableName"];

            try
            {
                using var con = new SqlConnection(connString);
                con.Open();

                using var bc = new SqlBulkCopy(con, SqlBulkCopyOptions.TableLock, null)
                {
                    DestinationTableName = tableName,
                    BatchSize = 5000,
                    BulkCopyTimeout = 0
                };

                // Add only mapped columns
                foreach (var map in _columnMappings)
                {
                    if (dt.Columns.Contains(map.Key))
                    {
                        bc.ColumnMappings.Add(map.Key, map.Value);
                    }
                }

                // Clone DataTable to contain only mapped columns
                var filteredTable = dt.DefaultView.ToTable(false, _columnMappings.Keys.ToArray());
                bc.WriteToServer(filteredTable);
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Bulk insert failed", ex);
            }
        }


        private void btnGetSqlSchema_Click(object sender, EventArgs e)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["SqlTableName"];
                _sqlSchema = GetTableSchema(tableName);

                var sqlHeaders = _sqlSchema.Rows
                    .Cast<DataRow>()
                    .Select(r => r["ColumnName"].ToString())
                    .ToList();

                lstSqlHeaders.DataSource = sqlHeaders;

                var excelHeaders = _excelTable.Columns
                    .Cast<DataColumn>()
                    .Select(c => c.ColumnName)
                    .ToList();

                ShowColumnMappingDialog(excelHeaders, sqlHeaders);
            }
            catch (Exception ex)
            {
                ShowError("Failed to retrieve SQL schema", ex);
            }
        }

        private static DataSet LoadExcelDataSet(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            return reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });
        }


        private void btnLoadExcel_Click(object sender, EventArgs e)
        {
            try
            {
                using var ofd = new OpenFileDialog { Filter = "Excel Files|*.xlsx" };
                if (ofd.ShowDialog() != DialogResult.OK) return;

                string filePath = ofd.FileName;
                _excelDataSet = LoadExcelDataSet(filePath);

                var sheetNames = _excelDataSet.Tables.Cast<DataTable>().Select(t => t.TableName).ToList();

                cmbSheets.Items.Clear();
                cmbSheets.Items.AddRange(sheetNames.ToArray());
                cmbSheets.Tag = filePath;

                if (sheetNames.Any())
                {
                    cmbSheets.SelectedIndex = 0;
                }

                lblStatus.Text = "Excel loaded – select a sheet to preview.";
            }
            catch (Exception ex)
            {
                ShowError("Failed to load Excel file", ex);
            }
        }


        private void lblStatus_Click(object sender, EventArgs e)
        {

        }

        private void ShowError(string title, Exception ex)
        {
            MessageBox.Show($"{title}:\n{ex.Message}\n\n{ex.InnerException?.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            lblStatus.Text = $"{title} ✖";
        }

        private void cmbSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string selectedSheet = cmbSheets.SelectedItem?.ToString();

                if (_excelDataSet == null || string.IsNullOrWhiteSpace(selectedSheet))
                    return;

                _excelTable = ReadExcelSheet(cmbSheets.Tag.ToString(), selectedSheet);

                lstExcelHeaders.Items.Clear();
                foreach (DataColumn column in _excelTable.Columns)
                {
                    lstExcelHeaders.Items.Add(column.ColumnName);
                }

                btnGetSqlSchema.Enabled = true;
                lblStatus.Text = $"Preview loaded from sheet: {selectedSheet}";
            }
            catch (Exception ex)
            {
                ShowError("Failed to preview selected sheet", ex);
            }
        }

        private void btnColumnMapping_Click(object sender, EventArgs e)
        {
            try
            {
                if (_excelTable == null || _sqlSchema == null)
                {
                    MessageBox.Show("Load Excel and SQL schema first.", "Missing Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var excelHeaders = _excelTable.Columns
                    .Cast<DataColumn>()
                    .Select(c => c.ColumnName)
                    .ToList();

                var sqlHeaders = _sqlSchema.Rows
                    .Cast<DataRow>()
                    .Select(r => r["ColumnName"].ToString())
                    .ToList();

                ShowColumnMappingDialog(excelHeaders, sqlHeaders);
            }
            catch (Exception ex)
            {
                ShowError("Failed to open mapping dialog", ex);
            }
        }

    }
}