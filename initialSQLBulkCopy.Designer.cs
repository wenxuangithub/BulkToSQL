namespace SQLBulkCopy
{
    partial class initialSQLBulkCopy
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnImport = new Button();
            btnLoadExcel = new Button();
            btnGetSqlSchema = new Button();
            lstExcelHeaders = new ListBox();
            lstSqlHeaders = new ListBox();
            lblStatus = new Label();
            cmbSheets = new ComboBox();
            btnColumnMapping = new Button();
            SuspendLayout();
            // 
            // btnImport
            // 
            btnImport.Location = new Point(570, 164);
            btnImport.Name = "btnImport";
            btnImport.Size = new Size(93, 23);
            btnImport.TabIndex = 0;
            btnImport.Text = "Import to SQL";
            btnImport.UseVisualStyleBackColor = true;
            btnImport.Click += btnImport_Click;
            // 
            // btnLoadExcel
            // 
            btnLoadExcel.Location = new Point(146, 236);
            btnLoadExcel.Name = "btnLoadExcel";
            btnLoadExcel.Size = new Size(75, 23);
            btnLoadExcel.TabIndex = 1;
            btnLoadExcel.Text = "Load Excel";
            btnLoadExcel.UseVisualStyleBackColor = true;
            btnLoadExcel.Click += btnLoadExcel_Click;
            // 
            // btnGetSqlSchema
            // 
            btnGetSqlSchema.Location = new Point(394, 236);
            btnGetSqlSchema.Name = "btnGetSqlSchema";
            btnGetSqlSchema.Size = new Size(110, 44);
            btnGetSqlSchema.TabIndex = 2;
            btnGetSqlSchema.Text = "Get SQL Schema";
            btnGetSqlSchema.UseVisualStyleBackColor = true;
            btnGetSqlSchema.Click += btnGetSqlSchema_Click;
            // 
            // lstExcelHeaders
            // 
            lstExcelHeaders.FormattingEnabled = true;
            lstExcelHeaders.ItemHeight = 15;
            lstExcelHeaders.Location = new Point(146, 128);
            lstExcelHeaders.Name = "lstExcelHeaders";
            lstExcelHeaders.Size = new Size(153, 94);
            lstExcelHeaders.TabIndex = 3;
            // 
            // lstSqlHeaders
            // 
            lstSqlHeaders.FormattingEnabled = true;
            lstSqlHeaders.ItemHeight = 15;
            lstSqlHeaders.Location = new Point(386, 128);
            lstSqlHeaders.Name = "lstSqlHeaders";
            lstSqlHeaders.Size = new Size(156, 94);
            lstSqlHeaders.TabIndex = 4;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(59, 79);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(35, 15);
            lblStatus.TabIndex = 5;
            lblStatus.Text = "Label";
            // 
            // cmbSheets
            // 
            cmbSheets.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbSheets.FormattingEnabled = true;
            cmbSheets.Location = new Point(146, 276);
            cmbSheets.Name = "cmbSheets";
            cmbSheets.Size = new Size(121, 23);
            cmbSheets.TabIndex = 6;
            cmbSheets.SelectedIndexChanged += cmbSheets_SelectedIndexChanged;
            // 
            // btnColumnMapping
            // 
            btnColumnMapping.Enabled = false;
            btnColumnMapping.Location = new Point(305, 149);
            btnColumnMapping.Name = "btnColumnMapping";
            btnColumnMapping.Size = new Size(75, 52);
            btnColumnMapping.TabIndex = 7;
            btnColumnMapping.Text = "Column Mapping";
            btnColumnMapping.UseVisualStyleBackColor = true;
            btnColumnMapping.Click += btnColumnMapping_Click;
            // 
            // initialSQLBulkCopy
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(btnColumnMapping);
            Controls.Add(cmbSheets);
            Controls.Add(lblStatus);
            Controls.Add(lstSqlHeaders);
            Controls.Add(lstExcelHeaders);
            Controls.Add(btnGetSqlSchema);
            Controls.Add(btnLoadExcel);
            Controls.Add(btnImport);
            Name = "initialSQLBulkCopy";
            Text = "initialSQLBulkCopy";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnImport;
        private Button btnLoadExcel;
        private Button btnGetSqlSchema;
        private ListBox lstExcelHeaders;
        private ListBox lstSqlHeaders;
        private Label lblStatus;
        private ComboBox cmbSheets;
        private Button btnColumnMapping;
    }
}