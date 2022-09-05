namespace EODAddIn.Forms
{
    partial class FrmScreener
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmScreener));
            this.dataGridViewFilters = new System.Windows.Forms.DataGridView();
            this.colField = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.colOperation = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.colValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblFilter = new System.Windows.Forms.Label();
            this.chk50d_new_lo = new System.Windows.Forms.CheckBox();
            this.chk50d_new_hi = new System.Windows.Forms.CheckBox();
            this.chk200d_new_lo = new System.Windows.Forms.CheckBox();
            this.chk200d_new_hi = new System.Windows.Forms.CheckBox();
            this.chkBookvalue_neg = new System.Windows.Forms.CheckBox();
            this.chkBookvalue_pos = new System.Windows.Forms.CheckBox();
            this.chkWallstreet_lo = new System.Windows.Forms.CheckBox();
            this.chkWallstreet_hi = new System.Windows.Forms.CheckBox();
            this.cboSortField = new System.Windows.Forms.ComboBox();
            this.lblField = new System.Windows.Forms.Label();
            this.rbtnSortAsc = new System.Windows.Forms.RadioButton();
            this.rbtnSortDesc = new System.Windows.Forms.RadioButton();
            this.lblLimit = new System.Windows.Forms.Label();
            this.numLimit = new System.Windows.Forms.NumericUpDown();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnAddFilter = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewFilters)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLimit)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridViewFilters
            // 
            this.dataGridViewFilters.AllowUserToAddRows = false;
            this.dataGridViewFilters.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewFilters.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridViewFilters.BackgroundColor = System.Drawing.Color.White;
            this.dataGridViewFilters.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewFilters.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colField,
            this.colOperation,
            this.colValue});
            this.dataGridViewFilters.Location = new System.Drawing.Point(12, 28);
            this.dataGridViewFilters.Name = "dataGridViewFilters";
            this.dataGridViewFilters.RowHeadersWidth = 20;
            this.dataGridViewFilters.Size = new System.Drawing.Size(501, 150);
            this.dataGridViewFilters.TabIndex = 0;
            this.dataGridViewFilters.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewFilters_CellValueChanged);
            this.dataGridViewFilters.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dataGridViewFilters_RowsAdded);
            // 
            // colField
            // 
            this.colField.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.colField.HeaderText = "Field";
            this.colField.Name = "colField";
            this.colField.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colField.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // colOperation
            // 
            this.colOperation.HeaderText = "Operation";
            this.colOperation.Name = "colOperation";
            // 
            // colValue
            // 
            this.colValue.HeaderText = "Value";
            this.colValue.Name = "colValue";
            // 
            // lblFilter
            // 
            this.lblFilter.AutoSize = true;
            this.lblFilter.Location = new System.Drawing.Point(12, 9);
            this.lblFilter.Name = "lblFilter";
            this.lblFilter.Size = new System.Drawing.Size(34, 13);
            this.lblFilter.TabIndex = 2;
            this.lblFilter.Text = "Filters";
            // 
            // chk50d_new_lo
            // 
            this.chk50d_new_lo.AutoSize = true;
            this.chk50d_new_lo.Location = new System.Drawing.Point(10, 19);
            this.chk50d_new_lo.Name = "chk50d_new_lo";
            this.chk50d_new_lo.Size = new System.Drawing.Size(84, 17);
            this.chk50d_new_lo.TabIndex = 3;
            this.chk50d_new_lo.Text = "50d_new_lo";
            this.chk50d_new_lo.UseVisualStyleBackColor = true;
            // 
            // chk50d_new_hi
            // 
            this.chk50d_new_hi.AutoSize = true;
            this.chk50d_new_hi.Location = new System.Drawing.Point(128, 19);
            this.chk50d_new_hi.Name = "chk50d_new_hi";
            this.chk50d_new_hi.Size = new System.Drawing.Size(84, 17);
            this.chk50d_new_hi.TabIndex = 4;
            this.chk50d_new_hi.Text = "50d_new_hi";
            this.chk50d_new_hi.UseVisualStyleBackColor = true;
            // 
            // chk200d_new_lo
            // 
            this.chk200d_new_lo.AutoSize = true;
            this.chk200d_new_lo.Location = new System.Drawing.Point(240, 19);
            this.chk200d_new_lo.Name = "chk200d_new_lo";
            this.chk200d_new_lo.Size = new System.Drawing.Size(90, 17);
            this.chk200d_new_lo.TabIndex = 5;
            this.chk200d_new_lo.Text = "200d_new_lo";
            this.chk200d_new_lo.UseVisualStyleBackColor = true;
            // 
            // chk200d_new_hi
            // 
            this.chk200d_new_hi.AutoSize = true;
            this.chk200d_new_hi.Location = new System.Drawing.Point(353, 19);
            this.chk200d_new_hi.Name = "chk200d_new_hi";
            this.chk200d_new_hi.Size = new System.Drawing.Size(90, 17);
            this.chk200d_new_hi.TabIndex = 6;
            this.chk200d_new_hi.Text = "200d_new_hi";
            this.chk200d_new_hi.UseVisualStyleBackColor = true;
            // 
            // chkBookvalue_neg
            // 
            this.chkBookvalue_neg.AutoSize = true;
            this.chkBookvalue_neg.Location = new System.Drawing.Point(10, 42);
            this.chkBookvalue_neg.Name = "chkBookvalue_neg";
            this.chkBookvalue_neg.Size = new System.Drawing.Size(100, 17);
            this.chkBookvalue_neg.TabIndex = 7;
            this.chkBookvalue_neg.Text = "bookvalue_neg";
            this.chkBookvalue_neg.UseVisualStyleBackColor = true;
            // 
            // chkBookvalue_pos
            // 
            this.chkBookvalue_pos.AutoSize = true;
            this.chkBookvalue_pos.Location = new System.Drawing.Point(128, 42);
            this.chkBookvalue_pos.Name = "chkBookvalue_pos";
            this.chkBookvalue_pos.Size = new System.Drawing.Size(99, 17);
            this.chkBookvalue_pos.TabIndex = 8;
            this.chkBookvalue_pos.Text = "bookvalue_pos";
            this.chkBookvalue_pos.UseVisualStyleBackColor = true;
            // 
            // chkWallstreet_lo
            // 
            this.chkWallstreet_lo.AutoSize = true;
            this.chkWallstreet_lo.Location = new System.Drawing.Point(10, 65);
            this.chkWallstreet_lo.Name = "chkWallstreet_lo";
            this.chkWallstreet_lo.Size = new System.Drawing.Size(84, 17);
            this.chkWallstreet_lo.TabIndex = 9;
            this.chkWallstreet_lo.Text = "wallstreet_lo";
            this.chkWallstreet_lo.UseVisualStyleBackColor = true;
            // 
            // chkWallstreet_hi
            // 
            this.chkWallstreet_hi.AutoSize = true;
            this.chkWallstreet_hi.Location = new System.Drawing.Point(128, 65);
            this.chkWallstreet_hi.Name = "chkWallstreet_hi";
            this.chkWallstreet_hi.Size = new System.Drawing.Size(84, 17);
            this.chkWallstreet_hi.TabIndex = 10;
            this.chkWallstreet_hi.Text = "wallstreet_hi";
            this.chkWallstreet_hi.UseVisualStyleBackColor = true;
            // 
            // cboSortField
            // 
            this.cboSortField.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cboSortField.FormattingEnabled = true;
            this.cboSortField.Location = new System.Drawing.Point(69, 306);
            this.cboSortField.Name = "cboSortField";
            this.cboSortField.Size = new System.Drawing.Size(121, 21);
            this.cboSortField.TabIndex = 12;
            // 
            // lblField
            // 
            this.lblField.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblField.AutoSize = true;
            this.lblField.Location = new System.Drawing.Point(15, 310);
            this.lblField.Name = "lblField";
            this.lblField.Size = new System.Drawing.Size(48, 13);
            this.lblField.TabIndex = 13;
            this.lblField.Text = "Sort field";
            // 
            // rbtnSortAsc
            // 
            this.rbtnSortAsc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.rbtnSortAsc.AutoSize = true;
            this.rbtnSortAsc.Checked = true;
            this.rbtnSortAsc.Location = new System.Drawing.Point(205, 307);
            this.rbtnSortAsc.Name = "rbtnSortAsc";
            this.rbtnSortAsc.Size = new System.Drawing.Size(43, 17);
            this.rbtnSortAsc.TabIndex = 14;
            this.rbtnSortAsc.TabStop = true;
            this.rbtnSortAsc.Text = "Asc";
            this.rbtnSortAsc.UseVisualStyleBackColor = true;
            this.rbtnSortAsc.Visible = false;
            // 
            // rbtnSortDesc
            // 
            this.rbtnSortDesc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.rbtnSortDesc.AutoSize = true;
            this.rbtnSortDesc.Location = new System.Drawing.Point(254, 307);
            this.rbtnSortDesc.Name = "rbtnSortDesc";
            this.rbtnSortDesc.Size = new System.Drawing.Size(50, 17);
            this.rbtnSortDesc.TabIndex = 15;
            this.rbtnSortDesc.Text = "Desc";
            this.rbtnSortDesc.UseVisualStyleBackColor = true;
            // 
            // lblLimit
            // 
            this.lblLimit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblLimit.AutoSize = true;
            this.lblLimit.Location = new System.Drawing.Point(344, 309);
            this.lblLimit.Name = "lblLimit";
            this.lblLimit.Size = new System.Drawing.Size(28, 13);
            this.lblLimit.TabIndex = 16;
            this.lblLimit.Text = "Limit";
            // 
            // numLimit
            // 
            this.numLimit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.numLimit.Location = new System.Drawing.Point(390, 307);
            this.numLimit.Name = "numLimit";
            this.numLimit.Size = new System.Drawing.Size(120, 20);
            this.numLimit.TabIndex = 17;
            this.numLimit.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            // 
            // btnOk
            // 
            this.btnOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOk.Location = new System.Drawing.Point(355, 351);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 18;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(436, 351);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 19;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnAddFilter
            // 
            this.btnAddFilter.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAddFilter.Image = ((System.Drawing.Image)(resources.GetObject("btnAddFilter.Image")));
            this.btnAddFilter.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddFilter.Location = new System.Drawing.Point(433, 3);
            this.btnAddFilter.Name = "btnAddFilter";
            this.btnAddFilter.Size = new System.Drawing.Size(80, 23);
            this.btnAddFilter.TabIndex = 20;
            this.btnAddFilter.Text = "      Add filter";
            this.btnAddFilter.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddFilter.UseVisualStyleBackColor = true;
            this.btnAddFilter.Click += new System.EventHandler(this.btnAddFilter_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.chk50d_new_lo);
            this.groupBox1.Controls.Add(this.chk50d_new_hi);
            this.groupBox1.Controls.Add(this.chk200d_new_lo);
            this.groupBox1.Controls.Add(this.chk200d_new_hi);
            this.groupBox1.Controls.Add(this.chkBookvalue_neg);
            this.groupBox1.Controls.Add(this.chkBookvalue_pos);
            this.groupBox1.Controls.Add(this.chkWallstreet_lo);
            this.groupBox1.Controls.Add(this.chkWallstreet_hi);
            this.groupBox1.Location = new System.Drawing.Point(11, 193);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(501, 99);
            this.groupBox1.TabIndex = 21;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Signals";
            // 
            // FrmScreener
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(525, 386);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnAddFilter);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.numLimit);
            this.Controls.Add(this.lblLimit);
            this.Controls.Add(this.rbtnSortDesc);
            this.Controls.Add(this.rbtnSortAsc);
            this.Controls.Add(this.lblField);
            this.Controls.Add(this.cboSortField);
            this.Controls.Add(this.lblFilter);
            this.Controls.Add(this.dataGridViewFilters);
            this.Name = "FrmScreener";
            this.ShowIcon = false;
            this.Text = "New screener";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewFilters)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLimit)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewFilters;
        private System.Windows.Forms.Label lblFilter;
        private System.Windows.Forms.CheckBox chk50d_new_lo;
        private System.Windows.Forms.CheckBox chk50d_new_hi;
        private System.Windows.Forms.CheckBox chk200d_new_lo;
        private System.Windows.Forms.CheckBox chk200d_new_hi;
        private System.Windows.Forms.CheckBox chkBookvalue_neg;
        private System.Windows.Forms.CheckBox chkBookvalue_pos;
        private System.Windows.Forms.CheckBox chkWallstreet_lo;
        private System.Windows.Forms.CheckBox chkWallstreet_hi;
        private System.Windows.Forms.ComboBox cboSortField;
        private System.Windows.Forms.Label lblField;
        private System.Windows.Forms.RadioButton rbtnSortAsc;
        private System.Windows.Forms.RadioButton rbtnSortDesc;
        private System.Windows.Forms.Label lblLimit;
        private System.Windows.Forms.NumericUpDown numLimit;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.DataGridViewComboBoxColumn colField;
        private System.Windows.Forms.DataGridViewComboBoxColumn colOperation;
        private System.Windows.Forms.DataGridViewTextBoxColumn colValue;
        private System.Windows.Forms.Button btnAddFilter;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}