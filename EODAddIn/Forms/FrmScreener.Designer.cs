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
            this.label1 = new System.Windows.Forms.Label();
            this.txtCode = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtExchange = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cboSector = new System.Windows.Forms.ComboBox();
            this.cboIndustry = new System.Windows.Forms.ComboBox();
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
            this.dataGridViewFilters.Location = new System.Drawing.Point(12, 99);
            this.dataGridViewFilters.Name = "dataGridViewFilters";
            this.dataGridViewFilters.RowHeadersWidth = 20;
            this.dataGridViewFilters.Size = new System.Drawing.Size(581, 226);
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
            this.cboSortField.Location = new System.Drawing.Point(69, 453);
            this.cboSortField.Name = "cboSortField";
            this.cboSortField.Size = new System.Drawing.Size(121, 21);
            this.cboSortField.TabIndex = 12;
            // 
            // lblField
            // 
            this.lblField.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblField.AutoSize = true;
            this.lblField.Location = new System.Drawing.Point(15, 457);
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
            this.rbtnSortAsc.Location = new System.Drawing.Point(214, 454);
            this.rbtnSortAsc.Name = "rbtnSortAsc";
            this.rbtnSortAsc.Size = new System.Drawing.Size(43, 17);
            this.rbtnSortAsc.TabIndex = 14;
            this.rbtnSortAsc.TabStop = true;
            this.rbtnSortAsc.Text = "Asc";
            this.rbtnSortAsc.UseVisualStyleBackColor = true;
            // 
            // rbtnSortDesc
            // 
            this.rbtnSortDesc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.rbtnSortDesc.AutoSize = true;
            this.rbtnSortDesc.Location = new System.Drawing.Point(263, 454);
            this.rbtnSortDesc.Name = "rbtnSortDesc";
            this.rbtnSortDesc.Size = new System.Drawing.Size(50, 17);
            this.rbtnSortDesc.TabIndex = 15;
            this.rbtnSortDesc.Text = "Desc";
            this.rbtnSortDesc.UseVisualStyleBackColor = true;
            // 
            // lblLimit
            // 
            this.lblLimit.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.lblLimit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblLimit.AutoSize = true;
            this.lblLimit.Location = new System.Drawing.Point(344, 456);
            this.lblLimit.Name = "lblLimit";
            this.lblLimit.Size = new System.Drawing.Size(28, 13);
            this.lblLimit.TabIndex = 16;
            this.lblLimit.Text = "Limit";
            // 
            // numLimit
            // 
            this.numLimit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.numLimit.Location = new System.Drawing.Point(390, 454);
            this.numLimit.Maximum = new decimal(new int[] {
            2000,
            0,
            0,
            0});
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
            this.btnOk.Location = new System.Drawing.Point(435, 498);
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
            this.btnCancel.Location = new System.Drawing.Point(516, 498);
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
            this.btnAddFilter.Location = new System.Drawing.Point(466, 70);
            this.btnAddFilter.Name = "btnAddFilter";
            this.btnAddFilter.Size = new System.Drawing.Size(127, 23);
            this.btnAddFilter.TabIndex = 20;
            this.btnAddFilter.Text = "      Add number filter";
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
            this.groupBox1.Location = new System.Drawing.Point(11, 340);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(581, 99);
            this.groupBox1.TabIndex = 21;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Signals";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 13);
            this.label1.TabIndex = 22;
            this.label1.Text = "Code";
            // 
            // txtCode
            // 
            this.txtCode.Location = new System.Drawing.Point(90, 11);
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(100, 20);
            this.txtCode.TabIndex = 23;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 24;
            this.label2.Text = "Name";
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(90, 37);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(100, 20);
            this.txtName.TabIndex = 25;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 66);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 13);
            this.label3.TabIndex = 26;
            this.label3.Text = "Exchange";
            // 
            // txtExchange
            // 
            this.txtExchange.Location = new System.Drawing.Point(90, 63);
            this.txtExchange.Name = "txtExchange";
            this.txtExchange.Size = new System.Drawing.Size(100, 20);
            this.txtExchange.TabIndex = 27;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(223, 14);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(38, 13);
            this.label4.TabIndex = 28;
            this.label4.Text = "Sector";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(224, 40);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(44, 13);
            this.label5.TabIndex = 29;
            this.label5.Text = "Industry";
            // 
            // cboSector
            // 
            this.cboSector.FormattingEnabled = true;
            this.cboSector.Location = new System.Drawing.Point(289, 10);
            this.cboSector.Name = "cboSector";
            this.cboSector.Size = new System.Drawing.Size(165, 21);
            this.cboSector.TabIndex = 30;
            this.cboSector.SelectedIndexChanged += new System.EventHandler(this.cboSector_SelectedIndexChanged);
            // 
            // cboIndustry
            // 
            this.cboIndustry.FormattingEnabled = true;
            this.cboIndustry.Location = new System.Drawing.Point(289, 37);
            this.cboIndustry.Name = "cboIndustry";
            this.cboIndustry.Size = new System.Drawing.Size(165, 21);
            this.cboIndustry.TabIndex = 31;
            // 
            // FrmScreener
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(605, 533);
            this.Controls.Add(this.cboIndustry);
            this.Controls.Add(this.cboSector);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtExchange);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtCode);
            this.Controls.Add(this.label1);
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
            this.Controls.Add(this.dataGridViewFilters);
            this.Name = "FrmScreener";
            this.ShowIcon = false;
            this.Text = "New screener";
            this.Load += new System.EventHandler(this.FrmScreener_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewFilters)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLimit)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewFilters;
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
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtCode;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtExchange;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboSector;
        private System.Windows.Forms.ComboBox cboIndustry;
    }
}