namespace EODAddIn.Forms
{
    partial class FrmGetLive
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmGetLive));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmiFindTicker = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiLoadTickers = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiFromTxt = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiFromExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiClearTicker = new System.Windows.Forms.ToolStripMenuItem();
            this.gridTickers = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chkIsTable = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cboTypeOfOutput = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.NudInterval = new System.Windows.Forms.NumericUpDown();
            this.BtnCreate = new System.Windows.Forms.Button();
            this.BtnFilters = new System.Windows.Forms.Button();
            this.LblFilters = new System.Windows.Forms.Label();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.findTickerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiDeleteRowDataGrid = new System.Windows.Forms.ToolStripMenuItem();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridTickers)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.NudInterval)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiFindTicker,
            this.tsmiLoadTickers,
            this.tsmiClearTicker});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(284, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.Click += new System.EventHandler(this.ClearTicker_Click);
            // 
            // tsmiFindTicker
            // 
            this.tsmiFindTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFindTicker.Image")));
            this.tsmiFindTicker.Name = "tsmiFindTicker";
            this.tsmiFindTicker.Size = new System.Drawing.Size(90, 20);
            this.tsmiFindTicker.Text = "Find ticker";
            this.tsmiFindTicker.Click += new System.EventHandler(this.TsmiFindTicker_Click);
            // 
            // tsmiLoadTickers
            // 
            this.tsmiLoadTickers.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiFromTxt,
            this.tsmiFromExcel});
            this.tsmiLoadTickers.Image = ((System.Drawing.Image)(resources.GetObject("tsmiLoadTickers.Image")));
            this.tsmiLoadTickers.Name = "tsmiLoadTickers";
            this.tsmiLoadTickers.Size = new System.Drawing.Size(71, 20);
            this.tsmiLoadTickers.Text = "Import";
            // 
            // tsmiFromTxt
            // 
            this.tsmiFromTxt.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFromTxt.Image")));
            this.tsmiFromTxt.Name = "tsmiFromTxt";
            this.tsmiFromTxt.Size = new System.Drawing.Size(180, 22);
            this.tsmiFromTxt.Text = "From file txt";
            this.tsmiFromTxt.Click += new System.EventHandler(this.TsmiFromTxt_Click);
            // 
            // tsmiFromExcel
            // 
            this.tsmiFromExcel.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFromExcel.Image")));
            this.tsmiFromExcel.Name = "tsmiFromExcel";
            this.tsmiFromExcel.Size = new System.Drawing.Size(180, 22);
            this.tsmiFromExcel.Text = "From Excel range";
            this.tsmiFromExcel.Click += new System.EventHandler(this.TsmiFromExcel_Click);
            // 
            // tsmiClearTicker
            // 
            this.tsmiClearTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiClearTicker.Image")));
            this.tsmiClearTicker.Name = "tsmiClearTicker";
            this.tsmiClearTicker.Size = new System.Drawing.Size(80, 20);
            this.tsmiClearTicker.Text = "Clear list";
            // 
            // gridTickers
            // 
            this.gridTickers.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridTickers.BackgroundColor = System.Drawing.SystemColors.Window;
            this.gridTickers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridTickers.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            this.gridTickers.Location = new System.Drawing.Point(8, 27);
            this.gridTickers.Name = "gridTickers";
            this.gridTickers.RowHeadersVisible = false;
            this.gridTickers.RowHeadersWidth = 20;
            this.gridTickers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridTickers.Size = new System.Drawing.Size(264, 186);
            this.gridTickers.TabIndex = 2;
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column1.HeaderText = "Tickers";
            this.Column1.Name = "Column1";
            // 
            // chkIsTable
            // 
            this.chkIsTable.Checked = true;
            this.chkIsTable.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIsTable.Location = new System.Drawing.Point(15, 279);
            this.chkIsTable.Name = "chkIsTable";
            this.chkIsTable.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkIsTable.Size = new System.Drawing.Size(110, 18);
            this.chkIsTable.TabIndex = 24;
            this.chkIsTable.Text = "Smart table";
            this.chkIsTable.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkIsTable.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 255);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 26;
            this.label1.Text = "Type of Output";
            // 
            // cboTypeOfOutput
            // 
            this.cboTypeOfOutput.FormattingEnabled = true;
            this.cboTypeOfOutput.Items.AddRange(new object[] {
            "One worksheet",
            "Separated"});
            this.cboTypeOfOutput.Location = new System.Drawing.Point(108, 252);
            this.cboTypeOfOutput.Name = "cboTypeOfOutput";
            this.cboTypeOfOutput.Size = new System.Drawing.Size(124, 21);
            this.cboTypeOfOutput.TabIndex = 25;
            this.cboTypeOfOutput.Text = "One worksheet";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 228);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(85, 13);
            this.label5.TabIndex = 28;
            this.label5.Text = "Period (minutes):";
            // 
            // NudInterval
            // 
            this.NudInterval.Location = new System.Drawing.Point(108, 226);
            this.NudInterval.Maximum = new decimal(new int[] {
            1440,
            0,
            0,
            0});
            this.NudInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.NudInterval.Name = "NudInterval";
            this.NudInterval.Size = new System.Drawing.Size(124, 20);
            this.NudInterval.TabIndex = 29;
            this.NudInterval.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // BtnCreate
            // 
            this.BtnCreate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCreate.Location = new System.Drawing.Point(187, 66);
            this.BtnCreate.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.BtnCreate.MinimumSize = new System.Drawing.Size(75, 23);
            this.BtnCreate.Name = "BtnCreate";
            this.BtnCreate.Size = new System.Drawing.Size(75, 23);
            this.BtnCreate.TabIndex = 30;
            this.BtnCreate.Text = "Create";
            this.BtnCreate.UseVisualStyleBackColor = true;
            this.BtnCreate.Click += new System.EventHandler(this.BtnCreate_Click);
            // 
            // BtnFilters
            // 
            this.BtnFilters.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnFilters.Location = new System.Drawing.Point(3, 21);
            this.BtnFilters.Margin = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.BtnFilters.Name = "BtnFilters";
            this.BtnFilters.Size = new System.Drawing.Size(77, 23);
            this.BtnFilters.TabIndex = 31;
            this.BtnFilters.Text = "Filters";
            this.BtnFilters.UseVisualStyleBackColor = true;
            this.BtnFilters.Click += new System.EventHandler(this.BtnFilters_Click);
            // 
            // LblFilters
            // 
            this.LblFilters.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LblFilters.AutoSize = true;
            this.LblFilters.Location = new System.Drawing.Point(86, 3);
            this.LblFilters.Margin = new System.Windows.Forms.Padding(3);
            this.LblFilters.MaximumSize = new System.Drawing.Size(155, 0);
            this.LblFilters.Name = "LblFilters";
            this.LblFilters.Size = new System.Drawing.Size(155, 59);
            this.LblFilters.TabIndex = 32;
            this.LblFilters.Text = "All";
            this.LblFilters.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.findTickerToolStripMenuItem,
            this.tsmiDeleteRowDataGrid});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(130, 48);
            // 
            // findTickerToolStripMenuItem
            // 
            this.findTickerToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("findTickerToolStripMenuItem.Image")));
            this.findTickerToolStripMenuItem.Name = "findTickerToolStripMenuItem";
            this.findTickerToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.findTickerToolStripMenuItem.Text = "Find ticker";
            this.findTickerToolStripMenuItem.Click += new System.EventHandler(this.TsmiFindTicker_Click);
            // 
            // tsmiDeleteRowDataGrid
            // 
            this.tsmiDeleteRowDataGrid.Image = ((System.Drawing.Image)(resources.GetObject("tsmiDeleteRowDataGrid.Image")));
            this.tsmiDeleteRowDataGrid.Name = "tsmiDeleteRowDataGrid";
            this.tsmiDeleteRowDataGrid.Size = new System.Drawing.Size(129, 22);
            this.tsmiDeleteRowDataGrid.Text = "Delete";
            this.tsmiDeleteRowDataGrid.Click += new System.EventHandler(this.TsmiDeleteRowDataGrid_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31.51261F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 68.4874F));
            this.tableLayoutPanel1.Controls.Add(this.BtnFilters, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.BtnCreate, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.LblFilters, 1, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(8, 308);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 71.59091F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 28.40909F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(265, 92);
            this.tableLayoutPanel1.TabIndex = 33;
            // 
            // FrmGetLive
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 423);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.NudInterval);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboTypeOfOutput);
            this.Controls.Add(this.chkIsTable);
            this.Controls.Add(this.gridTickers);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmGetLive";
            this.ShowIcon = false;
            this.Text = "Live (Delayed) Stock Prices";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridTickers)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.NudInterval)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsmiFindTicker;
        private System.Windows.Forms.ToolStripMenuItem tsmiLoadTickers;
        private System.Windows.Forms.ToolStripMenuItem tsmiFromTxt;
        private System.Windows.Forms.ToolStripMenuItem tsmiFromExcel;
        private System.Windows.Forms.ToolStripMenuItem tsmiClearTicker;
        private System.Windows.Forms.DataGridView gridTickers;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.CheckBox chkIsTable;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboTypeOfOutput;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown NudInterval;
        private System.Windows.Forms.Button BtnCreate;
        private System.Windows.Forms.Button BtnFilters;
        private System.Windows.Forms.Label LblFilters;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem findTickerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmiDeleteRowDataGrid;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
    }
}