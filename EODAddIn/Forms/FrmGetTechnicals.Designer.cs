namespace EODAddIn.Forms
{
    partial class FrmGetTechnicals
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmGetTechnicals));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmiFindTicker = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiLoadTickers = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiFromTxt = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiFromExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiClearTicker = new System.Windows.Forms.ToolStripMenuItem();
            this.gridTickers = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.rbtnAscOrder = new System.Windows.Forms.RadioButton();
            this.rbtnDescOrder = new System.Windows.Forms.RadioButton();
            this.order_label = new System.Windows.Forms.Label();
            this.cboFunction = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.dtpTo = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.dtpFrom = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.cboAggPeriod = new System.Windows.Forms.ComboBox();
            this.labelFirstOption = new System.Windows.Forms.Label();
            this.labelSecondOption = new System.Windows.Forms.Label();
            this.labelThirdOption = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cboTypeOfOutput = new System.Windows.Forms.ComboBox();
            this.tbSecondOption = new System.Windows.Forms.TextBox();
            this.tbThirdOption = new System.Windows.Forms.TextBox();
            this.tbFirstOption = new System.Windows.Forms.TextBox();
            this.chkIsTable = new System.Windows.Forms.CheckBox();
            this.btnLoad = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridTickers)).BeginInit();
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
            // 
            // tsmiFindTicker
            // 
            this.tsmiFindTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFindTicker.Image")));
            this.tsmiFindTicker.Name = "tsmiFindTicker";
            this.tsmiFindTicker.Size = new System.Drawing.Size(90, 20);
            this.tsmiFindTicker.Text = "Find ticker";
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
            this.tsmiFromTxt.Size = new System.Drawing.Size(165, 22);
            this.tsmiFromTxt.Text = "From file txt";
            // 
            // tsmiFromExcel
            // 
            this.tsmiFromExcel.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFromExcel.Image")));
            this.tsmiFromExcel.Name = "tsmiFromExcel";
            this.tsmiFromExcel.Size = new System.Drawing.Size(165, 22);
            this.tsmiFromExcel.Text = "From Excel range";
            // 
            // tsmiClearTicker
            // 
            this.tsmiClearTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiClearTicker.Image")));
            this.tsmiClearTicker.Name = "tsmiClearTicker";
            this.tsmiClearTicker.Size = new System.Drawing.Size(80, 20);
            this.tsmiClearTicker.Text = "Clear list";
            this.tsmiClearTicker.Click += new System.EventHandler(this.ClearTicker_Click);
            // 
            // gridTickers
            // 
            this.gridTickers.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridTickers.BackgroundColor = System.Drawing.SystemColors.Window;
            this.gridTickers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridTickers.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            this.gridTickers.Location = new System.Drawing.Point(12, 27);
            this.gridTickers.Name = "gridTickers";
            this.gridTickers.RowHeadersVisible = false;
            this.gridTickers.RowHeadersWidth = 20;
            this.gridTickers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridTickers.Size = new System.Drawing.Size(260, 152);
            this.gridTickers.TabIndex = 2;
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column1.HeaderText = "Tickers";
            this.Column1.Name = "Column1";
            // 
            // rbtnAscOrder
            // 
            this.rbtnAscOrder.AutoSize = true;
            this.rbtnAscOrder.Location = new System.Drawing.Point(176, 241);
            this.rbtnAscOrder.Name = "rbtnAscOrder";
            this.rbtnAscOrder.Size = new System.Drawing.Size(43, 17);
            this.rbtnAscOrder.TabIndex = 21;
            this.rbtnAscOrder.Text = "Asc";
            this.rbtnAscOrder.UseVisualStyleBackColor = true;
            // 
            // rbtnDescOrder
            // 
            this.rbtnDescOrder.AutoSize = true;
            this.rbtnDescOrder.Checked = true;
            this.rbtnDescOrder.Location = new System.Drawing.Point(120, 241);
            this.rbtnDescOrder.Name = "rbtnDescOrder";
            this.rbtnDescOrder.Size = new System.Drawing.Size(50, 17);
            this.rbtnDescOrder.TabIndex = 20;
            this.rbtnDescOrder.TabStop = true;
            this.rbtnDescOrder.Text = "Desc";
            this.rbtnDescOrder.UseVisualStyleBackColor = true;
            // 
            // order_label
            // 
            this.order_label.AutoSize = true;
            this.order_label.Location = new System.Drawing.Point(12, 245);
            this.order_label.Name = "order_label";
            this.order_label.Size = new System.Drawing.Size(33, 13);
            this.order_label.TabIndex = 19;
            this.order_label.Text = "Order";
            // 
            // cboFunction
            // 
            this.cboFunction.FormattingEnabled = true;
            this.cboFunction.Items.AddRange(new object[] {
            "Average Volume",
            "Average Volume by Price",
            "Simple Moving Average",
            "Exponential Moving Average",
            "Weighted Moving Average",
            "Volatility",
            "Relative Strength Index",
            "Standard Deviation",
            "Slope (Linear Regression)",
            "Directional Movement Index",
            "Average Directional Movement Index",
            "Average True Range",
            "Commodity Channel Index",
            "Bollinger Bands",
            "Split Adjusted Data",
            "Stochastic Technical Indicator",
            "Stochastic Relative Strength Index",
            "Moving Average Convergence/Divergence",
            "Parabolic SAR"});
            this.cboFunction.Location = new System.Drawing.Point(120, 269);
            this.cboFunction.Name = "cboFunction";
            this.cboFunction.Size = new System.Drawing.Size(152, 21);
            this.cboFunction.TabIndex = 14;
            this.cboFunction.SelectedIndexChanged += new System.EventHandler(this.CboFunction_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 274);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(48, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Function";
            // 
            // dtpTo
            // 
            this.dtpTo.Location = new System.Drawing.Point(120, 212);
            this.dtpTo.Name = "dtpTo";
            this.dtpTo.Size = new System.Drawing.Size(152, 20);
            this.dtpTo.TabIndex = 18;
            this.dtpTo.Value = new System.DateTime(2022, 9, 18, 13, 37, 0, 0);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 218);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(20, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "To";
            // 
            // dtpFrom
            // 
            this.dtpFrom.Location = new System.Drawing.Point(120, 186);
            this.dtpFrom.Name = "dtpFrom";
            this.dtpFrom.Size = new System.Drawing.Size(152, 20);
            this.dtpFrom.TabIndex = 16;
            this.dtpFrom.Value = new System.DateTime(2020, 9, 17, 0, 0, 0, 0);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 192);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "From";
            // 
            // cboAggPeriod
            // 
            this.cboAggPeriod.FormattingEnabled = true;
            this.cboAggPeriod.Items.AddRange(new object[] {
            "Daily",
            "Weekly",
            "Monthly"});
            this.cboAggPeriod.Location = new System.Drawing.Point(207, 297);
            this.cboAggPeriod.Name = "cboAggPeriod";
            this.cboAggPeriod.Size = new System.Drawing.Size(65, 21);
            this.cboAggPeriod.TabIndex = 23;
            this.cboAggPeriod.Visible = false;
            // 
            // labelFirstOption
            // 
            this.labelFirstOption.AutoSize = true;
            this.labelFirstOption.Location = new System.Drawing.Point(12, 301);
            this.labelFirstOption.Name = "labelFirstOption";
            this.labelFirstOption.Size = new System.Drawing.Size(48, 13);
            this.labelFirstOption.TabIndex = 22;
            this.labelFirstOption.Text = "Function";
            this.labelFirstOption.Visible = false;
            // 
            // labelSecondOption
            // 
            this.labelSecondOption.AutoSize = true;
            this.labelSecondOption.Location = new System.Drawing.Point(12, 328);
            this.labelSecondOption.Name = "labelSecondOption";
            this.labelSecondOption.Size = new System.Drawing.Size(48, 13);
            this.labelSecondOption.TabIndex = 24;
            this.labelSecondOption.Text = "Function";
            this.labelSecondOption.Visible = false;
            // 
            // labelThirdOption
            // 
            this.labelThirdOption.AutoSize = true;
            this.labelThirdOption.Location = new System.Drawing.Point(12, 355);
            this.labelThirdOption.Name = "labelThirdOption";
            this.labelThirdOption.Size = new System.Drawing.Size(48, 13);
            this.labelThirdOption.TabIndex = 26;
            this.labelThirdOption.Text = "Function";
            this.labelThirdOption.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 385);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 29;
            this.label1.Text = "Type of Output";
            // 
            // cboTypeOfOutput
            // 
            this.cboTypeOfOutput.FormattingEnabled = true;
            this.cboTypeOfOutput.Items.AddRange(new object[] {
            "One worksheet",
            "Separated with chart",
            "Separated without chart"});
            this.cboTypeOfOutput.Location = new System.Drawing.Point(120, 382);
            this.cboTypeOfOutput.Name = "cboTypeOfOutput";
            this.cboTypeOfOutput.Size = new System.Drawing.Size(152, 21);
            this.cboTypeOfOutput.TabIndex = 28;
            this.cboTypeOfOutput.Text = "One worksheet";
            // 
            // tbSecondOption
            // 
            this.tbSecondOption.Location = new System.Drawing.Point(207, 324);
            this.tbSecondOption.Name = "tbSecondOption";
            this.tbSecondOption.Size = new System.Drawing.Size(65, 20);
            this.tbSecondOption.TabIndex = 30;
            // 
            // tbThirdOption
            // 
            this.tbThirdOption.Location = new System.Drawing.Point(207, 350);
            this.tbThirdOption.Name = "tbThirdOption";
            this.tbThirdOption.Size = new System.Drawing.Size(65, 20);
            this.tbThirdOption.TabIndex = 31;
            // 
            // tbFirstOption
            // 
            this.tbFirstOption.Location = new System.Drawing.Point(207, 298);
            this.tbFirstOption.Name = "tbFirstOption";
            this.tbFirstOption.Size = new System.Drawing.Size(65, 20);
            this.tbFirstOption.TabIndex = 32;
            // 
            // chkIsTable
            // 
            this.chkIsTable.Checked = true;
            this.chkIsTable.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIsTable.Location = new System.Drawing.Point(12, 401);
            this.chkIsTable.Name = "chkIsTable";
            this.chkIsTable.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkIsTable.Size = new System.Drawing.Size(122, 33);
            this.chkIsTable.TabIndex = 34;
            this.chkIsTable.Text = "Smart Table";
            this.chkIsTable.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkIsTable.UseVisualStyleBackColor = true;
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(188, 437);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(84, 23);
            this.btnLoad.TabIndex = 33;
            this.btnLoad.Text = "Download";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // FrmGetTechnicals
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 467);
            this.Controls.Add(this.chkIsTable);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.tbFirstOption);
            this.Controls.Add(this.tbThirdOption);
            this.Controls.Add(this.tbSecondOption);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboTypeOfOutput);
            this.Controls.Add(this.labelThirdOption);
            this.Controls.Add(this.labelSecondOption);
            this.Controls.Add(this.cboAggPeriod);
            this.Controls.Add(this.labelFirstOption);
            this.Controls.Add(this.rbtnAscOrder);
            this.Controls.Add(this.rbtnDescOrder);
            this.Controls.Add(this.order_label);
            this.Controls.Add(this.cboFunction);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.dtpTo);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dtpFrom);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.gridTickers);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmGetTechnicals";
            this.ShowIcon = false;
            this.Text = "Technical Indicators";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridTickers)).EndInit();
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
        private System.Windows.Forms.RadioButton rbtnAscOrder;
        private System.Windows.Forms.RadioButton rbtnDescOrder;
        private System.Windows.Forms.Label order_label;
        private System.Windows.Forms.ComboBox cboFunction;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker dtpTo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtpFrom;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboAggPeriod;
        private System.Windows.Forms.Label labelFirstOption;
        private System.Windows.Forms.Label labelSecondOption;
        private System.Windows.Forms.Label labelThirdOption;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboTypeOfOutput;
        private System.Windows.Forms.TextBox tbSecondOption;
        private System.Windows.Forms.TextBox tbThirdOption;
        private System.Windows.Forms.TextBox tbFirstOption;
        private System.Windows.Forms.CheckBox chkIsTable;
        private System.Windows.Forms.Button btnLoad;
    }
}