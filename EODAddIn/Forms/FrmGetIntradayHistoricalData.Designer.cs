﻿
namespace EODAddIn.Forms
{
    partial class FrmGetIntradayHistoricalData
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmGetIntradayHistoricalData));
            this.label2 = new System.Windows.Forms.Label();
            this.dtpFrom = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.dtpTo = new System.Windows.Forms.DateTimePicker();
            this.btnLoad = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.cboInterval = new System.Windows.Forms.ComboBox();
            this.gridTickers = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.findTickerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiDeleteRowDataGrid = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmiFindTicker = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiLoadTickers = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiFromTxt = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiFromExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiClearTicker = new System.Windows.Forms.ToolStripMenuItem();
            this.chkChart = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.gridTickers)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 277);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "From";
            // 
            // dtpFrom
            // 
            this.dtpFrom.CustomFormat = "dd.MM.yyyy hh:mm:ss";
            this.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpFrom.Location = new System.Drawing.Point(86, 273);
            this.dtpFrom.Name = "dtpFrom";
            this.dtpFrom.Size = new System.Drawing.Size(152, 20);
            this.dtpFrom.TabIndex = 5;
            this.dtpFrom.Value = new System.DateTime(1970, 1, 1, 0, 0, 0, 0);
            this.dtpFrom.ValueChanged += new System.EventHandler(this.dtpFrom_ValueChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 303);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(20, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "To";
            // 
            // dtpTo
            // 
            this.dtpTo.CustomFormat = "dd.MM.yyyy hh:mm:ss";
            this.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpTo.Location = new System.Drawing.Point(86, 299);
            this.dtpTo.Name = "dtpTo";
            this.dtpTo.Size = new System.Drawing.Size(152, 20);
            this.dtpTo.TabIndex = 7;
            this.dtpTo.ValueChanged += new System.EventHandler(this.dtpTo_ValueChanged);
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(163, 345);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(75, 23);
            this.btnLoad.TabIndex = 8;
            this.btnLoad.Text = "Load";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 250);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(42, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Interval";
            // 
            // cboInterval
            // 
            this.cboInterval.FormattingEnabled = true;
            this.cboInterval.Items.AddRange(new object[] {
            "5m",
            "1h",
            "1m"});
            this.cboInterval.Location = new System.Drawing.Point(86, 245);
            this.cboInterval.Name = "cboInterval";
            this.cboInterval.Size = new System.Drawing.Size(152, 21);
            this.cboInterval.TabIndex = 3;
            this.cboInterval.SelectedIndexChanged += new System.EventHandler(this.cboInterval_SelectedIndexChanged);
            // 
            // gridTickers
            // 
            this.gridTickers.BackgroundColor = System.Drawing.SystemColors.Window;
            this.gridTickers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridTickers.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            this.gridTickers.ContextMenuStrip = this.contextMenuStrip1;
            this.gridTickers.Location = new System.Drawing.Point(15, 30);
            this.gridTickers.Name = "gridTickers";
            this.gridTickers.RowHeadersVisible = false;
            this.gridTickers.RowHeadersWidth = 20;
            this.gridTickers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridTickers.Size = new System.Drawing.Size(223, 209);
            this.gridTickers.TabIndex = 1;
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column1.HeaderText = "Tickers";
            this.Column1.Name = "Column1";
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.findTickerToolStripMenuItem,
            this.tsmiDeleteRowDataGrid});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(124, 48);
            // 
            // findTickerToolStripMenuItem
            // 
            this.findTickerToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("findTickerToolStripMenuItem.Image")));
            this.findTickerToolStripMenuItem.Name = "findTickerToolStripMenuItem";
            this.findTickerToolStripMenuItem.Size = new System.Drawing.Size(123, 22);
            this.findTickerToolStripMenuItem.Text = "Find ticker";
            this.findTickerToolStripMenuItem.Click += new System.EventHandler(this.TsmiFindTicker_Click);
            // 
            // tsmiDeleteRowDataGrid
            // 
            this.tsmiDeleteRowDataGrid.Image = ((System.Drawing.Image)(resources.GetObject("tsmiDeleteRowDataGrid.Image")));
            this.tsmiDeleteRowDataGrid.Name = "tsmiDeleteRowDataGrid";
            this.tsmiDeleteRowDataGrid.Size = new System.Drawing.Size(123, 22);
            this.tsmiDeleteRowDataGrid.Text = "Delete";
            this.tsmiDeleteRowDataGrid.Click += new System.EventHandler(this.TsmiDeleteRowDataGrid_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiFindTicker,
            this.tsmiLoadTickers,
            this.tsmiClearTicker});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(250, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tsmiFindTicker
            // 
            this.tsmiFindTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFindTicker.Image")));
            this.tsmiFindTicker.Name = "tsmiFindTicker";
            this.tsmiFindTicker.Size = new System.Drawing.Size(84, 20);
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
            this.tsmiLoadTickers.Size = new System.Drawing.Size(67, 20);
            this.tsmiLoadTickers.Text = "Import";
            // 
            // tsmiFromTxt
            // 
            this.tsmiFromTxt.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFromTxt.Image")));
            this.tsmiFromTxt.Name = "tsmiFromTxt";
            this.tsmiFromTxt.Size = new System.Drawing.Size(157, 22);
            this.tsmiFromTxt.Text = "From file txt";
            this.tsmiFromTxt.Click += new System.EventHandler(this.TsmiFromTxt_Click);
            // 
            // tsmiFromExcel
            // 
            this.tsmiFromExcel.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFromExcel.Image")));
            this.tsmiFromExcel.Name = "tsmiFromExcel";
            this.tsmiFromExcel.Size = new System.Drawing.Size(157, 22);
            this.tsmiFromExcel.Text = "From Excel range";
            this.tsmiFromExcel.Click += new System.EventHandler(this.TsmiFromExcel_Click);
            // 
            // tsmiClearTicker
            // 
            this.tsmiClearTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiClearTicker.Image")));
            this.tsmiClearTicker.Name = "tsmiClearTicker";
            this.tsmiClearTicker.Size = new System.Drawing.Size(76, 20);
            this.tsmiClearTicker.Text = "Clear list";
            this.tsmiClearTicker.Click += new System.EventHandler(this.ClearTicker_Click);
            // 
            // chkChart
            // 
            this.chkChart.Location = new System.Drawing.Point(12, 325);
            this.chkChart.Name = "chkChart";
            this.chkChart.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkChart.Size = new System.Drawing.Size(88, 24);
            this.chkChart.TabIndex = 9;
            this.chkChart.Text = "Chart";
            this.chkChart.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkChart.UseVisualStyleBackColor = true;
            // 
            // FrmGetIntradayHistoricalData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(250, 380);
            this.Controls.Add(this.chkChart);
            this.Controls.Add(this.gridTickers);
            this.Controls.Add(this.cboInterval);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.dtpTo);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dtpFrom);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmGetIntradayHistoricalData";
            this.ShowIcon = false;
            this.Text = "Intraday Historical Data";
            ((System.ComponentModel.ISupportInitialize)(this.gridTickers)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpFrom;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtpTo;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboInterval;
        private System.Windows.Forms.DataGridView gridTickers;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsmiFindTicker;
        private System.Windows.Forms.ToolStripMenuItem tsmiClearTicker;
        private System.Windows.Forms.ToolStripMenuItem findTickerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmiDeleteRowDataGrid;
        private System.Windows.Forms.ToolStripMenuItem tsmiLoadTickers;
        private System.Windows.Forms.ToolStripMenuItem tsmiFromTxt;
        private System.Windows.Forms.ToolStripMenuItem tsmiFromExcel;
        private System.Windows.Forms.CheckBox chkChart;
    }
}