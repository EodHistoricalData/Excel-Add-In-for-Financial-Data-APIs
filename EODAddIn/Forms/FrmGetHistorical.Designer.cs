﻿
namespace EODAddIn.Forms
{
    partial class FrmGetHistorical
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmGetHistorical));
            this.label2 = new System.Windows.Forms.Label();
            this.dtpFrom = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.dtpTo = new System.Windows.Forms.DateTimePicker();
            this.btnLoad = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.cboPeriod = new System.Windows.Forms.ComboBox();
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
            this.order_label = new System.Windows.Forms.Label();
            this.rbtnDescOrder = new System.Windows.Forms.RadioButton();
            this.rbtnAscOrder = new System.Windows.Forms.RadioButton();
            this.chkIsTable = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cboTypeOfOutput = new System.Windows.Forms.ComboBox();
            this.chkDateToWS = new System.Windows.Forms.CheckBox();
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
            this.dtpFrom.Location = new System.Drawing.Point(101, 275);
            this.dtpFrom.Name = "dtpFrom";
            this.dtpFrom.Size = new System.Drawing.Size(171, 20);
            this.dtpFrom.TabIndex = 5;
            this.dtpFrom.Value = new System.DateTime(2020, 9, 17, 0, 0, 0, 0);
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
            this.dtpTo.Location = new System.Drawing.Point(101, 301);
            this.dtpTo.Name = "dtpTo";
            this.dtpTo.Size = new System.Drawing.Size(171, 20);
            this.dtpTo.TabIndex = 7;
            this.dtpTo.Value = new System.DateTime(2022, 9, 18, 13, 37, 0, 0);
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(197, 427);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(75, 23);
            this.btnLoad.TabIndex = 8;
            this.btnLoad.Text = "Get Data";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 250);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(37, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Period";
            // 
            // cboPeriod
            // 
            this.cboPeriod.FormattingEnabled = true;
            this.cboPeriod.Items.AddRange(new object[] {
            "daily",
            "weekly",
            "monthly"});
            this.cboPeriod.Location = new System.Drawing.Point(101, 247);
            this.cboPeriod.Name = "cboPeriod";
            this.cboPeriod.Size = new System.Drawing.Size(171, 21);
            this.cboPeriod.TabIndex = 3;
            // 
            // gridTickers
            // 
            this.gridTickers.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridTickers.BackgroundColor = System.Drawing.SystemColors.Window;
            this.gridTickers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridTickers.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            this.gridTickers.ContextMenuStrip = this.contextMenuStrip1;
            this.gridTickers.Location = new System.Drawing.Point(12, 27);
            this.gridTickers.Name = "gridTickers";
            this.gridTickers.RowHeadersVisible = false;
            this.gridTickers.RowHeadersWidth = 20;
            this.gridTickers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridTickers.Size = new System.Drawing.Size(260, 209);
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
            this.contextMenuStrip1.Size = new System.Drawing.Size(136, 48);
            // 
            // findTickerToolStripMenuItem
            // 
            this.findTickerToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("findTickerToolStripMenuItem.Image")));
            this.findTickerToolStripMenuItem.Name = "findTickerToolStripMenuItem";
            this.findTickerToolStripMenuItem.Size = new System.Drawing.Size(135, 22);
            this.findTickerToolStripMenuItem.Text = "Find ticker";
            this.findTickerToolStripMenuItem.Click += new System.EventHandler(this.TsmiFindTicker_Click);
            // 
            // tsmiDeleteRowDataGrid
            // 
            this.tsmiDeleteRowDataGrid.Image = ((System.Drawing.Image)(resources.GetObject("tsmiDeleteRowDataGrid.Image")));
            this.tsmiDeleteRowDataGrid.Name = "tsmiDeleteRowDataGrid";
            this.tsmiDeleteRowDataGrid.Size = new System.Drawing.Size(135, 22);
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
            this.menuStrip1.Size = new System.Drawing.Size(284, 25);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tsmiFindTicker
            // 
            this.tsmiFindTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFindTicker.Image")));
            this.tsmiFindTicker.Name = "tsmiFindTicker";
            this.tsmiFindTicker.Size = new System.Drawing.Size(95, 21);
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
            this.tsmiLoadTickers.Size = new System.Drawing.Size(75, 21);
            this.tsmiLoadTickers.Text = "Import";
            // 
            // tsmiFromTxt
            // 
            this.tsmiFromTxt.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFromTxt.Image")));
            this.tsmiFromTxt.Name = "tsmiFromTxt";
            this.tsmiFromTxt.Size = new System.Drawing.Size(177, 22);
            this.tsmiFromTxt.Text = "From file txt";
            this.tsmiFromTxt.Click += new System.EventHandler(this.TsmiFromTxt_Click);
            // 
            // tsmiFromExcel
            // 
            this.tsmiFromExcel.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFromExcel.Image")));
            this.tsmiFromExcel.Name = "tsmiFromExcel";
            this.tsmiFromExcel.Size = new System.Drawing.Size(177, 22);
            this.tsmiFromExcel.Text = "From Excel range";
            this.tsmiFromExcel.Click += new System.EventHandler(this.TsmiFromExcel_Click);
            // 
            // tsmiClearTicker
            // 
            this.tsmiClearTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiClearTicker.Image")));
            this.tsmiClearTicker.Name = "tsmiClearTicker";
            this.tsmiClearTicker.Size = new System.Drawing.Size(86, 21);
            this.tsmiClearTicker.Text = "Clear list";
            this.tsmiClearTicker.Click += new System.EventHandler(this.ClearTicker_Click);
            // 
            // order_label
            // 
            this.order_label.AutoSize = true;
            this.order_label.Location = new System.Drawing.Point(12, 330);
            this.order_label.Name = "order_label";
            this.order_label.Size = new System.Drawing.Size(33, 13);
            this.order_label.TabIndex = 10;
            this.order_label.Text = "Order";
            // 
            // rbtnDescOrder
            // 
            this.rbtnDescOrder.AutoSize = true;
            this.rbtnDescOrder.Checked = true;
            this.rbtnDescOrder.Location = new System.Drawing.Point(101, 330);
            this.rbtnDescOrder.Name = "rbtnDescOrder";
            this.rbtnDescOrder.Size = new System.Drawing.Size(50, 17);
            this.rbtnDescOrder.TabIndex = 11;
            this.rbtnDescOrder.TabStop = true;
            this.rbtnDescOrder.Text = "Desc";
            this.rbtnDescOrder.UseVisualStyleBackColor = true;
            // 
            // rbtnAscOrder
            // 
            this.rbtnAscOrder.AutoSize = true;
            this.rbtnAscOrder.Location = new System.Drawing.Point(157, 330);
            this.rbtnAscOrder.Name = "rbtnAscOrder";
            this.rbtnAscOrder.Size = new System.Drawing.Size(43, 17);
            this.rbtnAscOrder.TabIndex = 12;
            this.rbtnAscOrder.Text = "Asc";
            this.rbtnAscOrder.UseVisualStyleBackColor = true;
            // 
            // chkIsTable
            // 
            this.chkIsTable.Checked = true;
            this.chkIsTable.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIsTable.Location = new System.Drawing.Point(12, 410);
            this.chkIsTable.Name = "chkIsTable";
            this.chkIsTable.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkIsTable.Size = new System.Drawing.Size(147, 20);
            this.chkIsTable.TabIndex = 23;
            this.chkIsTable.Text = "Smart Table";
            this.chkIsTable.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkIsTable.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 358);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 22;
            this.label1.Text = "Type of Output";
            // 
            // cboTypeOfOutput
            // 
            this.cboTypeOfOutput.FormattingEnabled = true;
            this.cboTypeOfOutput.Items.AddRange(new object[] {
            "One worksheet",
            "Separated with chart",
            "Separated without chart"});
            this.cboTypeOfOutput.Location = new System.Drawing.Point(111, 357);
            this.cboTypeOfOutput.Name = "cboTypeOfOutput";
            this.cboTypeOfOutput.Size = new System.Drawing.Size(161, 21);
            this.cboTypeOfOutput.TabIndex = 21;
            this.cboTypeOfOutput.Text = "One worksheet";
            // 
            // chkDateToWS
            // 
            this.chkDateToWS.Checked = true;
            this.chkDateToWS.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDateToWS.Location = new System.Drawing.Point(12, 384);
            this.chkDateToWS.Name = "chkDateToWS";
            this.chkDateToWS.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkDateToWS.Size = new System.Drawing.Size(147, 20);
            this.chkDateToWS.TabIndex = 24;
            this.chkDateToWS.Text = "Add date to sheet name";
            this.chkDateToWS.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDateToWS.UseVisualStyleBackColor = true;
            // 
            // FrmGetHistorical
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 462);
            this.Controls.Add(this.chkDateToWS);
            this.Controls.Add(this.chkIsTable);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboTypeOfOutput);
            this.Controls.Add(this.rbtnAscOrder);
            this.Controls.Add(this.rbtnDescOrder);
            this.Controls.Add(this.order_label);
            this.Controls.Add(this.gridTickers);
            this.Controls.Add(this.cboPeriod);
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
            this.Name = "FrmGetHistorical";
            this.ShowIcon = false;
            this.Text = "Historical data";
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
        private System.Windows.Forms.ComboBox cboPeriod;
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
        private System.Windows.Forms.Label order_label;
        private System.Windows.Forms.RadioButton rbtnDescOrder;
        private System.Windows.Forms.RadioButton rbtnAscOrder;
        private System.Windows.Forms.CheckBox chkIsTable;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboTypeOfOutput;
        private System.Windows.Forms.CheckBox chkDateToWS;
    }
}