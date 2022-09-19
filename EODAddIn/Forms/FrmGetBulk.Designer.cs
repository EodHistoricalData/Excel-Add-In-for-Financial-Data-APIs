namespace EODAddIn.Forms
{
    partial class FrmGetBulk
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmGetBulk));
            this.txtExchange = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnLoad = new System.Windows.Forms.Button();
            this.numOffset = new System.Windows.Forms.NumericUpDown();
            this.numLimit = new System.Windows.Forms.NumericUpDown();
            this.gridTickers = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmiFindTicker = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiClearTicker = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.numOffset)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLimit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridTickers)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtExchange
            // 
            this.txtExchange.Location = new System.Drawing.Point(76, 27);
            this.txtExchange.Name = "txtExchange";
            this.txtExchange.Size = new System.Drawing.Size(160, 20);
            this.txtExchange.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Exchange";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 270);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Offset";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(18, 296);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(28, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Limit";
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(161, 291);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(75, 23);
            this.btnLoad.TabIndex = 19;
            this.btnLoad.Text = "Load";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // numOffset
            // 
            this.numOffset.Location = new System.Drawing.Point(55, 268);
            this.numOffset.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.numOffset.Name = "numOffset";
            this.numOffset.Size = new System.Drawing.Size(79, 20);
            this.numOffset.TabIndex = 20;
            // 
            // numLimit
            // 
            this.numLimit.Location = new System.Drawing.Point(55, 294);
            this.numLimit.Maximum = new decimal(new int[] {
            500,
            0,
            0,
            0});
            this.numLimit.Name = "numLimit";
            this.numLimit.Size = new System.Drawing.Size(79, 20);
            this.numLimit.TabIndex = 21;
            this.numLimit.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            // 
            // gridTickers
            // 
            this.gridTickers.BackgroundColor = System.Drawing.SystemColors.Window;
            this.gridTickers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridTickers.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            this.gridTickers.Location = new System.Drawing.Point(13, 53);
            this.gridTickers.Name = "gridTickers";
            this.gridTickers.RowHeadersVisible = false;
            this.gridTickers.RowHeadersWidth = 20;
            this.gridTickers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridTickers.Size = new System.Drawing.Size(223, 209);
            this.gridTickers.TabIndex = 22;
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column1.HeaderText = "Tickers";
            this.Column1.Name = "Column1";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiFindTicker,
            this.tsmiClearTicker});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(248, 24);
            this.menuStrip1.TabIndex = 23;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tsmiFindTicker
            // 
            this.tsmiFindTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFindTicker.Image")));
            this.tsmiFindTicker.Name = "tsmiFindTicker";
            this.tsmiFindTicker.Size = new System.Drawing.Size(90, 20);
            this.tsmiFindTicker.Text = "Find ticker";
            this.tsmiFindTicker.Click += new System.EventHandler(this.TsmiFindTicker_Click);
            // 
            // tsmiClearTicker
            // 
            this.tsmiClearTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiClearTicker.Image")));
            this.tsmiClearTicker.Name = "tsmiClearTicker";
            this.tsmiClearTicker.Size = new System.Drawing.Size(80, 20);
            this.tsmiClearTicker.Text = "Clear list";
            this.tsmiClearTicker.Click += new System.EventHandler(this.ClearTicker_Click);
            // 
            // FrmGetBulk
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(248, 326);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.gridTickers);
            this.Controls.Add(this.numLimit);
            this.Controls.Add(this.numOffset);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtExchange);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmGetBulk";
            this.ShowIcon = false;
            this.Text = "Bulk Fundamentals";
            ((System.ComponentModel.ISupportInitialize)(this.numOffset)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLimit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridTickers)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtExchange;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.NumericUpDown numOffset;
        private System.Windows.Forms.NumericUpDown numLimit;
        private System.Windows.Forms.DataGridView gridTickers;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsmiFindTicker;
        private System.Windows.Forms.ToolStripMenuItem tsmiClearTicker;
    }
}