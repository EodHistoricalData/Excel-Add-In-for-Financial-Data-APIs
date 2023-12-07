namespace EODAddIn.Forms
{
    partial class LiveDownloaderDispatcher
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LiveDownloaderDispatcher));
            this.gridDownloaders = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmiFindTicker = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiLoadTickers = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiClearTicker = new System.Windows.Forms.ToolStripMenuItem();
            this.BtnAdd = new System.Windows.Forms.Button();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewImageColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewButtonColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewButtonColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewButtonColumn();
            ((System.ComponentModel.ISupportInitialize)(this.gridDownloaders)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // gridDownloaders
            // 
            this.gridDownloaders.BackgroundColor = System.Drawing.SystemColors.Window;
            this.gridDownloaders.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridDownloaders.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column6,
            this.Column3,
            this.Column4,
            this.Column5});
            this.gridDownloaders.Location = new System.Drawing.Point(12, 28);
            this.gridDownloaders.Name = "gridDownloaders";
            this.gridDownloaders.RowHeadersVisible = false;
            this.gridDownloaders.RowHeadersWidth = 20;
            this.gridDownloaders.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridDownloaders.Size = new System.Drawing.Size(569, 291);
            this.gridDownloaders.TabIndex = 3;
            this.gridDownloaders.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.GridDownloaders_CellContentClick);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiFindTicker,
            this.tsmiLoadTickers,
            this.tsmiClearTicker});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(593, 25);
            this.menuStrip1.TabIndex = 4;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tsmiFindTicker
            // 
            this.tsmiFindTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiFindTicker.Image")));
            this.tsmiFindTicker.Name = "tsmiFindTicker";
            this.tsmiFindTicker.Size = new System.Drawing.Size(80, 21);
            this.tsmiFindTicker.Text = "Start all";
            // 
            // tsmiLoadTickers
            // 
            this.tsmiLoadTickers.Image = ((System.Drawing.Image)(resources.GetObject("tsmiLoadTickers.Image")));
            this.tsmiLoadTickers.Name = "tsmiLoadTickers";
            this.tsmiLoadTickers.Size = new System.Drawing.Size(80, 21);
            this.tsmiLoadTickers.Text = "Stop all";
            // 
            // tsmiClearTicker
            // 
            this.tsmiClearTicker.Image = ((System.Drawing.Image)(resources.GetObject("tsmiClearTicker.Image")));
            this.tsmiClearTicker.Name = "tsmiClearTicker";
            this.tsmiClearTicker.Size = new System.Drawing.Size(90, 21);
            this.tsmiClearTicker.Text = "Delete all";
            // 
            // BtnAdd
            // 
            this.BtnAdd.Location = new System.Drawing.Point(506, 325);
            this.BtnAdd.Name = "BtnAdd";
            this.BtnAdd.Size = new System.Drawing.Size(75, 33);
            this.BtnAdd.TabIndex = 5;
            this.BtnAdd.Text = "Add";
            this.BtnAdd.UseVisualStyleBackColor = true;
            this.BtnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.Column1.HeaderText = "Name";
            this.Column1.Name = "Column1";
            this.Column1.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Column1.Width = 60;
            // 
            // Column2
            // 
            this.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column2.HeaderText = "Tickers";
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            this.Column2.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // Column6
            // 
            this.Column6.HeaderText = "Status";
            this.Column6.Name = "Column6";
            this.Column6.ReadOnly = true;
            this.Column6.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Column6.Width = 45;
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Start";
            this.Column3.Name = "Column3";
            this.Column3.ReadOnly = true;
            this.Column3.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Column3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.Column3.Width = 36;
            // 
            // Column4
            // 
            this.Column4.HeaderText = "Stop";
            this.Column4.Name = "Column4";
            this.Column4.ReadOnly = true;
            this.Column4.Width = 36;
            // 
            // Column5
            // 
            this.Column5.HeaderText = "Delete";
            this.Column5.Name = "Column5";
            this.Column5.ReadOnly = true;
            this.Column5.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Column5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.Column5.Text = "x";
            this.Column5.Width = 45;
            // 
            // LiveDownloaderDispatcher
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(593, 370);
            this.Controls.Add(this.BtnAdd);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.gridDownloaders);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LiveDownloaderDispatcher";
            this.ShowIcon = false;
            this.Text = "Live Data Downloaders";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.LiveDownloaderDispatcher_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.gridDownloaders)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView gridDownloaders;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsmiFindTicker;
        private System.Windows.Forms.ToolStripMenuItem tsmiLoadTickers;
        private System.Windows.Forms.ToolStripMenuItem tsmiClearTicker;
        private System.Windows.Forms.Button BtnAdd;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewImageColumn Column6;
        private System.Windows.Forms.DataGridViewButtonColumn Column3;
        private System.Windows.Forms.DataGridViewButtonColumn Column4;
        private System.Windows.Forms.DataGridViewButtonColumn Column5;
    }
}