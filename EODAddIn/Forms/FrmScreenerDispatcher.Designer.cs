namespace EODAddIn.Forms
{
    partial class FrmScreenerDispatcher
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.newScreenerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.getFundamentalToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadScreenerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.getFundamentalToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.getHistoricalToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.getIntradayToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.editToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dataGridViewData = new System.Windows.Forms.DataGridView();
            this.colScreenerName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewData)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newScreenerToolStripMenuItem,
            this.getFundamentalToolStripMenuItem,
            this.editToolStripMenuItem,
            this.deleteToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(436, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.menuStrip1_ItemClicked);
            // 
            // newScreenerToolStripMenuItem
            // 
            this.newScreenerToolStripMenuItem.Image = global::EODAddIn.Properties.Resources.icons8_add_16;
            this.newScreenerToolStripMenuItem.Name = "newScreenerToolStripMenuItem";
            this.newScreenerToolStripMenuItem.Size = new System.Drawing.Size(106, 20);
            this.newScreenerToolStripMenuItem.Text = "New screener";
            this.newScreenerToolStripMenuItem.Click += new System.EventHandler(this.NewScreenerToolStripMenuItem_Click);
            // 
            // getFundamentalToolStripMenuItem
            // 
            this.getFundamentalToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.loadScreenerToolStripMenuItem,
            this.getFundamentalToolStripMenuItem1,
            this.getHistoricalToolStripMenuItem1,
            this.getIntradayToolStripMenuItem1});
            this.getFundamentalToolStripMenuItem.Image = global::EODAddIn.Properties.Resources.greenStatus;
            this.getFundamentalToolStripMenuItem.Name = "getFundamentalToolStripMenuItem";
            this.getFundamentalToolStripMenuItem.Size = new System.Drawing.Size(97, 20);
            this.getFundamentalToolStripMenuItem.Text = "Commands";
            // 
            // loadScreenerToolStripMenuItem
            // 
            this.loadScreenerToolStripMenuItem.Name = "loadScreenerToolStripMenuItem";
            this.loadScreenerToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.loadScreenerToolStripMenuItem.Text = "Load screener";
            this.loadScreenerToolStripMenuItem.Click += new System.EventHandler(this.LoadScreenerToolStripMenuItem_Click);
            // 
            // getFundamentalToolStripMenuItem1
            // 
            this.getFundamentalToolStripMenuItem1.Name = "getFundamentalToolStripMenuItem1";
            this.getFundamentalToolStripMenuItem1.Size = new System.Drawing.Size(180, 22);
            this.getFundamentalToolStripMenuItem1.Text = "Get fundamental";
            this.getFundamentalToolStripMenuItem1.Click += new System.EventHandler(this.GetFundamentalToolStripMenuItem1_Click);
            // 
            // getHistoricalToolStripMenuItem1
            // 
            this.getHistoricalToolStripMenuItem1.Name = "getHistoricalToolStripMenuItem1";
            this.getHistoricalToolStripMenuItem1.Size = new System.Drawing.Size(180, 22);
            this.getHistoricalToolStripMenuItem1.Text = "Get historical";
            this.getHistoricalToolStripMenuItem1.Click += new System.EventHandler(this.GetHistoricalToolStripMenuItem1_Click);
            // 
            // getIntradayToolStripMenuItem1
            // 
            this.getIntradayToolStripMenuItem1.Name = "getIntradayToolStripMenuItem1";
            this.getIntradayToolStripMenuItem1.Size = new System.Drawing.Size(180, 22);
            this.getIntradayToolStripMenuItem1.Text = "Get intraday";
            this.getIntradayToolStripMenuItem1.Click += new System.EventHandler(this.GetIntradayToolStripMenuItem1_Click);
            // 
            // editToolStripMenuItem
            // 
            this.editToolStripMenuItem.Image = global::EODAddIn.Properties.Resources.icons8_edit_16__1_;
            this.editToolStripMenuItem.Name = "editToolStripMenuItem";
            this.editToolStripMenuItem.Size = new System.Drawing.Size(55, 20);
            this.editToolStripMenuItem.Text = "Edit";
            this.editToolStripMenuItem.Click += new System.EventHandler(this.EditToolStripMenuItem_Click);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Image = global::EODAddIn.Properties.Resources.icons8_close_16;
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(68, 20);
            this.deleteToolStripMenuItem.Text = "Delete";
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.DeleteToolStripMenuItem_Click);
            // 
            // dataGridViewData
            // 
            this.dataGridViewData.AllowUserToAddRows = false;
            this.dataGridViewData.AllowUserToDeleteRows = false;
            this.dataGridViewData.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridViewData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colScreenerName});
            this.dataGridViewData.Location = new System.Drawing.Point(12, 27);
            this.dataGridViewData.Name = "dataGridViewData";
            this.dataGridViewData.ReadOnly = true;
            this.dataGridViewData.RowHeadersVisible = false;
            this.dataGridViewData.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewData.Size = new System.Drawing.Size(412, 192);
            this.dataGridViewData.TabIndex = 1;
            // 
            // colScreenerName
            // 
            this.colScreenerName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colScreenerName.HeaderText = "Name";
            this.colScreenerName.Name = "colScreenerName";
            this.colScreenerName.ReadOnly = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 226);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(347, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "How to use: Create a new screener by pressing \"New screener\". When ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 239);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(325, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "the screener is ready, you can export its results to Sheets or perform";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 252);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(370, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "an EOD Bulk request and Fundamental data Bulk request. The result will also";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 265);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(205, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "be exported to your sheets in a new page.";
            // 
            // FrmScreenerDispatcher
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(436, 288);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridViewData);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmScreenerDispatcher";
            this.ShowIcon = false;
            this.Text = "Screener dispatcher";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem newScreenerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem getFundamentalToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem getFundamentalToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem getHistoricalToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem getIntradayToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
        private System.Windows.Forms.DataGridView dataGridViewData;
        private System.Windows.Forms.DataGridViewTextBoxColumn colScreenerName;
        private System.Windows.Forms.ToolStripMenuItem loadScreenerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}