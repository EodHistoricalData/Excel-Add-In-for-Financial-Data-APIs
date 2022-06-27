namespace EODAddIn.Forms
{
    partial class FrmSelectRange
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
            this.refEdit1 = new EODAddIn.Controls.RefEdit();
            this.btnImport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // refEdit1
            // 
            this.refEdit1.Location = new System.Drawing.Point(12, 12);
            this.refEdit1.Name = "refEdit1";
            this.refEdit1.Size = new System.Drawing.Size(318, 20);
            this.refEdit1.TabIndex = 0;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(255, 38);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 1;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.BtnImport_Click);
            // 
            // FrmSelectRange
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(343, 68);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.refEdit1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmSelectRange";
            this.ShowIcon = false;
            this.Text = "Range select";
            this.ResumeLayout(false);

        }

        #endregion

        private Controls.RefEdit refEdit1;
        private System.Windows.Forms.Button btnImport;
    }
}