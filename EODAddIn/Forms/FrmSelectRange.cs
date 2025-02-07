using EODAddIn.Controls;
using System;
using System.Drawing;
using System.Web.UI;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmSelectRange : Form
    {
        RefEdit RefEdit {  get; set; }
        public string RangeAddress { get; set; }
        public FrmSelectRange()
        {
            InitializeComponent();
            float scale = (float)Size.Width / 360;
            RefEdit = new RefEdit
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Location = new Point((int)(12 * scale), (int)(12 * scale)),
                Name = "refEdit1",
                Size = new Size((int)(318 * scale), (int)(20 * scale)),
                TabIndex = 0
            };
            this.Controls.Add(RefEdit);
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            RangeAddress = RefEdit.Text;
            Close();
        }
    }
}
