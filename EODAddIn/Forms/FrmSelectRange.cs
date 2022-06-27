using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmSelectRange : Form
    {
        public string RangeAddress { get; set; }
        public FrmSelectRange()
        {
            InitializeComponent();
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            RangeAddress = refEdit1.Text;
            Close();
        }
    }
}
