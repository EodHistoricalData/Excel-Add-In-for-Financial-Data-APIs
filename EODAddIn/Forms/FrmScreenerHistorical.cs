using EODAddIn.BL;
using EODAddIn.BL.Screener;
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
    public partial class FrmScreenerHistorical : Form
    {
        public FrmScreenerHistorical()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(Convert.ToString(period.SelectedItem)) || dateTimePicker2.Value.Subtract(dateTimePicker1.Value).TotalDays < 1 || DateTime.Today.Subtract(dateTimePicker2.Value).TotalDays < -1)
            {
                MessageBox.Show(
                    "incorrect input!",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
            ScreneerPrinter.PrintScreenerHistorical(dateTimePicker1.Value,dateTimePicker2.Value, Convert.ToString(period.SelectedItem));

            Close();
        }
    }
}
