using EODAddIn.BL;
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
    public partial class FrmScreenerIntraday : Form
    {
        public FrmScreenerIntraday()
        {
            InitializeComponent();
        }

        private void btnScreenLoadIntraday_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(Convert.ToString(comboBox1.SelectedItem)) || dateTimePicker2.Value.Subtract(dateTimePicker1.Value).TotalDays < 1 || DateTime.Today.Subtract(dateTimePicker2.Value).TotalDays < -1)
            {
                MessageBox.Show(
                    "incorrect input!",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
            LoadToExcel.PrintScreenerIntraday(dateTimePicker1.Value, dateTimePicker2.Value, Convert.ToString(comboBox1.SelectedItem));

            Close();
        }
    }
}
