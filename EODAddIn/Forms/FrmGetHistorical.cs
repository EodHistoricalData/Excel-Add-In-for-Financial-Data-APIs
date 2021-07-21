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
    public partial class FrmGetHistorical : Form
    {
        public List<Model.EndOfDay> Results;
        public string Tiker;
        public string Exchange;
        public string Period;
        public DateTime From;
        public DateTime To;

        public FrmGetHistorical()
        {
            InitializeComponent();
            cboPeriod.SelectedIndex = 0;
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return; 

            Tiker = txtCode.Text;
            Exchange = txtExchange.Text;

            Period = cboPeriod.SelectedItem.ToString().ToLower().Substring(0, 1);
            From = dtpFrom.Value;
            To = dtpTo.Value;

            Results = Utils.APIEOD.GetEOD($"{Tiker}.{Exchange}", From, To, Period);

            Close();
        }

        private bool CheckForm()
        {
            if (string.IsNullOrEmpty(txtCode.Text))
            {
                MessageBox.Show("Insert a tiсker", "Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
                return false;
            }
            if (string.IsNullOrEmpty(txtExchange.Text))
            {
                MessageBox.Show("Insert exchange", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (cboPeriod.SelectedIndex == -1)
            {
                MessageBox.Show("Select period", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (dtpFrom.Value == null)
            {
                MessageBox.Show("Select start date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (dtpTo.Value == null)
            {
                MessageBox.Show("Select end date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }
    }
}
