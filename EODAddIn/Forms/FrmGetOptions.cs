using EOD.Model.OptionsData;
using EODAddIn.Program;
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
    public partial class FrmGetOptions : Form
    {
        public string Ticker = Settings.SettingsFields.OptionsTicker;
        public DateTime From;
        public DateTime To;
        public DateTime? FromTrade;
        public DateTime? ToTrade;

        public FrmGetOptions()
        {
            InitializeComponent();
            txtCode.Text = Ticker;
            dtpFrom.Value = DateTime.Today;
            dtpTo.Value = new DateTime(2100, 1, 1);
            dtpFromTrade.Enabled = false;
            dtpToTrade.Enabled = false;
        }

        private void ChkTradeDate_CheckedChanged(object sender, EventArgs e)
        {
            if (chkTradeDate.Checked == true)
            {
                dtpFromTrade.Enabled = true;
                dtpToTrade.Enabled = true;
            }
            else
            {
                dtpFromTrade.Enabled = false;
                dtpToTrade.Enabled = false;
            }
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return;

            Ticker = txtCode.Text;
            From = dtpFrom.Value;
            To = dtpTo.Value;
            if (chkTradeDate.Checked == true)
            {
                FromTrade = dtpFromTrade.Value;
                ToTrade = dtpToTrade.Value;
            }

            Settings.SettingsFields.OptionsTicker = Ticker;
            Settings.Save();
            DialogResult = DialogResult.OK;
            Close();
        }

        private bool CheckForm()
        {
            if (string.IsNullOrEmpty(txtCode.Text))
            {
                MessageBox.Show("Insert a tiсker", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (!txtCode.Text.Contains("."))
            {
                MessageBox.Show("Insert exchange", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            int compare = DateTime.Compare(dtpFrom.Value, dtpTo.Value);
            if (compare > 0)
            {
                MessageBox.Show("The \"From\" Date must be not later than the \"To\" Date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            compare = DateTime.Compare(dtpFromTrade.Value, dtpToTrade.Value);
            if (compare > 0)
            {
                MessageBox.Show("The \"From\" Date must be not later than the \"To\" Date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }
    }
}
