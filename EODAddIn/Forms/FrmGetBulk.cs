using EOD.Model.BulkFundamental;
using EODAddIn.Program;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmGetBulk : Form
    {
        public Dictionary<string, BulkFundamentalData> Results;
        public string Exchange = Settings.SettingsFields.BulkFundamentalExchange;
        public List<string> Tickers = Settings.SettingsFields.BulkFundamentalTickers;
        public int Offset = Settings.SettingsFields.BulkFundamentalOffset;
        public int Limit = Settings.SettingsFields.BulkFundamentalLimit;

        public FrmGetBulk()
        {
            InitializeComponent();
            txtExchange.Text = Exchange;
            numOffset.Value = Offset;
            numLimit.Value = Limit;
            foreach (string ticker in Settings.SettingsFields.BulkFundamentalTickers)
            {
                int i = gridTickers.Rows.Add();
                gridTickers.Rows[i].Cells[0].Value = ticker;
            }
        }

        private void TsmiFindTicker_Click(object sender, EventArgs e)
        {
            FrmSearchTiker frm = new FrmSearchTiker();
            frm.ShowDialog();

            if (frm.Result.Code == null) return;

            int i = gridTickers.Rows.Add();

            gridTickers.Rows[i].Cells[0].Value = $"{frm.Result.Code}.{frm.Result.Exchange}";
        }

        private void ClearTicker_Click(object sender, EventArgs e)
        {
            gridTickers.Rows.Clear();
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return;

            try
            {
                Exchange = txtExchange.Text;
                Offset = (int)numOffset.Value;
                Limit = (int)numLimit.Value;
                Tickers.Clear();
                foreach (DataGridViewRow row in gridTickers.Rows)
                {
                    if (row.Cells[0].Value == null) continue;
                    string ticker = row.Cells[0].Value.ToString();
                    Tickers.Add(ticker);
                }
            }
            catch
            {

            }

            Settings.SettingsFields.BulkFundamentalExchange = Exchange;
            Settings.SettingsFields.BulkFundamentalTickers = Tickers;
            Settings.SettingsFields.BulkFundamentalOffset = Offset;
            Settings.SettingsFields.BulkFundamentalLimit = Limit;
            Settings.Save();
            string warning = "You are going to download " + Limit + " symbols.";
            if (Tickers.Count != 0)
            {
                warning = "You are going to download " + Tickers.Count + " symbols.";
            }
            else
            {
                if (Limit > 99)
                {
                    warning += " It might take longer.";
                }
            }
            DialogResult = MessageBox.Show(warning + " Do you want to proceed?", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
        }

        private bool CheckForm()
        {
            if (string.IsNullOrEmpty(txtExchange.Text))
            {
                MessageBox.Show("Insert an exchange", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }
    }
}
