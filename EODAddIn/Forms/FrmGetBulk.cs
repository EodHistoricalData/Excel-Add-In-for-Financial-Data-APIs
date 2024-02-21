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
        public List<string> Tickers = Settings.Data.BulkFundamentalTickers;
        public string BulkTypeOfOutput;
        public FrmGetBulk()
        {
            InitializeComponent();
            foreach (string ticker in Settings.Data.BulkFundamentalTickers)
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
            try
            {
                Tickers.Clear();
                foreach (DataGridViewRow row in gridTickers.Rows)
                {
                    if (row.Cells[0].Value == null) continue;
                    string ticker = row.Cells[0].Value.ToString();
                    Tickers.Add(ticker);
                }
                BulkTypeOfOutput = cboTypeOfOutput.Text;
            }
            catch
            {

            }
            Settings.Data.BulkFundamentalTickers = Tickers;
            Settings.Save();
            //string warning = "You are going to download " + Tickers.Count + " symbols.";
            DialogResult =  DialogResult.OK;
        }
    }
}
