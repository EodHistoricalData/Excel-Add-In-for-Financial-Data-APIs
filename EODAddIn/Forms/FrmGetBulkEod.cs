using EOD.Model.OptionsData;
using EODAddIn.BL;
using EODAddIn.Program;
using EODAddIn.Utils;
using MS.ProgressBar;
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
    public partial class FrmGetBulkEod : Form
    {
        public string Exchange = Settings.SettingsFields.BulkEodExchange;
        public string Type = Settings.SettingsFields.BulkEodType;
        public DateTime Date = Settings.SettingsFields.BulkEodDate;
        public List<string> Tickers = Settings.SettingsFields.BulkEodSymbols;

        public FrmGetBulkEod()
        {
            InitializeComponent();
            tbExchange.Text = Exchange;
            cboType.SelectedValue = Type;
            if (Date < dtpDate.MinDate)
            {
                dtpDate.Value = dtpDate.MinDate;
            }
            else
            {
                dtpDate.Value = Date;
            }
            foreach (string ticker in Tickers)
            {
                int i = gridTickers.Rows.Add();
                gridTickers.Rows[i].Cells[0].Value = ticker;
            }
        }

        private void BtnGet_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return;

            Exchange = tbExchange.Text;
            Type = cboType.SelectedItem.ToString();
            Date = dtpDate.Value;
            Tickers.Clear();

            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                Tickers.Add(row.Cells[0].Value.ToString());
            }

            Settings.SettingsFields.BulkEodExchange = Exchange;
            Settings.SettingsFields.BulkEodType = Type;
            Settings.SettingsFields.BulkEodDate = Date;
            Settings.SettingsFields.BulkEodSymbols = Tickers;
            Settings.Save();

            DialogResult = DialogResult.OK;
            Close();
        }

        private bool CheckForm()
        {
            return true;
        }
    }
}
