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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;


namespace EODAddIn.Forms
{
    public partial class FrmGetBulkEod : Form
    {
        public string Exchange = Settings.Data.BulkEodExchange;
        public string Type = Settings.Data.BulkEodType;
        public DateTime Date = Settings.Data.BulkEodDate;
        public List<string> Tickers = Settings.Data.BulkEodSymbols;

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

            Settings.Data.BulkEodExchange = Exchange;
            Settings.Data.BulkEodType = Type;
            Settings.Data.BulkEodDate = Date;
            Settings.Data.BulkEodSymbols = Tickers;
            Settings.Save();

            DialogResult = DialogResult.OK;
            Close();
        }

        private bool CheckForm()
        {
            if (cboType.SelectedIndex == -1)
            {
                MessageBox.Show("Select type please", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        private void tsmiFindTicker_Click(object sender, EventArgs e)
        {
            FrmSearchTiker frm = new FrmSearchTiker();
            frm.ShowDialog();

            if (frm.Result.Code == null) return;

            int i = gridTickers.Rows.Add();

            gridTickers.Rows[i].Cells[0].Value = $"{frm.Result.Code}.{frm.Result.Exchange}";
        }

        private void tsmiClearTicker_Click(object sender, EventArgs e)
        {
            gridTickers.Rows.Clear();
        }

        private void tsmiFromTxt_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "txt files (*.txt)|*.txt";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;

                    using (StreamReader fstream = new StreamReader(filePath))
                    {
                        while (!fstream.EndOfStream)
                        {
                            string text = fstream.ReadLine();
                            int i = gridTickers.Rows.Add();
                            gridTickers.Rows[i].Cells[0].Value = text;
                        }
                        fstream.Close();
                    }
                }
            }
        }

        private void tsmiFromExcel_Click(object sender, EventArgs e)
        {
            FrmSelectRange frm = new FrmSelectRange();
            frm.ShowDialog(new WinHwnd());
            tsmiFromExcel.Enabled = false;
            frm.FormClosing += FrmSelectRangeClosing;
        }

        private void FrmSelectRangeClosing(object sender, FormClosingEventArgs e)
        {
            FrmSelectRange frm = (FrmSelectRange)sender;
            if (ExcelUtils.IsRange(frm.RangeAddress))
            {
                Excel.Range range = Globals.ThisAddIn.Application.Range[frm.RangeAddress];

                foreach (Excel.Range cell in range)
                {
                    string txt = cell.Text;
                    if (!string.IsNullOrEmpty(txt))
                    {
                        int i = gridTickers.Rows.Add();
                        gridTickers.Rows[i].Cells[0].Value = cell.Text;
                    }
                }
            }
            tsmiFromExcel.Enabled = true;
        }
    }
}
