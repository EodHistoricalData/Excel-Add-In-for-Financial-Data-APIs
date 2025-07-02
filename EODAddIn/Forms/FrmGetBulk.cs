using EOD.Model.BulkFundamental;
using EODAddIn.Program;
using EODAddIn.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

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
            Close();
        }

        private void FromTxt_Click(object sender, EventArgs e)
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

        private void FromExcel_Click(object sender, EventArgs e)
        {
            FrmSelectRange frm = new FrmSelectRange();
            tsmiFromExcel.Enabled = false;
            frm.Show(new WinHwnd());
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
