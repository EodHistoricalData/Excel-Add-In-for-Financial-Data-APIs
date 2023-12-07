using EODAddIn.BL.Live;
using EODAddIn.Program;
using EODAddIn.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using static EODAddIn.Utils.ExcelUtils;

namespace EODAddIn.Forms
{
    public partial class FrmGetLive : Form
    {
        private List<(string, bool)> Filters = null;
        internal LiveDownloader LiveDownloader = null;
        public FrmGetLive()
        {
            InitializeComponent();
            Filters = NewFilters();
        }

        private void ClearTicker_Click(object sender, EventArgs e)
        {
            gridTickers.Rows.Clear();
        }

        private void TsmiDeleteRowDataGrid_Click(object sender, EventArgs e)
        {
            if (gridTickers.SelectedRows.Count == 0) return;
            gridTickers.Rows.Remove(gridTickers.SelectedRows[0]);
        }

        private void TsmiFindTicker_Click(object sender, EventArgs e)
        {
            FrmSearchTiker frm = new FrmSearchTiker();
            frm.ShowDialog();

            if (frm.Result.Code == null) return;

            int i = gridTickers.Rows.Add();

            gridTickers.Rows[i].Cells[0].Value = $"{frm.Result.Code}.{frm.Result.Exchange}";
        }

        private void TsmiFromTxt_Click(object sender, EventArgs e)
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

        private void TsmiFromExcel_Click(object sender, EventArgs e)
        {
            FrmSelectRange frm = new FrmSelectRange();
            frm.Show(new WinHwnd());
            frm.FormClosing += FrmSelectRangeClosing;
        }

        private void FrmSelectRangeClosing(object sender, FormClosingEventArgs e)
        {
            FrmSelectRange frm = (FrmSelectRange)sender;
            if (IsRange(frm.RangeAddress))
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
        }

        private void BtnFilters_Click(object sender, EventArgs e)
        {
            FrmLiveFilters frm = new FrmLiveFilters(Filters);
            if (frm.ShowDialog(new WinHwnd()) == DialogResult.OK)
            {
                Filters = frm.Filters;
                if (Filters.All(x => x.Item2 == true))
                {
                    LblFilters.Text = "All";
                }
                else
                {
                    LblFilters.Text = string.Join(", ", Filters.FindAll(x => x.Item2 == true).Select(x => x.Item1));
                }
            }
        }

        private List<(string, bool)> NewFilters()
        {
            List<(string, bool)> pairs = new List<(string, bool)>
            {
                { ("Code", true) },
                { ("Timestamp", true) },
                { ("Gmtoffset", true) },
                { ("Open", true) },
                { ("High", true) },
                { ("Low", true) },
                { ("Close", true) },
                { ("Volume", true) },
                { ("PreviousClose", true) },
                { ("Change", true) },
                { ("Change_p", true) }
            };
            return pairs;
        }

        private void BtnCreate_Click(object sender, EventArgs e)
        {
            if (!CheckForm())
            {
                return;
            }
            int interval = Convert.ToInt32(NudInterval.Value);
            bool smart = chkIsTable.Checked;
            List<(string, string)> tickers = GetTickers();
            int output = cboTypeOfOutput.SelectedIndex;

            int i = 1;
            string downloaderName;
            do
            {
                downloaderName = "Live Downloader " + i;
                if (Settings.SettingsFields.LiveDownloaderNames.Contains(downloaderName))
                {
                    i++;
                }
                else
                {
                    break;
                }
            }
            while (true);

            Settings.SettingsFields.LiveDownloaderNames.Add(downloaderName);
            Settings.Save();
            LiveDownloader = new LiveDownloader(tickers, interval, output, smart, Filters, downloaderName);
            Close();
        }

        private List<(string, string)> GetTickers()
        {
            List<(string, string)> tickers = new List<(string, string)>();
            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                string ticker = row.Cells[0].Value.ToString();
                List<string> codeexchange = ticker.Split('.').ToList();
                if (codeexchange.Count == 2)
                {
                    tickers.Add((codeexchange[0], codeexchange[1]));
                }
            }
            return tickers;
        }

        private bool CheckForm()
        {
            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value != null && !row.Cells[0].Value.ToString().Contains("."))
                {
                    MessageBox.Show("Enter the ticker in the Ticker.Exchange format", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            return true;
        }
    }
}
