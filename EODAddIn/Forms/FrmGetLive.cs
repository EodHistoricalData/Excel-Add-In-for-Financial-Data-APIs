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
using EODAddIn.BL;

namespace EODAddIn.Forms
{
    public partial class FrmGetLive : Form
    {
        private List<Filter> Filters = null;
        internal LiveDownloader LiveDownloader = null;
        public FrmGetLive()
        {
            InitializeComponent();
            Filters = NewFilters();
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
            tsmiFromExcel.Enabled = false;
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
            tsmiFromExcel.Enabled = true;
        }

        private void BtnFilters_Click(object sender, EventArgs e)
        {
            FrmLiveFilters frm = new FrmLiveFilters(Filters);
            if (frm.ShowDialog(new WinHwnd()) == DialogResult.OK)
            {
                Filters = frm.Filters;
                if (Filters.All(x => x.IsChecked == true))
                {
                    LblFilters.Text = "All";
                }
                else
                {
                    LblFilters.Text = string.Join(", ", Filters.FindAll(x => x.IsChecked == true).Select(x => x.Name));
                }
            }
        }

        private List<Filter> NewFilters()
        {
            List<(string, bool)> pairs = new List<(string, bool)>
            {
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

            var filters = new List<Filter>();

            foreach (var pair in pairs)
            {
                var filter = new Filter()
                {
                    Name = pair.Item1,
                    IsChecked = pair.Item2,
                };
                filters.Add(filter);
            }

            return filters;
        }

        private void BtnCreate_Click(object sender, EventArgs e)
        {
            if (!CheckForm())
            {
                return;
            }
            try
            {
                int interval = Convert.ToInt32(NudInterval.Value);
                bool smart = chkIsTable.Checked;
                List<Ticker> tickers = GetTickers();

                int i = 1;
                string downloaderName;
                do
                {
                    downloaderName = "Live Downloader " + i;
                    var count = LiveDownloaderManager.LiveDownloaders.Where(x => x.Name == downloaderName).Count();
                    if (count == 0) break;
                    i++;
                }
                while (true);

                LiveDownloader = new LiveDownloader(tickers, interval, smart, Filters, downloaderName, Globals.ThisAddIn.Application.ActiveWorkbook);

                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<Ticker> GetTickers()
        {
            List<Ticker> tickers = new List<Ticker>();
            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                string ticker = row.Cells[0].Value.ToString();
                List<string> codeexchange = ticker.Split('.').ToList();
                if (codeexchange.Count == 2)
                {
                    var newTicker = new Ticker()
                    {
                        Name = codeexchange[0],
                        Exchange = codeexchange[1]
                    };
                    tickers.Add(newTicker);
                }
            }
            return tickers;
        }

        private bool CheckForm()
        {
            if (gridTickers.Rows.Count == 0)
            {
                MessageBox.Show("Enter the ticker in the Ticker.Exchange format", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
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

        private void tsmiClearTicker_Click(object sender, EventArgs e)
        {
            gridTickers.Rows.Clear();
        }
    }
}
