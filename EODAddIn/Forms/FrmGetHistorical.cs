using EODAddIn.BL;
using EODAddIn.Program;
using EODAddIn.Utils;

using MS.ProgressBar;

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.Forms
{
    public partial class FrmGetHistorical : Form
    {

        public FrmGetHistorical()
        {
            InitializeComponent();

            switch (Settings.SettingsFields.EndOfDayPeriod)
            {
                case "d":
                    cboPeriod.SelectedIndex = 0;
                    break;
                case "w":
                    cboPeriod.SelectedIndex = 1;
                    break;
                case "m":
                    cboPeriod.SelectedIndex = 2;
                    break;
                default:
                    cboPeriod.SelectedIndex = 0;
                    break;
            }

            dtpFrom.Value = Settings.SettingsFields.EndOfDayFrom;
            dtpTo.Value = DateTime.Today.AddDays(-1);


            foreach (string ticker in Settings.SettingsFields.EndOfDayTickers)
            {
                int i = gridTickers.Rows.Add();
                gridTickers.Rows[i].Cells[0].Value = ticker;
            }
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return;

            string period = cboPeriod.SelectedItem.ToString().ToLower().Substring(0, 1);
            DateTime from = dtpFrom.Value;
            DateTime to = dtpTo.Value;

            List<string> tikers = new List<string>();

            Progress progress = new Progress("Load end of data", gridTickers.Rows.Count - 1);
            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                progress.TaskStart(row.Cells[0].Value?.ToString(), 1);

                string ticker = row.Cells[0].Value.ToString();
                tikers.Add(ticker);
                try
                {
                    List<Model.EndOfDay> res = APIEOD.GetEOD(ticker, from, to, period);
                    LoadToExcel.LoadEndOfDay(res, ticker, period);
                }
                catch (APIException ex)
                {
                    MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }
                catch (Exception ex)
                {
                    ErrorReport error = new ErrorReport(ex);
                    error.ShowAndSend();
                    continue;
                }


            }
            progress.Finish();
            Settings.SettingsFields.EndOfDayPeriod = period;
            Settings.SettingsFields.EndOfDayTo = to;
            Settings.SettingsFields.EndOfDayFrom = from;
            Settings.SettingsFields.EndOfDayTickers = tikers;
            Settings.Save();

            Close();
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
        }
    }
}
