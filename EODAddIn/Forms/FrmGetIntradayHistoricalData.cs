using EODAddIn.BL;
using EODAddIn.BL.IntradayAPI;
using EODAddIn.BL.IntradayPrinter;
using EODAddIn.BL.Screener;
using EODAddIn.Program;
using EODAddIn.Utils;
using MS.ProgressBar;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Policy;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.Forms
{
    public partial class FrmGetIntradayHistoricalData : Form
    {

        public FrmGetIntradayHistoricalData()
        {
            InitializeComponent();

            /*  5m,1h,1m  */
            switch (Settings.SettingsFields.IntradayInterval)
            {
                case "5m":
                    cboInterval.SelectedIndex = 0;
                    break;
                case "1h":
                    cboInterval.SelectedIndex = 1;
                    break;
                case "1m":
                    cboInterval.SelectedIndex = 2;
                    break;
                default:
                    cboInterval.SelectedIndex = 0;
                    break;
            }

            dtpFrom.Value = Settings.SettingsFields.IntradayFrom;
            dtpTo.Value = DateTime.Now.AddDays(-1);

            foreach (string ticker in Settings.SettingsFields.IntradayTickers)
            {
                int i = gridTickers.Rows.Add();
                gridTickers.Rows[i].Cells[0].Value = ticker;
            }
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return;
            bool isListCreated = false;
            string interval = cboInterval.SelectedItem.ToString().ToLower();
            DateTime from = dtpFrom.Value;
            DateTime to = dtpTo.Value;

            List<string> tikers = new List<string>();
            int rowIntraday = 3;
            Progress progress = new Progress("Load end of data", gridTickers.Rows.Count - 1);
            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                progress.TaskStart(row.Cells[0].Value?.ToString(), 1);

                string ticker = row.Cells[0].Value.ToString();
                tikers.Add(ticker);
                try
                {
                    List<EOD.Model.IntradayHistoricalStockPrice> res = IntradayAPI.GetIntraday(ticker, from, to, interval);
                    if (rbtnAscOrder.Checked)
                    {
                        res.Reverse();
                    }
                    switch (cboTypeOfOutput.SelectedItem.ToString())
                    {
                        case "Separated with chart":
                            rowIntraday = IntradayPrinter.PrintIntraday(res, ticker, interval, true, chkIsTable.Checked);
                            break;
                        case "Separated without chart":
                            rowIntraday=IntradayPrinter.PrintIntraday(res, ticker, interval, false, chkIsTable.Checked);
                            break;
                        case "One worksheet":
                            rowIntraday=IntradayPrinter.PrintIntradaySummary(res, ticker, interval, rowIntraday,isListCreated);
                            isListCreated = true;
                            break;
                    }

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
            if (isListCreated&&chkIsTable.Checked)
            {
                ExcelUtils.MakeTable("A2", "J" + rowIntraday.ToString(), Globals.ThisAddIn.Application.ActiveSheet, "Intraday", 9);
            }
            progress.Finish();
            Settings.SettingsFields.IntradayInterval = interval;
            Settings.SettingsFields.IntradayTo = to;
            Settings.SettingsFields.IntradayFrom = from;
            Settings.SettingsFields.IntradayTickers = tikers;
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

            if (cboInterval.SelectedIndex == -1)
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
            // check from <= to
            int compare = DateTime.Compare(dtpFrom.Value, dtpTo.Value);
            if (compare > 0)
            {
                MessageBox.Show("The \"From\" Date must be not later than the \"To\" Date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void dtpFrom_ValueChanged(object sender, EventArgs e)
        {
            if (!CheckDateInterval(dtpFrom.Value, dtpTo.Value))
            {
                dtpTo.Value = GetPossibleDateTo(dtpFrom.Value);
            }
        }

        private void dtpTo_ValueChanged(object sender, EventArgs e)
        {
            if (!CheckDateInterval(dtpFrom.Value, dtpTo.Value))
            {
                dtpFrom.Value = GetPossibleDateFrom(dtpTo.Value);
            }
        }

        private void cboInterval_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!CheckDateInterval(dtpFrom.Value, dtpTo.Value))
            {
                dtpFrom.Value = GetPossibleDateFrom(dtpTo.Value);
            }
        }

        /// <summary>
        /// Проверка диапазона дат на соответствие стандартам API
        /// </summary>
        /// <returns></returns>
        private bool CheckDateInterval(DateTime from, DateTime to)
        {
            double selectedDateRange = to.Subtract(from).TotalSeconds;
            double possibleDateRange;
            switch (cboInterval.SelectedIndex)
            {
                case 1: //"1h":
                    possibleDateRange = TimeSpan.FromDays(7200).TotalSeconds;
                    break;
                case 2: //"1m":
                    possibleDateRange = TimeSpan.FromDays(120).TotalSeconds;
                    break;
                default: // 0 - "5m"
                    possibleDateRange = TimeSpan.FromDays(600).TotalSeconds;
                    break;
            }

            return  selectedDateRange<=possibleDateRange ;
        }
        private DateTime GetPossibleDateFrom(DateTime dateTo)
        {
            DateTime dateFrom;
            switch (cboInterval.SelectedIndex)
            {
                case 1: //"1h":
                    dateFrom = dateTo.AddSeconds(-TimeSpan.FromDays(7200).TotalSeconds);
                    break;
                case 2: //"1m":
                    dateFrom = dateTo.AddSeconds(-TimeSpan.FromDays(120).TotalSeconds);
                    break;
                default: // 0 - "5m"
                    dateFrom = dateTo.AddSeconds(-TimeSpan.FromDays(600).TotalSeconds);
                    break;
            }
            return dateFrom;
        }
        private DateTime GetPossibleDateTo(DateTime dateFrom)
        {
            DateTime dateTo;
            switch (cboInterval.SelectedIndex)
            {
                case 1: //"1h":
                    dateTo = dateFrom.AddSeconds(TimeSpan.FromDays(7200).TotalSeconds);
                    break;
                case 2: //"1m":
                    dateTo = dateFrom.AddSeconds(TimeSpan.FromDays(120).TotalSeconds);
                    break;
                default: // 0 - "5m"
                    dateTo = dateFrom.AddSeconds(TimeSpan.FromDays(600).TotalSeconds);
                    break;
            }
            return dateTo;
        }
    }
}
