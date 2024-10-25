using EODAddIn.BL;
using EODAddIn.BL.IntradayAPI;
using EODAddIn.BL.IntradayPrinter;
using EODAddIn.BL.Screener;
using EODAddIn.Program;
using EODAddIn.Utils;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;

using MS.ProgressBar;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.Forms
{
    public partial class FrmGetIntradayHistoricalData : Form
    {

        public FrmGetIntradayHistoricalData()
        {
            InitializeComponent();

            /*  5m,1h,1m  */
            switch (Settings.Data.IntradayInterval)
            {
                case "1m":
                    cboInterval.SelectedIndex = 0;
                    break;
                case "5m":
                    cboInterval.SelectedIndex = 1;
                    break;
                case "15m":
                    cboInterval.SelectedIndex = 2;
                    break;
                case "30m":
                    cboInterval.SelectedIndex = 3;
                    break;
                case "1h":
                    cboInterval.SelectedIndex = 4;
                    break;
                default:
                    cboInterval.SelectedIndex = -1;
                    break;
            }

            dtpFrom.Value = Settings.Data.IntradayFrom;
            dtpTo.Value = DateTime.Now.AddDays(-1);

            foreach (string ticker in Settings.Data.IntradayTickers)
            {
                int i = gridTickers.Rows.Add();
                gridTickers.Rows[i].Cells[0].Value = ticker;
            }
        }

        private async void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return;

            Dictionary<string, string> fails = new Dictionary<string, string>();

            Excel.Worksheet worksheet = null;
            string sheetName;
            bool isSummary = false;
            string intervalString = cboInterval.SelectedItem.ToString().ToLower();
            EOD.API.IntradayHistoricalInterval interval;
            Hide();
            int period = 0;
            switch (intervalString)
            {
                case "15m":
                    {
                        period = 15;
                        interval = EOD.API.IntradayHistoricalInterval.FiveMinutes;
                        break;
                    }
                case "30m":
                    {
                        period = 30;
                        interval = EOD.API.IntradayHistoricalInterval.FiveMinutes;
                        break;
                    }

                case "1m": {
                        interval = EOD.API.IntradayHistoricalInterval.OneMinute; 
                        break;
                    }

                case "1h":
                    {
                        interval = EOD.API.IntradayHistoricalInterval.OneHour;
                        break;
                    }
                default:
                    {
                        interval = EOD.API.IntradayHistoricalInterval.FiveMinutes;
                        break;
                    }
            }
            DateTime from = dtpFrom.Value;
            DateTime to = dtpTo.Value;
            List<string> tikers = new List<string>();
            int rowIntraday = 2;
            Progress progress = new Progress("Get intraday data", gridTickers.Rows.Count - 1);

            if (cboTypeOfOutput.SelectedItem.ToString() == "One worksheet")
            {
                isSummary = true;
                sheetName = ExcelUtils.GetWorksheetNewName("Intraday summary");
                worksheet = ExcelUtils.AddSheet(sheetName);
            }

            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value == null) continue;
               
                progress.TaskStart(row.Cells[0].Value?.ToString() + " loading...", 1);
                string ticker = row.Cells[0].Value.ToString();
                tikers.Add(ticker);
                try
                {
                    List<EOD.Model.IntradayHistoricalStockPrice> res = await IntradayAPI.GetIntraday(ticker, from, to, interval);
                    if (res.Count == 0)
                    {
                        throw new APIException(200, "There is no available data for the selected parameters.");
                    }
                    if (period == 15 || period == 30)
                        res = CollapseRows(res, period);
                    if (rbtnAscOrder.Checked)
                    {
                        res.Reverse();
                    }

                    switch (cboTypeOfOutput.SelectedItem.ToString())
                    {
                        case "Separated with chart":
                            rowIntraday = IntradayPrinter.PrintIntraday(res, ticker, intervalString, true, chkIsTable.Checked);
                            break;
                        case "Separated without chart":
                            rowIntraday = IntradayPrinter.PrintIntraday(res, ticker, intervalString, false, chkIsTable.Checked);
                            break;
                        case "One worksheet":
                            if (gridTickers.Rows.Count > 2)
                            {
                                rowIntraday = IntradayPrinter.PrintIntradaySummary(res, ticker, rowIntraday, worksheet);
                            }
                            else
                            {
                                rowIntraday = IntradayPrinter.PrintIntraday(res, ticker, intervalString, false, chkIsTable.Checked);
                            }
                            break;
                    }
                }
                catch (APIException ex)
                {
                    fails.Add(ticker, ex.StatusError);
                    //MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }
                catch (Exception ex)
                {
                    fails.Add(ticker, ex.Message);
                    ErrorReport error = new ErrorReport(ex);
                    continue;
                }
            }

            if (fails.Count != 0)
            {
                ErrorReport error = new ErrorReport(fails);
                error.ShowAndSend();
            }

            if (isSummary && chkIsTable.Checked && gridTickers.Rows.Count > 2)
            {
                ExcelUtils.MakeTable("A1", "I" + (rowIntraday - 1).ToString(), Globals.ThisAddIn.Application.ActiveSheet, "Intraday", 9);
            }
            progress.Finish();
            Settings.Data.IntradayInterval = intervalString;
            Settings.Data.IntradayTo = to;
            Settings.Data.IntradayFrom = from;
            Settings.Data.IntradayTickers = tikers;
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

        private List<EOD.Model.IntradayHistoricalStockPrice> CollapseRows(List<EOD.Model.IntradayHistoricalStockPrice> res, int v)
        {
            List<EOD.Model.IntradayHistoricalStockPrice> temp = new List<EOD.Model.IntradayHistoricalStockPrice>();
            List<EOD.Model.IntradayHistoricalStockPrice> collapsed = new List<EOD.Model.IntradayHistoricalStockPrice>();
            int count = 0;
            int max = v / 5;
            foreach (EOD.Model.IntradayHistoricalStockPrice row in res)
            {
                temp.Add(row);
                count++;
                if (count == max)
                {
                    double? open = temp[0].Open;
                    double? close = temp[count - 1].Close;
                    double? high = temp.Max(x => x.High);
                    double? low = temp.Min(x => x.Low);
                    decimal? volume = temp.Sum(x => x.Volume);
                    DateTime? date = temp[0].DateTime;
                    long? timestamp = temp[0].Timestamp;
                    double? gmtoffset = temp[0].Gmtoffset;
                    collapsed.Add(new EOD.Model.IntradayHistoricalStockPrice()
                    {
                        Open = open,
                        Close = close,
                        High = high,
                        Low = low,
                        Volume = volume,
                        DateTime = date,
                        Gmtoffset = gmtoffset,
                        Timestamp = timestamp
                    });
                    count = 0;
                    temp.Clear();
                }
            }
            return collapsed;
        }

        private void ClearTicker_Click(object sender, EventArgs e)
        {
            gridTickers.Rows.Clear();
        }

        private void TsmiDeleteRowDataGrid_Click(object sender, EventArgs e)
        {
            if (gridTickers.SelectedRows.Count == 0) return;
            try
            {
                gridTickers.Rows.Remove(gridTickers.SelectedRows[0]);
            }
            catch
            {

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
                case 4:
                    possibleDateRange = TimeSpan.FromDays(7200).TotalSeconds;
                    break;
                case 0:
                    possibleDateRange = TimeSpan.FromDays(120).TotalSeconds;
                    break;
                case -1:
                    return true;
                default:
                    possibleDateRange = TimeSpan.FromDays(600).TotalSeconds;
                    break;
            }

            return selectedDateRange <= possibleDateRange;
        }
        private DateTime GetPossibleDateFrom(DateTime dateTo)
        {
            DateTime dateFrom;
            switch (cboInterval.SelectedIndex)
            {
                case 4:
                    dateFrom = dateTo.AddSeconds(-TimeSpan.FromDays(7200).TotalSeconds);
                    break;
                case 0:
                    dateFrom = dateTo.AddSeconds(-TimeSpan.FromDays(120).TotalSeconds);
                    break;
                default:
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
                case 4:
                    dateTo = dateFrom.AddSeconds(TimeSpan.FromDays(7200).TotalSeconds);
                    break;
                case 0:
                    dateTo = dateFrom.AddSeconds(TimeSpan.FromDays(120).TotalSeconds);
                    break;
                default:
                    dateTo = dateFrom.AddSeconds(TimeSpan.FromDays(600).TotalSeconds);
                    break;
            }
            return dateTo;
        }
    }
}
