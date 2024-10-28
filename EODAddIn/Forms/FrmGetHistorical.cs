using EOD.Model;
using EODAddIn.BL;
using EODAddIn.BL.HistoricalAPI;
using EODAddIn.BL.HistoricalPrinter;
using EODAddIn.Program;
using EODAddIn.Utils;

using Microsoft.Office.Interop.Excel;

using MS.ProgressBar;

using System;
using System.Collections.Generic;
using System.Diagnostics;
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

            switch (Settings.Data.GetHistoricalForm.Period)
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

            dtpFrom.Value = Settings.Data.GetHistoricalForm.From;
            dtpTo.Value = DateTime.Today.AddDays(-1);

            if (Settings.Data.GetHistoricalForm.OrderDesc) rbtnDescOrder.Checked = true;
            chkIsTable.Checked = Settings.Data.GetHistoricalForm.SmartTable;
            chkDateToWS.Checked = Settings.Data.GetHistoricalForm.AddDate;

            switch (Settings.Data.GetHistoricalForm.TypeOfOutput)
            {
                case "Separated with chart":
                    cboTypeOfOutput.SelectedIndex = 1;
                    break;
                case "Separated without chart":
                    cboTypeOfOutput.SelectedIndex = 2;
                    break;
                case "One worksheet":
                    cboTypeOfOutput.SelectedIndex = 0;
                    break;
            }

            foreach (string ticker in Settings.Data.GetHistoricalForm.Tickers)
            {
                int i = gridTickers.Rows.Add();
                gridTickers.Rows[i].Cells[0].Value = ticker;
            }
        }

        private async void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return;

            Dictionary<string, string> fails = new Dictionary<string, string>();

            string sheetName;
            bool isSummary = false;
            Worksheet worksheet = null;
            string period = cboPeriod.SelectedItem.ToString().ToLower().Substring(0, 1);
            DateTime from = dtpFrom.Value;
            DateTime to = dtpTo.Value;
            List<string> tikers = new List<string>();
            int rowHistorical = 2;
            Progress progress = new Progress("Load end of data", gridTickers.Rows.Count - 1);

            if (cboTypeOfOutput.SelectedItem.ToString() == "One worksheet")
            {
                isSummary = true;
                sheetName = ExcelUtils.GetWorksheetNewName("End of day summary");
                worksheet = ExcelUtils.AddSheet(sheetName);
            }

            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                progress.TaskStart(row.Cells[0].Value?.ToString(), 1);
                string ticker = row.Cells[0].Value.ToString();

                if (tikers.Contains(ticker))
                    continue;

                tikers.Add(ticker);
                try
                {
                    List<HistoricalStockPrice> res = await HistoricalAPI.GetHistoricalStockPrice(ticker, from, to, period);
                    if (rbtnAscOrder.Checked)
                    {
                        res.Reverse();
                    }
                    switch (cboTypeOfOutput.SelectedItem.ToString())
                    {
                        case "Separated with chart":
                            rowHistorical = HistoricalPrinter.PrintEndOfDay(res, ticker, period, true, chkIsTable.Checked, chkDateToWS.Checked);
                            break;
                        case "Separated without chart":
                            rowHistorical = HistoricalPrinter.PrintEndOfDay(res, ticker, period, false, chkIsTable.Checked, chkDateToWS.Checked);
                            break;
                        case "One worksheet":
                            if (gridTickers.Rows.Count > 1)
                            {
                                rowHistorical = HistoricalPrinter.PrintEndOfDaySummary(res, ticker, period, rowHistorical, worksheet);
                            }
                            else
                            {
                                rowHistorical = HistoricalPrinter.PrintEndOfDay(res, ticker, period, false, chkIsTable.Checked, chkDateToWS.Checked);
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

            if (isSummary && chkIsTable.Checked)
            {
                ExcelUtils.MakeTable("A1", "K" + (rowHistorical - 1).ToString(), worksheet, "Historical", 9);
            }
            progress.Finish();
            FormSettingsSave(tikers);



            Close();
        }

        private void FormSettingsSave(List<string> tikers)
        {
            Settings.Data.GetHistoricalForm.Period = cboPeriod.SelectedItem.ToString().ToLower().Substring(0, 1);
            Settings.Data.GetHistoricalForm.To = dtpTo.Value;
            Settings.Data.GetHistoricalForm.From = dtpFrom.Value;
            Settings.Data.GetHistoricalForm.OrderDesc = rbtnDescOrder.Checked;
            Settings.Data.GetHistoricalForm.SmartTable = chkIsTable.Checked;
            Settings.Data.GetHistoricalForm.AddDate = chkDateToWS.Checked;
            Settings.Data.GetHistoricalForm.Tickers = tikers;
            Settings.Data.GetHistoricalForm.TypeOfOutput = cboTypeOfOutput.SelectedItem.ToString();
            Settings.Save();
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
