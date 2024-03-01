using EOD.Model.OptionsData;

using EODAddIn.BL.HistoricalPrinter;
using EODAddIn.BL.TechnicalIndicatorData;
using EODAddIn.Program;
using EODAddIn.Utils;
using EODHistoricalData.Wrapper.Model.TechnicalIndicators;

using Microsoft.Office.Interop.Excel;

using MS.ProgressBar;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static EOD.API;
using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.Forms
{
    public partial class FrmGetTechnicals : Form
    {
        readonly Dictionary<int, string> Functions = new Dictionary<int, string>()
        {
            {0, "avgvol"},
            {1, "avgvolccy"},
            {2, "sma"},
            {3, "ema"},
            {4, "wma"},
            {5, "volatility"},
            {6, "rsi"},
            {7, "stddev"},
            {8, "slope"},
            {9, "dmi"},
            {10, "adx"},
            {11, "atr"},
            {12, "cci"},
            {13, "bbands"},
            {14, "splitadjusted"},
            {15, "stochastic"},
            {16, "stochrsi"},
            {17, "macd"},
            {18, "sar"}
        };

        public FrmGetTechnicals()
        {
            InitializeComponent();

            cboFunction.SelectedIndex = Settings.Data.TechnicalsFunctionId;
            dtpFrom.Value = Settings.Data.TechnicalsFrom != DateTime.MinValue ? Settings.Data.TechnicalsFrom : new DateTime(2020, 1, 1);
            dtpTo.Value =  DateTime.Today;

            foreach (string ticker in Settings.Data.TechnicalsTickers)
            {
                int i = gridTickers.Rows.Add();
                gridTickers.Rows[i].Cells[0].Value = ticker;
            }
        }

        private void CboFunction_SelectedIndexChanged(object sender, EventArgs e)
        {
            var select = (ComboBox)sender;

            if (select.SelectedIndex < 14)
            {
                labelFirstOption.Text = "Period";
                labelFirstOption.Visible = true;
                tbFirstOption.Visible = true;

                labelSecondOption.Visible = false;
                labelThirdOption.Visible = false;
                cboAggPeriod.Visible = false;
                tbSecondOption.Visible = false;
                tbThirdOption.Visible = false;
            }
            if (select.SelectedIndex == 14)
            {
                labelFirstOption.Text = "Aggregation period";
                labelFirstOption.Visible = true;
                cboAggPeriod.Visible = true;

                labelSecondOption.Visible = false;
                labelThirdOption.Visible = false;
                tbFirstOption.Visible = false;
                tbSecondOption.Visible = false;
                tbThirdOption.Visible = false;
            }
            if (select.SelectedIndex == 15)
            {
                labelFirstOption.Text = "Fast K-period";
                labelSecondOption.Text = "Slow K-period";
                labelThirdOption.Text = "Slow D-period";
                labelFirstOption.Visible = true;
                labelSecondOption.Visible = true;
                labelThirdOption.Visible = true;
                tbFirstOption.Visible = true;
                tbSecondOption.Visible = true;
                tbThirdOption.Visible = true;

                cboAggPeriod.Visible = false;
            }
            if (select.SelectedIndex == 16)
            {
                labelFirstOption.Text = "Fast K-period";
                labelSecondOption.Text = "Fast D-period";
                labelFirstOption.Visible = true;
                labelSecondOption.Visible = true;
                tbFirstOption.Visible = true;
                tbSecondOption.Visible = true;

                cboAggPeriod.Visible = false;
                labelThirdOption.Visible = false;
                tbThirdOption.Visible = false;
            }
            if (select.SelectedIndex == 17)
            {
                labelFirstOption.Text = "Fast period";
                labelSecondOption.Text = "Slow period";
                labelThirdOption.Text = "Signal D-period";
                labelFirstOption.Visible = true;
                labelSecondOption.Visible = true;
                labelThirdOption.Visible = true;
                tbFirstOption.Visible = true;
                tbSecondOption.Visible = true;
                tbThirdOption.Visible = true;

                cboAggPeriod.Visible = false;
            }
            if (select.SelectedIndex == 18)
            {
                labelFirstOption.Text = "Acceleration Factor";
                labelSecondOption.Text = "Acceleration Factor Maximum value";
                labelFirstOption.Visible = true;
                labelSecondOption.Visible = true;
                tbFirstOption.Visible = true;
                tbSecondOption.Visible = true;

                cboAggPeriod.Visible = false;
                labelThirdOption.Visible = false;
                tbThirdOption.Visible = false;
            }
        }

        private async void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return;
            string sheetName;
            bool isSummary = false;
            Worksheet worksheet = null;
            int row = 1;
            Order order = rbtnAscOrder.Checked ? Order.Ascending : Order.Descending;
            DateTime from = dtpFrom.Value;
            DateTime to = dtpTo.Value;

            Settings.Data.TechnicalsFunctionId = cboFunction.SelectedIndex;
            Settings.Data.TechnicalsTo = to;
            Settings.Data.TechnicalsFrom = from;
            Settings.Save();

            btnLoad.Enabled = false;
            List<IndicatorParameters> parameters = GetParameters();
            List<string> tikers = new List<string>();
            Progress progress = new Progress("Loading data", gridTickers.Rows.Count - 1);

            if (cboTypeOfOutput.SelectedItem.ToString() == "One worksheet")
            {
                isSummary = true;
                string function = parameters.First(x => x.Name == "function").Value;
                sheetName = ExcelUtils.GetWorksheetNewName($"Technical summary - {function}");
                worksheet = ExcelUtils.AddSheet(sheetName);
            }

            foreach (DataGridViewRow item in gridTickers.Rows)
            {            
                if (item.Cells[0].Value == null) continue;
                progress.TaskStart(item.Cells[0].Value?.ToString(), 1);
                string ticker = item.Cells[0].Value.ToString();
                tikers.Add(ticker);
                try
                {
                    var result = await TechnicalIndicatorAPI.GetTechnicalIndicatorsData(ticker, from, to, order, parameters);

                    if (result.Count == 0)
                    {
                        MessageBox.Show("There is no available data for the selected parameters.", "No data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (rbtnAscOrder.Checked)
                    {
                        result.Reverse();
                    }
                    switch (cboTypeOfOutput.SelectedItem.ToString())
                    {
                        case "Separated with chart":
                            row = TechnicalsPrinter.PrintTechnicals(result, ticker, parameters, true, chkIsTable.Checked);
                            break;
                        case "Separated without chart":
                            row = TechnicalsPrinter.PrintTechnicals(result, ticker, parameters, false, chkIsTable.Checked);
                            break;
                        case "One worksheet":
                            row = TechnicalsPrinter.PrintTechnicalsSummary(result, ticker, row, parameters, worksheet);
                            isSummary = true;
                            break;
                    }
                }
                catch (APIException ex)
                {
                    MessageBox.Show(ex.StatusError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }
                catch (Exception ex)
                {
                    ErrorReport error = new ErrorReport(ex);
                    error.ShowAndSend();
                    Close();
                    return;
                }
            }

            if (isSummary && chkIsTable.Checked)
            {
                ExcelUtils.MakeTable("A2", "K" + row.ToString(), Globals.ThisAddIn.Application.ActiveSheet, "Historical", 9);
            }

            progress.Finish();
            btnLoad.Enabled = true;
            Settings.Data.TechnicalsTickers = tikers;
            Settings.Save();
            DialogResult = DialogResult.OK;
            Close();
        }

        private List<IndicatorParameters> GetParameters()
        {
            List<IndicatorParameters> parameters = new List<IndicatorParameters>();
            string function = Functions[cboFunction.SelectedIndex];
            parameters.Add(new IndicatorParameters("function", function));
            if (labelFirstOption.Visible)
            {
                if (!string.IsNullOrEmpty(tbFirstOption.Text))
                    parameters.Add(new IndicatorParameters(labelFirstOption.Text.ToLower(), tbFirstOption.Text.ToLower()));
            }
            if (labelSecondOption.Visible)
            {
                if (!string.IsNullOrEmpty(tbSecondOption.Text))
                    parameters.Add(new IndicatorParameters(labelSecondOption.Text.ToLower(), tbSecondOption.Text.ToLower()));
            }
            if (labelThirdOption.Visible)
            {
                if (!string.IsNullOrEmpty(tbThirdOption.Text))
                    parameters.Add(new IndicatorParameters(labelThirdOption.Text.ToLower(), tbThirdOption.Text.ToLower()));
            }
            return parameters;
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

            if (cboFunction.SelectedIndex == -1)
            {
                MessageBox.Show("Select function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void tsmiFindTicker_Click(object sender, EventArgs e)
        {
            FrmSearchTiker frm = new FrmSearchTiker();
            frm.ShowDialog();

            if (frm.Result.Code == null) return;

            int i = gridTickers.Rows.Add();

            gridTickers.Rows[i].Cells[0].Value = $"{frm.Result.Code}.{frm.Result.Exchange}";
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
