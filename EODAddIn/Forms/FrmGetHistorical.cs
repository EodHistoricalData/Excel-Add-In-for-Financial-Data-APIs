using EODAddIn.BL;
using EODAddIn.Program;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MS.ProgressBar;
using EODAddIn.Utils;

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
            if (Settings.SettingsFields.EndOfDayTo != DateTime.MinValue)
            {
                dtpTo.Value = Settings.SettingsFields.EndOfDayTo;
            }
            
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

            Progress progress = new Progress("Load end of data", gridTickers.Rows.Count-1);
            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                progress.TaskStart(row.Cells[0].Value?.ToString(), 1);
                
                string ticker = row.Cells[0].Value.ToString();
                try
                {
                    List<Model.EndOfDay> res = APIEOD.GetEOD(ticker, from, to, period);
                    LoadToExcel.LoadEndOfDay(res, ticker, period);
                }
                catch (APIException ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                }
                catch (Exception ex)
                {
                    ErrorReport error = new ErrorReport(ex);
                    error.ShowAndSend();
                    break;
                }
                
                tikers.Add(ticker);
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
                    MessageBox.Show("Enter the ticket in the Ticket.Exchange format", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

    }
}
