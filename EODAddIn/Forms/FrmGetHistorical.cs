using EODAddIn.BL;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmGetHistorical : Form
    {

        public FrmGetHistorical()
        {
            InitializeComponent();
            cboPeriod.SelectedIndex = 0;
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (!CheckForm()) return; 

            string period = cboPeriod.SelectedItem.ToString().ToLower().Substring(0, 1);
            DateTime from = dtpFrom.Value;
            DateTime to = dtpTo.Value;

            foreach (DataGridViewRow row in gridTickers.Rows)
            {
                if (row.Cells[0].Value == null) continue;
                string ticker = row.Cells[0].Value.ToString();
                List<Model.EndOfDay> res = Utils.APIEOD.GetEOD(ticker, from, to, period);
                LoadToExcel.LoadEndOfDay(res, ticker, period);
            }
            
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
