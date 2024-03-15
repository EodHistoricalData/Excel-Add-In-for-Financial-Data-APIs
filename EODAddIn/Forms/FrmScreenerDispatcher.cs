using EODAddIn.BL.Screener;

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmScreenerDispatcher : Form
    {
        private readonly ScreenerManager _screenerManager;
        public FrmScreenerDispatcher(ScreenerManager screenerManager)
        {
            InitializeComponent();

            _screenerManager = screenerManager;
            UpdateTable();
        }

        private async void NewScreenerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Hide();
            FrmScreener frm = new FrmScreener();
            frm.ShowDialog(this);
            if (frm.DialogResult == DialogResult.OK)
            {
                _screenerManager.Screeners.Add(frm.Screener);
                _screenerManager.Save();
                await _screenerManager.LoadAndPrintScreener(frm.Screener);
                Close();
            }
            else
            {
                Show();
            }

        }

        private void UpdateTable()
        {
            dataGridViewData.Rows.Clear();
            foreach (var item in _screenerManager.Screeners)
            {
                var i = dataGridViewData.Rows.Add();
                dataGridViewData.Rows[i].Cells[0].Value = item.NameScreener;
                dataGridViewData.Rows[i].Tag = item;
            }
        }

        private void DeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = dataGridViewData.Rows.Count - 1; i >= 0; i--)
            {
                if (dataGridViewData.Rows[i].Selected)
                {
                    _screenerManager.Screeners.Remove((Screener)dataGridViewData.Rows[i].Tag);
                    dataGridViewData.Rows.RemoveAt(i);
                }
            }
            _screenerManager.Save();
        }

        private async void GetFundamentalToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridViewData.Rows)
            {
                try
                {
                    if (!row.Selected) continue;
                    Screener screener = (Screener)row.Tag;
                    List<string> tickers = new List<string>();

                    var data = await _screenerManager.LoadData(screener);
                    foreach (var item in data.Data)
                    {
                        tickers.Add($"{item.Code}.{item.Exchange}");
                    }

                    ScreenerPrinter.PrintScreenerBulk(tickers);
                }
                catch (Exception ex)
                {
                    Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                    errorReport.ShowAndSend();
                }
            }
        }

        private void GetHistoricalToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Close();
            foreach (DataGridViewRow row in dataGridViewData.Rows)
            {
                try
                {
                    if (!row.Selected) continue;
                    Screener screener = (Screener)row.Tag;
                    FrmScreenerHistorical frmScreener = new FrmScreenerHistorical(screener, _screenerManager);
                    frmScreener.ShowDialog(this);
                    return;
                }
                catch (Exception ex)
                {
                    Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                    errorReport.ShowAndSend();
                }
            }

            /// 
        }

        private void GetIntradayToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Close();
            foreach (DataGridViewRow row in dataGridViewData.Rows)
            {
                try
                {
                    if (!row.Selected) continue;
                    Screener screener = (Screener)row.Tag;
                    FrmScreenerIntraday frmScreener = new FrmScreenerIntraday(screener, _screenerManager);
                    frmScreener.ShowDialog(this);
                    return;
                }
                catch (Exception ex)
                {
                    Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                    errorReport.ShowAndSend();
                }
            }
        }

        private async void LoadScreenerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
            foreach (DataGridViewRow row in dataGridViewData.Rows)
            {
                if (!row.Selected) continue;
                try
                {
                    Screener screener = (Screener)row.Tag;
                    await _screenerManager.LoadAndPrintScreener(screener);

                }
                catch (Exception ex)
                {
                    Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                    errorReport.ShowAndSend();
                }
            }

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private async void EditToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            foreach (DataGridViewRow row in dataGridViewData.Rows)
            {
                if (!row.Selected) continue;
                try
                {
                    Hide();
                    Screener screener = (Screener)row.Tag;
                    FrmScreener frm = new FrmScreener(screener);
                    frm.ShowDialog(this);
                    _screenerManager.Save();
                    await _screenerManager.LoadAndPrintScreener(screener);
                    UpdateTable();
                    Show();
                    return;
                }
                catch (Exception ex)
                {
                    Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                    errorReport.ShowAndSend();
                }
            }

        }
    }
}
