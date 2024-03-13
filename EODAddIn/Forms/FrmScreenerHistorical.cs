using EODAddIn.BL.Screener;

using System;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class FrmScreenerHistorical : Form
    {
        Screener _screener;
        ScreenerManager _manager;

        public FrmScreenerHistorical(Screener screener, ScreenerManager manager)
        {
            InitializeComponent();
            _screener = screener;
            _manager = manager;
        }

        private async void btnLoad_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(Convert.ToString(period.SelectedItem)) || dateTimePicker2.Value.Subtract(dateTimePicker1.Value).TotalDays < 1 || DateTime.Today.Subtract(dateTimePicker2.Value).TotalDays < -1)
            {
                MessageBox.Show(
                    "incorrect input!",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            var data = await _manager.LoadData(_screener);
            ScreenerPrinter.PrintScreenerHistorical(_screener.NameScreener, data, dateTimePicker1.Value, dateTimePicker2.Value, Convert.ToString(period.SelectedItem));

            Close();
        }
    }
}
