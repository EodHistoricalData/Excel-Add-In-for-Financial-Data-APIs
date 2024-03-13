using EODAddIn.BL;
using EODAddIn.BL.IntradayPrinter;
using EODAddIn.BL.Screener;
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
    public partial class FrmScreenerIntraday : Form
    {
        Screener _screener;
        ScreenerManager _manager;
        public FrmScreenerIntraday(Screener screener, ScreenerManager manager)
        {
            InitializeComponent();
            _screener = screener;
            _manager = manager;
        }

        private async void btnScreenLoadIntraday_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(Convert.ToString(comboBox1.SelectedItem)) || dateTimePicker2.Value.Subtract(dateTimePicker1.Value).TotalDays < 1 || DateTime.Today.Subtract(dateTimePicker2.Value).TotalDays < -1)
            {
                MessageBox.Show(
                    "incorrect input!",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            EOD.API.IntradayHistoricalInterval interval;
            switch (Convert.ToString(comboBox1.SelectedItem))
            {
                case "1m":
                    interval = EOD.API.IntradayHistoricalInterval.OneMinute; break;
                case "5m":
                    interval = EOD.API.IntradayHistoricalInterval.FiveMinutes; break;
                case "1h":
                    interval = EOD.API.IntradayHistoricalInterval.OneHour; break;
                default:
                    interval = EOD.API.IntradayHistoricalInterval.FiveMinutes;
                    break;
            }

            var data = await _manager.LoadData(_screener);
            ScreenerPrinter.PrintScreenerIntraday(_screener.NameScreener, data, dateTimePicker1.Value, dateTimePicker2.Value, interval);

            Close();
        }
    }
}
