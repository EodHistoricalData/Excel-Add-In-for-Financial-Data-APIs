using EODAddIn.BL;
using EODAddIn.Model;
using EODAddIn.Utils;

using Microsoft.Office.Tools.Ribbon;

using System;
using System.Diagnostics;
using System.Windows.Threading;

namespace EODAddIn
{
    public partial class Ribbon
    {

        private DispatcherTimer timer = null;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            timer = new DispatcherTimer();
            timer.Tick += new EventHandler(UpdateRequests);
            timer.Interval = new TimeSpan(0, 0, 0, 20, 0);
            timer.Start();
        }

        private void BtnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            Program.FrmAbout frm = new Program.FrmAbout();
            frm.ShowDialog();
        }

        private void BtnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            Program.FrmAPIKey frm = new Program.FrmAPIKey();
            frm.ShowDialog();
        }

        private void GetHistorical_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetHistorical frm = new Forms.FrmGetHistorical();
            frm.ShowDialog(new WinHwnd());
        }

        private void GetEtf_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetEtf frm = new Forms.FrmGetEtf();
            frm.ShowDialog(new WinHwnd());

            FundamentalData res = frm.Results;
            if (res == null) return;
            LoadToExcel.PrintEtf(res, frm.Tiker);
        }

        private void SplitbtnFundamental_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            Model.FundamentalData res = frm.Results;
            if (res == null) return;
            LoadToExcel.PrintFundamentalAll(res, frm.Tiker);

        }

        private void CheckUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            Program.Program.CheckUpdates();
        }

        private void UpdateRequests(object sender, EventArgs e)
        {
            string key = Program.Settings.SettingsFields.APIKey;
            if (string.IsNullOrEmpty(key))
            {
                lblRequest.Label = "-";
                lblRequestLeft.Label = "-";
                return;
            }

            try
            {
                Model.User user = APIEOD.User(key);

                lblRequest.Label = user.ApiRequests?.ToString("# ##0");
                lblRequestLeft.Label = (user.DailyRateLimit - user.ApiRequests)?.ToString("# ##0");
            }
            catch
            {
                lblRequest.Label = "-";
                lblRequestLeft.Label = "-";
            }
        }

        /// <summary>
        /// Отправка предложения по программе
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SendIdea_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Process.Start("mailto:support@eodhistoricaldata.com" +
                          "?subject=Proposal for Excel AddIn ver. " + Program.Program.Version.Text);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        /// <summary>
        /// Отправка сообщения об ошибке по почте
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ErrorMessage_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Process.Start("mailto:support@eodhistoricaldata.com" +
                                          "?subject=Error in Excel AddIn ver. " + Program.Program.Version.Text);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }

        }

        private void BtnGetGeneral_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            FundamentalData res = frm.Results;
            LoadToExcel.PrintFundamentalGeneral(res);
        }

        private void BtnGetHighlights_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            FundamentalData res = frm.Results;
            LoadToExcel.PrintFundamentalHighlights(res);
        }

        private void BtnGetBalanceSheet_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            FundamentalData res = frm.Results;
            LoadToExcel.PrintFundamentalBalanceSheet(res);
        }

        private void BtnGetIncomeStatement_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            FundamentalData res = frm.Results;
            LoadToExcel.PrintFundamentalIncomeStatement(res);
        }

        private void BtnGetEarnings_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            FundamentalData res = frm.Results;
            LoadToExcel.PrintFundamentalEarnings(res);
        }

        private void BtnGetCashFlow_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            FundamentalData res = frm.Results;
            LoadToExcel.PrintFundamentalCashFlow(res);
        }
        private void BtnFundamentalAllData_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            FundamentalData res = frm.Results;
            LoadToExcel.PrintFundamentalAll(res, frm.Tiker);
        }

        private void BtnGetIntradayHistoricalData_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetIntradayHistoricalData frm = new Forms.FrmGetIntradayHistoricalData();
            frm.ShowDialog(new WinHwnd());
        }

        private async void BtnOptions_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Forms.FrmGetOptions frm = new Forms.FrmGetOptions();
                frm.ShowDialog(new WinHwnd());

                if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    BtnOptions.Label = "Processing";
                    BtnOptions.Enabled = false;

                    EOD.Model.OptionsData.OptionsData res = await GetOptions.GetOptionsData(frm.Ticker, frm.From, frm.To, frm.FromTrade, frm.ToTrade);
                    LoadToExcel.PrintOptions(res, frm.Ticker);

                    BtnOptions.Label = "Get Options";
                    BtnOptions.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void btnCreateScreener_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Forms.FrmScreener frm = new Forms.FrmScreener();
                frm.ShowDialog(new WinHwnd());

                if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    BtnOptions.Label = "Processing";
                    BtnOptions.Enabled = false;


                    BtnOptions.Label = "Get Options";
                    BtnOptions.Enabled = true;
                }

            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }
    }
}
