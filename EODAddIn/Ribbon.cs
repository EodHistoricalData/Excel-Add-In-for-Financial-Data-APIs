using EODAddIn.BL;
using EODAddIn.BL.BulkFundamental;
using EODAddIn.BL.ETFPrinter;
using EODAddIn.BL.FundamentalDataPrinter;
using EODAddIn.BL.OptionsAPI;
using EODAddIn.BL.OptionsPrinter;
using EODAddIn.BL.Screener;
using EODAddIn.Utils;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Threading;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

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

            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            if (res == null) return;
            ETFPrinter.PrintEtf(res, frm.Tiker);
        }

        private void SplitbtnFundamental_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            if (res == null) return;
            FundamentalDataPrinter.PrintFundamentalAll(res, frm.Tiker);

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
                EOD.Model.User user = JsonConvert.DeserializeObject<EOD.Model.User>(Response.GET("https://eodhistoricaldata.com/api/user", "api_token=" + key));
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
            if (frm.DialogResult != DialogResult.OK)
            {
                return;
            }
            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            FundamentalDataPrinter.PrintFundamentalGeneral(res);
        }

        private void BtnGetHighlights_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());
            if (frm.DialogResult != DialogResult.OK)
            {
                return;
            }
            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            FundamentalDataPrinter.PrintFundamentalHighlights(res);
        }

        private void BtnGetBalanceSheet_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());
            if (frm.DialogResult != DialogResult.OK)
            {
                return;
            }
            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            FundamentalDataPrinter.PrintFundamentalBalanceSheet(res);
        }

        private void BtnGetIncomeStatement_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());
            if (frm.DialogResult != DialogResult.OK)
            {
                return;
            }
            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            FundamentalDataPrinter.PrintFundamentalIncomeStatement(res);
        }

        private void BtnGetEarnings_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());
            if (frm.DialogResult != DialogResult.OK)
            {
                return;
            }
            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            FundamentalDataPrinter.PrintFundamentalEarnings(res);
        }

        private void BtnGetCashFlow_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());
            if (frm.DialogResult != DialogResult.OK)
            {
                return;
            }
            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            FundamentalDataPrinter.PrintFundamentalCashFlow(res);
        }
        private void BtnFundamentalAllData_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());
            if (frm.DialogResult!= DialogResult.OK)
            {
                return;
            }
            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            FundamentalDataPrinter.PrintFundamentalAll(res, frm.Tiker);
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

                    EOD.Model.OptionsData.OptionsData res = await OptionsAPI.GetOptionsData(frm.Ticker, frm.From, frm.To, frm.FromTrade, frm.ToTrade);
                    OptionsPrinter.PrintOptions(res, frm.Ticker);
                }
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
            finally
            {
                BtnOptions.Label = "Get Options";
                BtnOptions.Enabled = true;
            }
        }

        private void BtnGetBulk_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Forms.FrmGetBulk frm = new Forms.FrmGetBulk();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult ==DialogResult.OK)
                {
                    BtnGetBulk.Label = "Processing";
                    BtnGetBulk.Enabled = false;
                    BulkFundamentalPrinter.PrintBulkFundamentals(frm.Tickers);
                }
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
            finally
            {
                BtnGetBulk.Label = "Bulk Fundamentals";
                BtnGetBulk.Enabled = true;
            }
        }

        private async void btnCreateScreener_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Forms.FrmScreener frm = new Forms.FrmScreener();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult ==DialogResult.OK)
                {
                    BtnOptions.Label = "Processing";
                    BtnOptions.Enabled = false;

                    var res = await ScreenerAPI.GetScreener(frm.Filters, frm.Signals, frm.Sort, frm.Limit);
                    ScreneerPrinter.PrintScreener(res);

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

        private void btnGetScreenerFundamenat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ScreneerPrinter.PrintScreenerBulk();
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void btnGetSreenerHistorical_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Forms.FrmScreenerHistorical frm = new Forms.FrmScreenerHistorical();
                frm.ShowDialog(new WinHwnd());
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Forms.FrmScreenerIntraday frm = new Forms.FrmScreenerIntraday();
                frm.ShowDialog(new WinHwnd());
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }
    }
}
