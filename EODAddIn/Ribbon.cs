using EOD.Model.Bulks;

using EODAddIn.BL.BulkEod;
using EODAddIn.BL.BulkFundamental;
using EODAddIn.BL.ETFPrinter;
using EODAddIn.BL.FundamentalDataPrinter;
using EODAddIn.BL.Live;
using EODAddIn.BL.OptionsAPI;
using EODAddIn.BL.OptionsPrinter;
using EODAddIn.BL.Screener;
using EODAddIn.Forms;
using EODAddIn.Utils;
using EODAddIn.View.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Threading;
using System.Xml.Serialization;

using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn
{
    public partial class Ribbon
    {
        private Excel.Application _xlapp;
        private DispatcherTimer timer = null;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            timer = new DispatcherTimer();
            timer.Tick += new EventHandler(UpdateRequests);
            timer.Interval = new TimeSpan(0, 0, 10, 0, 0);
            timer.Start();

            _xlapp = Globals.ThisAddIn.Application;
            _xlapp.WorkbookOpen += OnWorkbookOpen;
            _xlapp.WorkbookBeforeClose += OnWorkbookBeforeClose;
        }

        

        private void BtnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            Program.FrmAbout frm = new Program.FrmAbout();
            frm.ShowDialog();
        }

        private void BtnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var panels = Globals.ThisAddIn.CustomTaskPanes;
                var panel = panels.FirstOrDefault(p => p.Title == "Info");
                if (panel == null)
                {
                    var panelInfo = new Panels.PanelInfo();
                    panelInfo.ShowPanel();
                }
                else
                {
                    panel.Visible = panel.Visible ? false : true;
                }
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }

        }

        private void GetHistorical_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                FormShower.FrmGetHistoricalShow();

            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void GetEtf_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;

                FrmGetEtf frm = new FrmGetEtf();
                frm.ShowDialog(new WinHwnd());

                EOD.Model.Fundamental.FundamentalData res = frm.Results;
                if (res == null) return;
                ETFPrinter.PrintEtf(res, frm.Tiker);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }

        }

        private void SplitbtnFundamental_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
                frm.ShowDialog(new WinHwnd());

                EOD.Model.Fundamental.FundamentalData res = frm.Results;
                if (res == null) return;
                FundamentalDataPrinter.PrintFundamentalAll(res, frm.Tiker);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void CheckUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Program.Program.CheckUpdates();

            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void UpdateRequests(object sender, EventArgs e)
        {
            try
            {
                string key = Program.Settings.Data.APIKey;
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
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
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
            try
            {
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult != DialogResult.OK)
                {
                    return;
                }
                EOD.Model.Fundamental.FundamentalData res = frm.Results;
                FundamentalDataPrinter.PrintFundamentalGeneral(res);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }

        }

        private void BtnGetHighlights_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult != DialogResult.OK)
                {
                    return;
                }
                EOD.Model.Fundamental.FundamentalData res = frm.Results;
                FundamentalDataPrinter.PrintFundamentalHighlights(res);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }

        }

        private void BtnGetBalanceSheet_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult != DialogResult.OK)
                {
                    return;
                }
                EOD.Model.Fundamental.FundamentalData res = frm.Results;
                FundamentalDataPrinter.PrintFundamentalBalanceSheet(res);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void BtnGetIncomeStatement_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult != DialogResult.OK)
                {
                    return;
                }
                EOD.Model.Fundamental.FundamentalData res = frm.Results;
                FundamentalDataPrinter.PrintFundamentalIncomeStatement(res);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void BtnGetEarnings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult != DialogResult.OK)
                {
                    return;
                }
                EOD.Model.Fundamental.FundamentalData res = frm.Results;
                FundamentalDataPrinter.PrintFundamentalEarnings(res);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void BtnGetCashFlow_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult != DialogResult.OK)
                {
                    return;
                }
                EOD.Model.Fundamental.FundamentalData res = frm.Results;
                FundamentalDataPrinter.PrintFundamentalCashFlow(res);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }
        private void BtnFundamentalAllData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult != DialogResult.OK)
                {
                    return;
                }
                EOD.Model.Fundamental.FundamentalData res = frm.Results;
                FundamentalDataPrinter.PrintFundamentalAll(res, frm.Tiker);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void BtnGetIntradayHistoricalData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                FormShower.FrmGetIntradayHistoricalDataShow();

            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }


        private async void BtnOptions_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;
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
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetBulk frm = new Forms.FrmGetBulk();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult == DialogResult.OK)
                {
                    BtnGetBulk.Label = "Processing";
                    BtnGetBulk.Enabled = false;
                    switch (frm.BulkTypeOfOutput)
                    {
                        case "Separated":
                            BulkFundamentalPrinter.PrintBulkFundamentals(frm.Tickers);
                            break;
                        case "One worksheet":
                            ScreenerPrinter.PrintScreenerBulk(frm.Tickers);
                            break;
                    }
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

       

        private void BtnListOfExchanges_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhd.com/financial-apis/list-supported-exchanges/?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");

        }

        private void BtnListOfCRYPTOCurrencies_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhd.com/financial-apis/list-supported-crypto-currencies/?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");

        }

        private void BtnListOfFutures_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhd.com/financial-apis/list-supported-futures-commodities/?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");

        }

        private void BtnListOfForexCurrencies_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhd.com/financial-apis/list-supported-forex-currencies/?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");

        }



        private async void BtnBulkEod_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (FormShower.ShowActiveForm()) return;
                Forms.FrmGetBulkEod frm = new Forms.FrmGetBulkEod();
                frm.ShowDialog(new WinHwnd());

                if (frm.DialogResult == DialogResult.OK)
                {
                    string exchange = frm.Exchange;
                    string type = "end-of-day data";

                    DateTime date = frm.Date;
                    string tickers = string.Join(",", frm.Tickers);
                    BtnBulkEod.Label = "Processing";
                    BtnBulkEod.Enabled = false;

                    List<Bulk> res = await GetBulkEod.GetBulkEodData(exchange, EODHistoricalData.Wrapper.Model.Bulks.BulkQueryTypes.EndOfDay, date, tickers);
                    BulkEodPrinter.PrintBulkEod(res, exchange, date, tickers, type);
                }
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
            finally
            {
                BtnBulkEod.Label = "Get Bulk Eod";
                BtnBulkEod.Enabled = true;
            }
        }

        private void BtnTechnicals_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                FormShower.FrmGetTechnicalsShow();
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void BtnGetLive_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                LiveDownloaderDispatcher frm = FormShower.LiveDownloaderDispatcherShow();
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void OnWorkbookOpen(Excel.Workbook Wb)
        {
            LiveDownloaderManager.LoadWorkbook(Wb);

        }


        private void OnWorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            LiveDownloaderManager.CloseWorkbook(Wb);
        }

        private async void BtnScreener_Click(object sender, RibbonControlEventArgs e)
        {
            ScreenerManager manager = new ScreenerManager();
            await manager.AddNewScreener();
        }
    }
}
