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

using Microsoft.Office.Tools;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Threading;

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

                var activeWindow = _xlapp.ActiveWindow;
                var panels = Globals.ThisAddIn.CustomTaskPanes;
                var infoPanels = panels.Where(p => p.Title == "Info").ToList();
                if (infoPanels.Count == 0)
                {
                    var panelInfo = new Panels.PanelInfo();
                    panelInfo.ShowPanel();
                }
                else
                {
                    CustomTaskPane taskPane = null;
                    foreach (var panel in infoPanels)
                    {
                        var panelWindow = (Window)panel.Window;
                        if (panelWindow.Hwnd == activeWindow.Hwnd)
                        {
                            taskPane = panel;
                        }
                    }
                    if (taskPane == null)
                    {
                        var panelInfo = new Panels.PanelInfo();
                        panelInfo.ShowPanel();
                    }
                    else
                    {
                        taskPane.Visible = !taskPane.Visible;
                    }
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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

                FormShower.FrmGetBulkShow();

            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

                if (FormShower.ShowActiveForm()) return;
                FrmGetBulkEod frm = new FrmGetBulkEod();
                frm.Show(new WinHwnd());

                if (frm.DialogResult == DialogResult.OK)
                {
                    string exchange = string.IsNullOrEmpty(frm.Exchange) ? "US" : frm.Exchange;
                    string type = "end-of-day data";

                    DateTime date = frm.Date;
                    string tickers = string.Join(",", frm.Tickers);
                    BtnBulkEod.Label = "Processing";
                    BtnBulkEod.Enabled = false;

                    List<Bulk> res = await GetBulkEod.GetBulkEodData(exchange, EODHistoricalData.Wrapper.Model.Bulks.BulkQueryTypes.EndOfDay, date, tickers);
                    if (!frm.IsExchange)
                    {
                        exchange = "Tickers";
                    }
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
                BtnBulkEod.Label = "Get Bulk EOD";
                BtnBulkEod.Enabled = true;
            }
        }

        private void BtnTechnicals_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
                if (!ExcelUtils.WindowAvailable())
                {
                    return;
                }

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
            if (!ExcelUtils.WindowAvailable())
            {
                return;
            }

            ScreenerManager manager = new ScreenerManager();
            await manager.AddNewScreener();
        }
    }
}
