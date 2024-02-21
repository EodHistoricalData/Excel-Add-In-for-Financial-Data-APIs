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
using static EODAddIn.Utils.ExcelUtils;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Threading;
using System.IO;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;
using EODAddIn.View.Forms;

namespace EODAddIn
{
    public partial class Ribbon
    {
        private Excel.Application _xlapp;
        private DispatcherTimer timer = null;

        private Dictionary<LiveDownloader, CustomXMLPart> LiveDownloaders = new Dictionary<LiveDownloader, CustomXMLPart>();
        private Dictionary<LiveDownloader, CancellationTokenSource> CancellationTokenSources = new Dictionary<LiveDownloader, CancellationTokenSource>();
        private bool DispatcherIsOpened = false;
        private delegate void Download();
        private Download _download;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            timer = new DispatcherTimer();
            timer.Tick += new EventHandler(UpdateRequests);
            timer.Interval = new TimeSpan(0, 0, 0, 20, 0);
            timer.Start();

            _xlapp = Globals.ThisAddIn.Application;
            _xlapp.WorkbookOpen += Xlapp_WorkbookOpen;
        }

        private void BtnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            Program.FrmAbout frm = new Program.FrmAbout();
            frm.ShowDialog();
        }

        private void BtnSettings_Click(object sender, RibbonControlEventArgs e)
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

        private void GetHistorical_Click(object sender, RibbonControlEventArgs e) => FormShower.FrmGetHistoricalShow();

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
            if (frm.DialogResult != DialogResult.OK)
            {
                return;
            }
            EOD.Model.Fundamental.FundamentalData res = frm.Results;
            FundamentalDataPrinter.PrintFundamentalAll(res, frm.Tiker);
        }

        private void BtnGetIntradayHistoricalData_Click(object sender, RibbonControlEventArgs e) => FormShower.FrmGetIntradayHistoricalDataShow();


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

        private async void btnCreateScreener_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Forms.FrmScreener frm = new Forms.FrmScreener();
                frm.ShowDialog(new WinHwnd());
                if (frm.DialogResult == DialogResult.OK)
                {
                    BtnOptions.Label = "Processing";
                    BtnOptions.Enabled = false;

                    var res = await ScreenerAPI.GetScreener(frm.Filters, frm.Signals, frm.Sort, frm.Limit);
                    ScreenerPrinter.PrintScreener(res);

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

        private void btnGetScreenerFundamental_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!ScreenerPrinter.CheckIsScreenerResult(Globals.ThisAddIn.Application.ActiveSheet))
                {
                    return;
                }
                ScreenerPrinter.PrintScreenerBulk(ScreenerPrinter.GetTickersFromScreener());
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

        private void BtnListOfIndices_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhd.com/financial-apis/list-supported-indices/?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");
        }

        private async void BtnBulkEod_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Forms.FrmGetBulkEod frm = new Forms.FrmGetBulkEod();
                frm.ShowDialog(new WinHwnd());

                if (frm.DialogResult == DialogResult.OK)
                {
                    string exchange = frm.Exchange;
                    string type = null;
                    switch (frm.Type)
                    {
                        case "end-of-day":
                            break;
                        case "":
                            break;
                        default:
                            type = frm.Type;
                            break;
                    }
                    DateTime date = frm.Date;
                    string tickers = string.Join(",", frm.Tickers);
                    BtnBulkEod.Label = "Processing";
                    BtnBulkEod.Enabled = false;

                    List<Bulk> res = await GetBulkEod.GetBulkEodData(exchange, type, date, tickers);
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

        private void BtnTechnicals_Click(object sender, RibbonControlEventArgs e) => FormShower.FrmGetTechnicalsShow();

        private void BtnGetLive_Click(object sender, RibbonControlEventArgs e)
        {
            LiveDownloaderDispatcher frm;
            if (!DispatcherIsOpened)
            {
                if (LiveDownloaders.Count == 0)
                {
                    frm = new LiveDownloaderDispatcher();
                }
                else
                {
                    frm = new LiveDownloaderDispatcher(LiveDownloaders, CancellationTokenSources);                 
                }
                frm.FormClosing += Frm_FormClosing;
                FormShower.FrmShow(frm);
                DispatcherIsOpened = true;
            }
        }

        private void Xlapp_WorkbookOpen(Excel.Workbook Wb)
        {
            var xml = GetXmlPart();
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(LiveDownloader));
            if (xml == null) return;
            foreach (CustomXMLPart item in xml)
            {
                LiveDownloader liveDownloader = null;
                try
                {
                    using (TextReader reader = new StringReader(item.XML))
                    {
                        liveDownloader = xmlSerializer.Deserialize(reader) as LiveDownloader;
                    }
                }
                catch
                {

                }
                finally
                {
                    if (liveDownloader != null)
                        LiveDownloaders.Add(liveDownloader, item);
                }
            }

            foreach (var pair in LiveDownloaders)
            {
                if (pair.Key.IsActive == true)
                    StartDowloader(pair.Key);
            }
        }

        private void StartDowloader(LiveDownloader downloader)
        {
            CancellationTokenSource src = new CancellationTokenSource();
            CancellationTokenSources.Add(downloader, src);

            async void Live() => await downloader.RequestAndPrint(src.Token);
            _download = Live;
            _download.Invoke();
        }

        private void Frm_FormClosing(object sender, FormClosingEventArgs e)
        {
            LiveDownloaderDispatcher frm = (LiveDownloaderDispatcher)sender;
            LiveDownloaders = frm.GetDownloaders();
            CancellationTokenSources = frm.GetTokens();
            DispatcherIsOpened = false;
        }
    }
}
