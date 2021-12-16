﻿using EODAddIn.BL;
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

        private void SplitbtnFundamental_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetFundamental frm = new Forms.FrmGetFundamental();
            frm.ShowDialog(new WinHwnd());

            Model.FundamentalData res = frm.Results;
            if (res == null) return;
            LoadToExcel.PrintFundamentalAll(res);

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
            LoadToExcel.PrintFundamentalAll(res);
        }

        private void BtnGetIntradayHistoricalData_Click(object sender, RibbonControlEventArgs e)
        {
            Forms.FrmGetIntradayHistoricalData frm = new Forms.FrmGetIntradayHistoricalData();
            frm.ShowDialog(new WinHwnd());
        }
    }
}
