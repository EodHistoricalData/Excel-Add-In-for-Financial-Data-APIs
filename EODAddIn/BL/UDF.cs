﻿using EODAddIn.Utils;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EOD.Model;
namespace EODAddIn.BL
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class UDF
    {
        #region EOD Historical
        public double? EOD_GetOpen(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.HistoricalStockPrice> res = HistoricalAPI.HistoricalAPI.GetEOD(ticker, date, date, "d");
                return res[0].Open;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public double? EOD_GetHigh(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.HistoricalStockPrice> res = HistoricalAPI.HistoricalAPI.GetEOD(ticker, date, date, "d");
                return res[0].High;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public double? EOD_GetLow(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.HistoricalStockPrice> res = HistoricalAPI.HistoricalAPI.GetEOD(ticker, date, date, "d");
                return res[0].Low;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public double? EOD_GetClose(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.HistoricalStockPrice> res = HistoricalAPI.HistoricalAPI.GetEOD(ticker, date, date, "d");
                return res[0].Close;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public double? EOD_GetAdjustedClose(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.HistoricalStockPrice> res = HistoricalAPI.HistoricalAPI.GetEOD(ticker, date, date, "d");
                return res[0].Adjusted_close;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public double? EOD_GetVolume(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.HistoricalStockPrice> res = HistoricalAPI.HistoricalAPI.GetEOD(ticker, date, date, "d");
                return res[0].Volume;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        #endregion

        #region Intraday
        public double? Intraday_GetOpen(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.IntradayHistoricalStockPrice> res = IntradayAPI.IntradayAPI.GetIntraday(ticker, date, date,  EOD.API.IntradayHistoricalInterval.OneMinute).Result;
                return res[0].Open;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public double? Intraday_GetHigh(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.IntradayHistoricalStockPrice> res = IntradayAPI.IntradayAPI.GetIntraday(ticker, date, date,  EOD.API.IntradayHistoricalInterval.OneMinute).Result;
                return res[0].High;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public double? Intraday_GetLow(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.IntradayHistoricalStockPrice> res = IntradayAPI.IntradayAPI.GetIntraday(ticker, date, date, EOD.API.IntradayHistoricalInterval.OneMinute).Result;
                return res[0].Low;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public double? Intraday_GetClose(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.IntradayHistoricalStockPrice> res = IntradayAPI.IntradayAPI.GetIntraday(ticker, date, date, EOD.API.IntradayHistoricalInterval.OneMinute).Result;
                return res[0].Close;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public decimal? Intraday_GetVolume(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.IntradayHistoricalStockPrice> res = IntradayAPI.IntradayAPI.GetIntraday(ticker, date, date, EOD.API.IntradayHistoricalInterval.OneMinute).Result;
                return res[0].Volume;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public long? Intraday_GetTimeStamp(string ticker, DateTime date)
        {
            try
            {
                List<EOD.Model.IntradayHistoricalStockPrice> res = IntradayAPI.IntradayAPI.GetIntraday(ticker, date, date, EOD.API.IntradayHistoricalInterval.OneMinute).Result;
                return res[0].Timestamp;
            }
            catch (APIException ex)
            {
                MessageBox.Show(ex.StatusError, "Error load " + ticker, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        #endregion
    }
}
