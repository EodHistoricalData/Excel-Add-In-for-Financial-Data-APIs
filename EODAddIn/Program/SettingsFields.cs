using EODAddIn.Program.SettingsModel;

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace EODAddIn.Program
{
    [Serializable]
    public class SettingsFields
    {
        /// <summary>
        /// API key
        /// </summary>
        public string APIKey = "demo";
        /// <summary>
        /// Program name
        /// </summary>
        public string AppName = "EOD Excel Plug-in";

        public GetHistoricalFormSettings GetHistoricalForm = new GetHistoricalFormSettings();

        public List<string> EndOfDayTickers;
        public string EndOfDayPeriod;
        public DateTime EndOfDayFrom = new DateTime(1970, 1, 1);
        public DateTime EndOfDayTo;

        public List<string> IntradayTickers;
        public string IntradayInterval;
        public DateTime IntradayFrom = new DateTime(1970, 1, 1);
        public DateTime IntradayTo;

        public string FundamentalTicker;

        public string EtfTicker;

        public string OptionsTicker;

        public string BulkFundamentalExchange;
        public List<string> BulkFundamentalTickers;
        public int BulkFundamentalOffset = 0;
        public int BulkFundamentalLimit = 500;

        public string ScreenerSector;
        public string ScreenerIndustry;
        public string ScreenerCode;
        public string ScreenerName;
        public string ScreenerExchange;
        public int ScreenerLimit=100;

        public CheckState Screener200d_New_Lo;
        public CheckState Screener200d_New_Hi;
        public CheckState ScreenerBookValue_Neg;
        public CheckState ScreenerBookValue_Pos;
        public CheckState ScreenerWallStreet_Lo;
        public CheckState ScreenerWallStreet_Hi;
        public bool ScreenerRbtnSortAsc;
        public bool ScreenerRbtnSortDesc;
        public List<(string, string, string)> ScreenerDataGridViewFilters;
        public bool IsInfoShowed;

        public string BulkEodExchange;
        public string BulkEodType;
        public DateTime BulkEodDate;
        public List<string> BulkEodSymbols;

        public List<string> TechnicalsTickers;
        public DateTime TechnicalsFrom = new DateTime(2020, 1, 1);
        public DateTime TechnicalsTo = DateTime.Today;
        public int TechnicalsFunctionId;

        public SettingsFields()
        {

        }
    }
}
