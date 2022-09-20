using System;
using System.Collections.Generic;

namespace EODAddIn.Program
{
    [Serializable]
    public class SettingsFields
    {
        /// <summary>
        /// API key
        /// </summary>
        public string APIKey = "OeAFFmMliFG5orCUuwAKQ8l4WWFQ67YX";
        /// <summary>
        /// Program name
        /// </summary>
        public string AppName = "EOD Excel Plug-in";

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

        public SettingsFields()
        {

        }
    }
}
