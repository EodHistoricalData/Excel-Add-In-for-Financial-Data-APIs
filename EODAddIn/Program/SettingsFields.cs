using System;
using System.Collections.Generic;

namespace EODAddIn.Program
{
    [Serializable]
    public class SettingsFields
    {
        /// <summary>
        /// API ключ
        /// </summary>
        public string APIKey = "OeAFFmMliFG5orCUuwAKQ8l4WWFQ67YX";

        public List<string> EndOfDayTickers;
        public string EndOfDayPeriod;
        public DateTime EndOfDayFrom = new DateTime(1970, 1, 1);
        public DateTime EndOfDayTo;

        public List<string> IntradayTickers;
        public string IntradayInterval;
        public DateTime IntradayFrom = new DateTime(1970, 1, 1);
        public DateTime IntradayTo;

        public string FundamentalTicker;

        public SettingsFields()
        {

        }
    }
}
