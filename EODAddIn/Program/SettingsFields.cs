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
        public string APIKey = "";

        public List<string> EndOfDayTickers;
        public string EndOfDayPeriod;
        public DateTime EndOfDayFrom = new DateTime(1970, 1, 1);
        public DateTime EndOfDayTo;

        public string FundamentalTicker;

        public SettingsFields()
        {

        }
    }
}
