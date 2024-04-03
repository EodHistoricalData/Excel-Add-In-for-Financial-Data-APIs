using System;
using System.Collections.Generic;

namespace EODAddIn.Program.SettingsModel
{
    [Serializable]
    public class GetHistoricalFormSettings
    {
        public List<string> Tickers = new List<string>();

        public DateTime From = DateTime.Today.AddYears(-1);
        public DateTime To = DateTime.Today;

        public string Period = string.Empty;

        public bool SmartTable = true;

        public bool AddDate = true;

        public bool OrderDesc = true;

        public string TypeOfOutput = string.Empty;
    }
}
