using EOD.Model.ExchangeDetails;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EODAddIn.BL.Live
{
    public class ExchangeDownloadRules
    {
        public string Exchange { get; set; }
        public List<DateTime> Holidays { get; set; } = new List<DateTime>();
        public DateTime Open { get; set; }
        public DateTime Close { get; set; }
        public List<string> DaysofWeek { get; set; } = new List<string>();
        public int Gtmoffset { get; set; }
        public ExchangeDownloadRules()
        {

        }
        public ExchangeDownloadRules(ExchangeDetail details)
        {
            Exchange = details.Code;
            Open = DateTime.Parse(details.TradingHours.Open);
            Close = DateTime.Parse(details.TradingHours.Close);
            Gtmoffset = (DateTime.Parse(details.TradingHours.Open) - DateTime.Parse(details.TradingHours.OpenUTC)).Hours;
            DaysofWeek = details.TradingHours.WorkingDays.Split(',').ToList();
            foreach (var item in details.ExchangeHolidays)
            {
                Holidays.Add(item.Value.Date == null ? DateTime.MinValue : (DateTime)item.Value.Date);
            }
        }
    }
}
