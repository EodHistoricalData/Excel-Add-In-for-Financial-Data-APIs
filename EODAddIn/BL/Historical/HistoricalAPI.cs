using EOD.Model;
using EODAddIn.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EODAddIn.BL.HistoricalAPI
{
    public class HistoricalAPI
    {
        public static List<HistoricalStockPrice> GetEOD(string code, DateTime from, DateTime to, string period = "d")
        {
            string url = $"https://eodhistoricaldata.com/api/eod/{code}";
            string data = $"fmt=json&period={period}&from={from:yyyy-MM-dd}&to={to:yyyy-MM-dd}&api_token={Program.Program.APIKey}";
            string s = Response.GET(url, data);
            return JsonConvert.DeserializeObject<List<HistoricalStockPrice>>(s);
        }

        public static async Task<List<HistoricalStockPrice>> GetHistoricalStockPrice(string code, DateTime from, DateTime to, string period = "d")
        {
            EOD.API api = new EOD.API(Program.Program.APIKey);
            EOD.API.HistoricalPeriod historicalPeriod;
            switch (period)
            {
                case "w":
                    historicalPeriod = EOD.API.HistoricalPeriod.Weekly;
                    break;
                case "m":
                    historicalPeriod = EOD.API.HistoricalPeriod.Monthly;
                    break;
                default:
                    historicalPeriod = EOD.API.HistoricalPeriod.Daily;
                    break;
            }
            var result = await api.GetEndOfDayHistoricalStockPriceAsync(code, from, to, historicalPeriod);
            return result;
        }
    }
}
