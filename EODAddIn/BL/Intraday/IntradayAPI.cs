using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using EODAddIn.Utils;
using EODHistoricalData.Wrapper.Model.TechnicalIndicators;
using static EOD.API;
using System.Threading.Tasks;

namespace EODAddIn.BL.IntradayAPI
{
    public class IntradayAPI
    {
        public static async Task<List<EOD.Model.IntradayHistoricalStockPrice>> GetIntraday(string code, DateTime from, DateTime to, IntradayHistoricalInterval interval)
        {
            EOD.API api = new EOD.API(Program.Program.APIKey);
            var result = await api.GetIntradayHistoricalStockPriceAsync(code, from, to, interval);
            return result;
            //long unixFrom = ((DateTimeOffset)from).ToUnixTimeSeconds();
            //long unixTo = ((DateTimeOffset)to).ToUnixTimeSeconds();

            //string url = $"https://eodhistoricaldata.com/api/intraday/{code}";
            ////string data = $"fmt=json&interval={interval}&from={from:yyyy-MM-dd}&to={to:yyyy-MM-dd}&api_token={Program.Program.APIKey}";
            //string data = $"api_token={Program.Program.APIKey}&interval={interval}&fmt=json&from={unixFrom}&to={unixTo}";
            //string s = Response.GET(url, data);
            //return JsonConvert.DeserializeObject<List<EOD.Model.IntradayHistoricalStockPrice>>(s);
        }

    }
}
