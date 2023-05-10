using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using EODAddIn.Utils;

namespace EODAddIn.BL.IntradayAPI
{
    public class IntradayAPI
    {
        public static List<EOD.Model.IntradayHistoricalStockPrice> GetIntraday(string code, DateTime from, DateTime to, string interval = "5m")
        {
            long unixFrom = ((DateTimeOffset)from).ToUnixTimeSeconds();
            long unixTo = ((DateTimeOffset)to).ToUnixTimeSeconds();

            string url = $"https://eodhistoricaldata.com/api/intraday/{code}";
            //string data = $"fmt=json&interval={interval}&from={from:yyyy-MM-dd}&to={to:yyyy-MM-dd}&api_token={Program.Program.APIKey}";
            string data = $"api_token={Program.Program.APIKey}&interval={interval}&fmt=json&from={unixFrom}&to={unixTo}";
            string s = Response.GET(url, data);
            return JsonConvert.DeserializeObject<List<EOD.Model.IntradayHistoricalStockPrice>>(s);
        }
    }
}
