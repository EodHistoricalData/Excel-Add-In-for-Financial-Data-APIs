using EODAddIn.Model;
using EODAddIn.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.BL.HistoricalAPI
{
    public class HistoricalAPI
    {
        public static List<EndOfDay> GetEOD(string code, DateTime from, DateTime to, string period = "d")
        {
            string url = $"https://eodhistoricaldata.com/api/eod/{code}";
            string data = $"fmt=json&period={period}&from={from:yyyy-MM-dd}&to={to:yyyy-MM-dd}&api_token={Program.Program.APIKey}";
            string s = Response.GET(url, data);
            return JsonConvert.DeserializeObject<List<EndOfDay>>(s);
        }
    }
}
