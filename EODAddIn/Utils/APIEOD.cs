using EODAddIn.Model;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;

namespace EODAddIn.Utils
{
    public static class APIEOD
    {
        /// <summary>
        /// Запрос на получение информации о пользователе
        /// </summary>
        /// <param name="api_token"></param>
        /// <returns></returns>
        public static User User(string api_token)
        {
            string url = "https://eodhistoricaldata.com/api/user";
            string s = Response.GET(url, "api_token=" + api_token);
            return JsonConvert.DeserializeObject<User>(s);
        }

        /// <summary>
        /// Запрос на получение информации о пользователе
        /// </summary>
        /// <param name="api_token"></param>
        /// <returns></returns>
        public static List<SearchResult> Search (string queryString)
        {
            string url = $"https://eodhistoricaldata.com/api/query-search-extended/?q={queryString}";
            
            try
            {
                string s = Response.GET(url);
                return JsonConvert.DeserializeObject<List<SearchResult>>(s);
            }
            catch (Exception)
            {
                return new List<SearchResult>();
            }
        }

        public static List<EndOfDay> GetEOD(string code,  DateTime from, DateTime to, string period = "d")
        {
            string url = $"https://eodhistoricaldata.com/api/eod/{code}";
            string data = $"fmt=json&period={period}&from={from:yyyy-MM-dd}&to={to:yyyy-MM-dd}&api_token={Program.Program.APIKey}";
            string s = Response.GET(url, data);
            return JsonConvert.DeserializeObject<List<EndOfDay>>(s);
        }

        public static List<Intraday> GetIntraday(string code, DateTime from, DateTime to, string interval = "5m")
        {            
            long unixFrom = ((DateTimeOffset)from).ToUnixTimeSeconds();
            long unixTo = ((DateTimeOffset)to).ToUnixTimeSeconds();

            string url = $"https://eodhistoricaldata.com/api/intraday/{code}";
            //string data = $"fmt=json&interval={interval}&from={from:yyyy-MM-dd}&to={to:yyyy-MM-dd}&api_token={Program.Program.APIKey}";
            string data = $"api_token={Program.Program.APIKey}&interval={interval}&fmt=json&from={unixFrom}&to={unixTo}";
            string s = Response.GET(url, data);
            return JsonConvert.DeserializeObject<List<Intraday>>(s);
        }

        public static FundamentalData GetFundamental(string code)
        {
            string url = $"https://eodhistoricaldata.com/api/fundamentals/{code}";
            string data = $"api_token={Program.Program.APIKey}";
            string s = Response.GET(url, data);
            return JsonConvert.DeserializeObject<FundamentalData>(s);
        }
    }
}
