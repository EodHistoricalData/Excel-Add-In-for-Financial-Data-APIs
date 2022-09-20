using EOD.Model.OptionsData;
using EOD;

using EODAddIn.Model;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace EODAddIn.Utils
{
    public static class APIEOD
    {
        /// <summary>
        ///User information request
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
        /// User information request
        /// </summary>
        /// <param name="api_token"></param>
        /// <returns></returns>
        public static List<SearchResult> Search(string queryString)
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

    }
}
