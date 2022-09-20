using EOD.Model.OptionsData;
using EOD;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using EODAddIn.Utils;

namespace EODAddIn.BL.FundamentalDataAPI
{
    public class FundamentalDataAPI
    {
        public static EOD.Model.Fundamental.FundamentalData GetFundamental(string code)
        {
            string url = $"https://eodhistoricaldata.com/api/fundamentals/{code}";
            string data = $"api_token={Program.Program.APIKey}";
            string s = Response.GET(url, data);
            return JsonConvert.DeserializeObject<EOD.Model.Fundamental.FundamentalData>(s);
        }
    }
}
