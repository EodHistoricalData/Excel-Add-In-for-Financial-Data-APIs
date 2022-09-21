using Newtonsoft.Json;
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
