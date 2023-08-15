using Newtonsoft.Json;
using EODAddIn.Utils;
using System.Threading.Tasks;
using EOD.Model.Fundamental;

namespace EODAddIn.BL.FundamentalDataAPI
{
    public class FundamentalDataAPI
    {
        //public static EOD.Model.Fundamental.FundamentalData GetFundamental(string code)
        //{
        //    string url = $"https://eodhistoricaldata.com/api/fundamentals/{code}";
        //    string data = $"api_token={Program.Program.APIKey}";
        //    string s = Response.GET(url, data);
        //    return JsonConvert.DeserializeObject<EOD.Model.Fundamental.FundamentalData>(s);
        //}

        public static async Task<FundamentalData> GetFundamentalData(string code)
        {
            EOD.API api = new EOD.API(Program.Program.APIKey);
            var result = await api.GetFundamentalDataAsync(code);
            return result;
        }
    }
}
