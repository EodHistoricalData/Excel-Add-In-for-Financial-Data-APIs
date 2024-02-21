using EOD;
using EOD.Model.OptionsData;
using EODAddIn.Program;
using System;
using System.Threading.Tasks;

namespace EODAddIn.BL.OptionsAPI
{
    public class OptionsAPI
    {
        private static string AppName = Settings.Data.AppName;
        private static string ApiKey = Settings.Data.APIKey;
        public static async Task<OptionsData> GetOptionsData(string ticker, DateTime from, DateTime to, DateTime? fromTrade, DateTime? toTrade)
        {
            API api = new API(ApiKey, null, AppName);
            OptionsData optionsData = await api.GetOptionsDataAsync(ticker, from, to, fromTrade, toTrade, null);

            return optionsData;
        }
    }
}
