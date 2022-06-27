using EOD;
using EOD.Model.OptionsData;
using EODAddIn.Program;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Utils
{
    internal class GetOptions
    {
        private static string AppName = Settings.SettingsFields.AppName;
        private static string ApiKey = Settings.SettingsFields.APIKey;
        public static async Task<OptionsData> GetOptionsData(string ticker, DateTime from, DateTime to, DateTime? fromTrade, DateTime? toTrade)
        {
            API api = new API(ApiKey, null, AppName);
            OptionsData optionsData = await api.GetOptionsDataAsync(ticker, from, to, fromTrade, toTrade, null);

            return optionsData;
        }
    }
}
