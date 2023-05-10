using EOD;
using EOD.Model.Bulks;
using EODAddIn.Program;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EODAddIn.BL.BulkEod
{
    internal class GetBulkEod
    {
        private static string AppName = Settings.SettingsFields.AppName;
        private static string ApiKey = Settings.SettingsFields.APIKey;
        public static async Task<List<Bulk>> GetBulkEodData(string exchange, string type, DateTime? date, string symbols)
        {
            API api = new API(ApiKey, null, AppName);
            List<Bulk> bulkEodData = await api.GetBulksAsync(exchange, type, date, symbols);

            return bulkEodData;
        }
    }
}
