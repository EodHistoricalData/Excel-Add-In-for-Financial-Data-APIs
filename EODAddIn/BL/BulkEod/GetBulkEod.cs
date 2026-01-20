using EOD;
using EOD.Model.Bulks;
using EODAddIn.Program;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EODAddIn.BL.BulkEod
{
    public class GetBulkEod
    {
        private static string AppName = Settings.Data.AppName;
        private static string ApiKey = Settings.Data.APIKey;
        public static async Task<List<Bulk>> GetBulkEodData(string exchange, EODHistoricalData.Wrapper.Model.Bulks.BulkQueryTypes type, DateTime? date, string symbols)
        {
            API api = new API(ApiKey, null, AppName);
            List<Bulk> bulkEodData = await api.GetBulksAsync(exchange, type, date, symbols);

            return bulkEodData;
        }
    }
}
