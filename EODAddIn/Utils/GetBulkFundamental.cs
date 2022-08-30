﻿using EOD;
using EOD.Model.BulkFundamental;
using EODAddIn.Program;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Utils
{
    internal class GetBulkFundamental
    {
        private static string AppName = Settings.SettingsFields.AppName;
        private static string ApiKey = Settings.SettingsFields.APIKey;
        public static async Task<Dictionary<string, BulkFundamentalData>> GetBulkData(string exchange, List<string> symbols, int offset, int limit)
        {
            string symbolStr = string.Empty;
            if (symbols.Count != 0)
            {
                symbolStr = string.Join(",", symbols.ToArray());
            }
            API api = new API(ApiKey, null, AppName);
            Dictionary<string, BulkFundamentalData> bulkFundamentalData = await api.GetBulkFundamentalsDataAsync(exchange, offset, limit, symbolStr);

            return bulkFundamentalData;
        }
    }
}
