using EOD.Model.OptionsData;
using EOD;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace EODAddIn.BL.Screener
{
    public class ScreenerAPI
    {
        public static async Task<EOD.Model.Screener.StockMarkerScreener> GetScreener(List<(API.Field, API.Operation, string)> filters, List<API.Signal> signals, (API.Field, API.Order)? sort, int limit)
        {
            API api = new API(Program.Settings.SettingsFields.APIKey, null, Program.Settings.SettingsFields.AppName);
            
            EOD.Model.Screener.StockMarkerScreener screener = new EOD.Model.Screener.StockMarkerScreener();
            screener.Data = new List<EOD.Model.Screener.ScreenerData>();
            for (int i = 0; i < limit; i += 100)
            {
                int ilimit;
                if (limit > 100)
                {
                    ilimit = 100;
                }
                else
                {
                    ilimit = limit;
                }

                EOD.Model.Screener.StockMarkerScreener request = await api.GetStockMarketScreenerAsync(filters, signals, sort, ilimit, i);
                if (request.Data.Count == 0) break;
                screener.Data.AddRange(request.Data);
                limit -= ilimit;
            }

            return screener;
        }
    }
}
