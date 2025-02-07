using EOD;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EODAddIn.BL.Screener
{
    public class ScreenerAPI
    {
        public static async Task<EOD.Model.Screener.StockMarkerScreener> GetScreener(List<(API.Field, API.Operation, string)> filters, List<API.Signal> signals, (API.Field, API.Order)? sort, int limit)
        {
            API api = new API(Program.Settings.Data.APIKey, null, Program.Settings.Data.AppName);

            EOD.Model.Screener.StockMarkerScreener screener = new EOD.Model.Screener.StockMarkerScreener
            {
                Data = new List<EOD.Model.Screener.ScreenerData>()
            };
            var fixed_filters = new List<(API.Field, API.Operation, string)>();
            for (int i = 0; i < filters.Count; i++)
            {
                (API.Field, API.Operation, string) temp;
                temp.Item1 = filters[i].Item1;
                temp.Item2 = filters[i].Item2;
                temp.Item3 = filters[i].Item3.Replace("&", "%26");
                fixed_filters.Add(temp);
            }

            for (int i = 0, remains = limit; i < limit; i += 100)
            {
                int ilimit;
                if (remains > 100)
                {
                    ilimit = 100;
                }
                else
                {
                    ilimit = remains;
                }

                EOD.Model.Screener.StockMarkerScreener request = await api.GetStockMarketScreenerAsync(fixed_filters, signals, sort, ilimit, i);
                if (request.Data.Count == 0)
                    break;
                screener.Data.AddRange(request.Data);
                remains -= 100;
            }

            return screener;
        }
    }
}
