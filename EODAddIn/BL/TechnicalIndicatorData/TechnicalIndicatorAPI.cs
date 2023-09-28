using EODHistoricalData.Wrapper.Model.TechnicalIndicators;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using static EOD.API;

namespace EODAddIn.BL.TechnicalIndicatorData
{
    internal class TechnicalIndicatorAPI
    {
        public static async Task<List<TechnicalIndicator>> GetTechnicalIndicatorsData(string code, DateTime? from, DateTime? to, Order? order, List<IndicatorParameters> parameters)
        {
            EOD.API api = new EOD.API(Program.Program.APIKey);
            var result = await api.GetTechnicalIndicatorsAsync(code, from, to, order, parameters);
            return result;
        }
    }
}
