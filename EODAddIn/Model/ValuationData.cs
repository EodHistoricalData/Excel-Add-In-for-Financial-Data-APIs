using Newtonsoft.Json;

namespace EODAddIn.Model
{
    public class ValuationData
    {
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Price/Prospective Earnings")]
        public string PriceToProspectiveEarnings { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Price/Book")]
        public string PriceToBook { get; set;}
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Price/Sales")]
        public string PriceToSales { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Price/Cash Flow")]
        public string PriceToCashFlow { get; set;}
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Dividend-Yield Factor")]
        public string DividendYieldFactor { get; set; }
    }
}
