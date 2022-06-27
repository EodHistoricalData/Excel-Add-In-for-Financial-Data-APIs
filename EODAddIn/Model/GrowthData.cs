using Newtonsoft.Json;

namespace EODAddIn.Model
{
    public class GrowthData
    {
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Long-Term Projected Earnings Growth")]
        public string LongTermProjectedEarningsGrowth { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Historical Earnings Growth")]
        public string HistoricalEarningsGrowth { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Sales Growth")]
        public string SalesGrowth { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Cash-Flow Growth")]
        public string CashFlowGrowth { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Book-Value Growth")]
        public string BookValueGrowth { get; set; }
    }
}
