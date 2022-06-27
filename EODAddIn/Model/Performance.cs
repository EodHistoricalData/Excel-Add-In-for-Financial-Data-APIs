using Newtonsoft.Json;

namespace EODAddIn.Model
{
    public class Performance
    {
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("1y_Volatility")]
        public string Volatility1y { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("3y_Volatility")]
        public string Volatility3y { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("3y_ExpReturn")]
        public string ExpReturn3y { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("3y_SharpRatio")]
        public string SharpRatio3y { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string Returns_YTD { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string Returns_1Y { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string Returns_3Y { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string Returns_5Y { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string Returns_10Y { get; set; }
    }
}
