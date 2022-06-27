using Newtonsoft.Json;

namespace EODAddIn.Model
{
    public class WorldRegionsData
    {
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Equity_%")]
        public string EquityPercent { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Relative_to_Category")]
        public string RelativeToCategory { get; set; }
    }
}
