using Newtonsoft.Json;

namespace EODAddIn.Model
{
    public class FixedIncomeData
    {
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Fund_%")]
        public string FundPercent { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Relative_to_Category")]
        public string RelativeToCategory { get; set; }
    }
}
