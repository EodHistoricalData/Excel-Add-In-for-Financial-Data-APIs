using Newtonsoft.Json;

namespace EODAddIn.Model
{
    public class HoldingData
    {
        /// <summary>
        /// 
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 
        /// </summary>
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Assets_%")]
        public string AssetsPercent { get; set; }
    }
}
