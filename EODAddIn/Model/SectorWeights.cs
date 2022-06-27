using Newtonsoft.Json;

namespace EODAddIn.Model
{
    public class SectorWeights
    {
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Basic Materials")]
        public WorldRegionsData BasicMaterials { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Consumer Cyclicals")]
        public WorldRegionsData ConsumerCyclicals { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Financial Services")]
        public WorldRegionsData FinancialServices { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Real Estate")]
        public WorldRegionsData RealEstate { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Communication Services")]
        public WorldRegionsData CommunicationServices { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public WorldRegionsData Energy { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public WorldRegionsData Industrials { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public WorldRegionsData Technology { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Consumer Defensive")]
        public WorldRegionsData ConsumerDefensive { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public WorldRegionsData Healthcare { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public WorldRegionsData Utilities { get; set; }
    }
}
