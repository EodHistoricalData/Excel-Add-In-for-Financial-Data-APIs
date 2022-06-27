using Newtonsoft.Json;

namespace EODAddIn.Model
{
    public class WorldRegions
    {
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("North America")]
        public WorldRegionsData NorthAmerica { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("United Kingdom")]
        public WorldRegionsData UnitedKingdom { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Europe Developed")]
        public WorldRegionsData EuropeDeveloped { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Europe Emerging")]
        public WorldRegionsData EuropeEmerging { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Africa/Middle East")]
        public WorldRegionsData AfricaMiddleEast { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public WorldRegionsData Japan { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public WorldRegionsData Australasia { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Asia Developed")]
        public WorldRegionsData AsiaDeveloped { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Asia Emerging")]
        public WorldRegionsData AsiaEmerging { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Latin America")]
        public WorldRegionsData LatinAmerica { get; set; }
    }
}
