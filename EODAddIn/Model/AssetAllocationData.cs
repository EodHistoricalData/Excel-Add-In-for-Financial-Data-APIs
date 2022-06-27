using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class AssetAllocationData
    {
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Long_%")]
        public string LongPercent { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Short_%")]
        public string ShortPercent { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Net_Assets_%")]
        public string NetAssetsPercent { get; set; }
    }
}
