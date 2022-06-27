using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class AssetAllocation
    {
        /// <summary>
        /// 
        /// </summary>
        public AssetAllocationData Cash { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public AssetAllocationData NotClassified { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Stock non-Us")]
        public AssetAllocationData StockNonUs { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public AssetAllocationData Other { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("Stock Us")]
        public AssetAllocationData StockUs { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public AssetAllocationData Bond { get; set; }
    }
}
