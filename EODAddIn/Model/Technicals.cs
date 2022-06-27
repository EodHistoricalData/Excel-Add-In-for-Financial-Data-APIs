using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    /// <summary>
    /// ETFs Supported Data
    /// </summary>
    public class Technicals
    {
        /// <summary>
        /// 
        /// </summary>
        public double? Beta { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("52WeekHigh")]
        public double? WeekHigh52 { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("52WeekLow")]
        public double? WeekLow52 { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("50DayMA")]
        public double? DayMA50 { get; set; }
        /// <summary>
        /// 
        /// </summary>
        [JsonProperty("200DayMA")]
        public double? DayMA200 { get; set; }
    }
}
