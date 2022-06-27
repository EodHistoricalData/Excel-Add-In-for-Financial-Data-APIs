using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    /// <summary>
    /// Записи значений в течении дня
    /// </summary>
    public class Intraday
    {
        //  public DateTime? Timestamp { get; set; }
        public long? Timestamp { get; set; }
        public double? Gmtoffset { get; set; }
        public DateTime? DateTime { get; set; }
        public double? Open { get; set; }
        public double? High { get; set; }
        public double? Low { get; set; }
        public double? Close { get; set; }
        public decimal? Volume { get; set; }      

        //public DateTime? Date 
        //{ get =>{return DateTime. }
        //    set; }
    }
}
