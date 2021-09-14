using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class EarningsHistoryData
    {
        public DateTime? ReportDate { get; set; }
        public DateTime? Date { get; set; }
        public string BeforeAfterMarket { get; set; }
        public string Currency { get; set; }
        public double? EpsActual { get; set; }
        public double? EpsEstimate { get; set; }
        public double? EpsDifference { get; set; }
        public double? SurprisePercent { get; set; }
    }
}
