using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class Earnings
    {
        public Dictionary<DateTime, EarningsHistoryData> History { get; set; }
        public Dictionary<DateTime, EarningsTrendData> Trend { get; set; }
    }
}
