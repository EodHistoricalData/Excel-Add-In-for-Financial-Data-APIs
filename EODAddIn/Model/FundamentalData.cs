using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class FundamentalData
    {
        public General General { get; set; }
        public Highlights Highlights { get; set; }
        public Earnings Earnings { get; set; }
        public Financials Financials { get; set; }
    }
}
