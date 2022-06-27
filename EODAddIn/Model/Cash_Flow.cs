using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class Cash_Flow
    {
        public string Currency_symbol { get; set; }
        public Dictionary<DateTime, Cash_FlowData> Quarterly { get; set; }
        public Dictionary<DateTime, Cash_FlowData> Yearly { get; set; }
    }
}
