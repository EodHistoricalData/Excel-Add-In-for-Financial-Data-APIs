using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class ValuationGrowthData
    {
        public ValuationData Valuations_Rates_Portfolio { get; set; }
        public ValuationData Valuations_Rates_To_Category { get; set; }
        public GrowthData Growth_Rates_Portfolio { get; set; }
        public GrowthData Growth_Rates_To_Category { get; set; }
    }
}
