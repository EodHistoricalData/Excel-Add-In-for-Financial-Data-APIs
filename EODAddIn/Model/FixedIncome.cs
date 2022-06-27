using Newtonsoft.Json;

namespace EODAddIn.Model
{
    public class FixedIncome
    {
        /// <summary>
        /// 
        /// </summary>
        public FixedIncomeData EffectiveDuration { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public FixedIncomeData ModifiedDuration { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public FixedIncomeData EffectiveMaturity { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public FixedIncomeData YieldToMaturity { get; set; }
    }
}
