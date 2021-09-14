using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class Highlights
    {
        public long? MarketCapitalization { get; set; }
        public double? MarketCapitalizationMln { get; set; }
        public long? EBITDA { get; set; }
        public double? PERatio { get; set; }
        public double? PEGRatio { get; set; }
        public double? WallStreetTargetPrice { get; set; }
        public double? BookValue { get; set; }
        public double? DividendShare { get; set; }
        public double? DividendYield { get; set; }
        public double? EarningsShare { get; set; }
        public double? EPSEstimateCurrentYear { get; set; }
        public double? EPSEstimateNextYear { get; set; }
        public double? EPSEstimateNextQuarter { get; set; }
        public double? EPSEstimateCurrentQuarter { get; set; }
        public DateTime? MostRecentQuarter { get; set; }
        public double? ProfitMargin { get; set; }
        public double? OperatingMarginTTM { get; set; }
        public double? ReturnOnAssetsTTM { get; set; }
        public double? ReturnOnEquityTTM { get; set; }
        public long? RevenueTTM { get; set; }
        public double? RevenuePerShareTTM { get; set; }
        public double? QuarterlyRevenueGrowthYOY { get; set; }
        public long? GrossProfitTTM { get; set; }
        public double? DilutedEpsTTM { get; set; }
        public int? QuarterlyEarningsGrowthYOY { get; set; }
    }
}
