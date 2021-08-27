using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class EarningsTrendData
    {
        public DateTime Date { get; set; }
        public string Period { get; set; }
        public double? Growth { get; set; }
        public double? EarningsEstimateAvg { get; set; }
        public double? EarningsEstimateLow { get; set; }
        public double? EarningsEstimateHigh { get; set; }
        public double? EarningsEstimateYearAgoEps { get; set; }
        public double? EarningsEstimateNumberOfAnalysts { get; set; }
        public double? EarningsEstimateGrowth { get; set; }
        public double? RevenueEstimateAvg { get; set; }
        public double? RevenueEstimateLow { get; set; }
        public double? RevenueEstimateHigh { get; set; }
        public double? RevenueEstimateYearAgoEps { get; set; }
        public double? RevenueEstimateNumberOfAnalysts { get; set; }
        public double? RevenueEstimateGrowth { get; set; }
        public double? EpsTrendCurrent { get; set; }
        public double? EpsTrend7daysAgo { get; set; }
        public double? EpsTrend30daysAgo { get; set; }
        public double? EpsTrend60daysAgo { get; set; }
        public double? EpsTrend90daysAgo { get; set; }
        public double? EpsRevisionsUpLast7days { get; set; }
        public double? EpsRevisionsUpLast30days { get; set; }
        public double? EpsRevisionsDownLast30days { get; set; }
        public double? EpsRevisionsDownLast90days { get; set; }
    }
}
