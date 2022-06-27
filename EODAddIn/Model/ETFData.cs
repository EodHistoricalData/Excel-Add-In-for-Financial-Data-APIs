using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class ETFData
    {
        /// <summary>
        /// ETFs Supported Data
        /// </summary>
        public string ISIN { get; set; }
        public string Company_Name { get; set; }
        public string Company_URL { get; set; }
        public string ETF_URL { get; set; }
        public string Domicile { get; set; }
        public string Index_Name { get; set; }
        public string Yield { get; set; }
        public string Dividend_Paying_Frequency { get; set; }
        public string Inception_Date { get; set; }
        public string Max_Annual_Mgmt_Charge { get; set; }
        public string Ongoing_Charge { get; set; }
        public string Date_Ongoing_Charge { get; set; }
        public string NetExpenseRatio { get; set; }
        public string AnnualHoldingsTurnover { get; set; }
        public string TotalAssets { get; set; }
        public string Average_Mkt_Cap_Mil { get; set; }
        public MarketCapitalization Market_Capitalisation { get; set; }
        public AssetAllocation Asset_Allocation { get; set; }
        public WorldRegions World_Regions { get; set; }
        public SectorWeights Sector_Weights { get; set; }
        public FixedIncome Fixed_Income { get; set; }
        public int Holdings_Count { get; set; }
        public Dictionary<string, HoldingData> Top_10_Holdings { get; set; }
        public Dictionary<string, HoldingData> Holdings { get; set; }
        public ValuationGrowthData Valuations_Growth { get; set; }
        public MorningStar MorningStar { get; set; }
        public Performance Performance { get; set; }
    }
}
