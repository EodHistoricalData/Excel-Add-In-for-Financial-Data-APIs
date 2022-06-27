using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class Income_StatementData
    {
        public DateTime? Date { get; set; }
        public DateTime? Filing_date { get; set; }
        public string Currency_symbol { get; set; }
        public double? ResearchDevelopment { get; set; }
        public double? EffectOfAccountingCharges { get; set; }
        public double? IncomeBeforeTax { get; set; }
        public double? MinorityInterest { get; set; }
        public double? NetIncome { get; set; }
        public double? SellingGeneralAdministrative { get; set; }
        public double? SellingAndMarketingExpenses { get; set; }
        public double? GrossProfit { get; set; }
        public double? ReconciledDepreciation { get; set; }
        public double? Ebit { get; set; }
        public double? Ebitda { get; set; }
        public double? DepreciationAndAmortization { get; set; }
        public double? NonOperatingIncomeNetOther { get; set; }
        public double? OperatingIncome { get; set; }
        public double? OtherOperatingExpenses { get; set; }
        public double? InterestExpense { get; set; }
        public double? TaxProvision { get; set; }
        public double? InterestIncome { get; set; }
        public double? NetInterestIncome { get; set; }
        public double? ExtraordinaryItems { get; set; }
        public double? NonRecurring { get; set; }
        public double? OtherItems { get; set; }
        public double? IncomeTaxExpense { get; set; }
        public double? TotalRevenue { get; set; }
        public double? TotalOperatingExpenses { get; set; }
        public double? CostOfRevenue { get; set; }
        public double? TotalOtherIncomeExpenseNet { get; set; }
        public double? DiscontinuedOperations { get; set; }
        public double? NetIncomeFromContinuingOps { get; set; }
        public double? NetIncomeApplicableToCommonShares { get; set; }
        public double? PreferredStockAndOtherAdjustments { get; set; }
    }
}
