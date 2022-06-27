using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class Cash_FlowData
    {
        public DateTime? Date { get; set; }
        public DateTime? Filing_date { get; set; }
        public string Currency_symbol { get; set; }
        public double? Investments { get; set; }
        public double? ChangeToLiabilities { get; set; }
        public double? TotalCashflowsFromInvestingActivities { get; set; }
        public double? NetBorrowings { get; set; }
        public double? TotalCashFromFinancingActivities { get; set; }
        public double? HangeToOperatingActivities { get; set; }
        public double? NetIncome { get; set; }
        public double? HangeInCash { get; set; }
        public double? BeginPeriodCashFlow { get; set; }
        public double? EndPeriodCashFlow { get; set; }
        public double? TotalCashFromOperatingActivities { get; set; }
        public double? Depreciation { get; set; }
        public double? OtherCashflowsFromInvestingActivities { get; set; }
        public double? DividendsPaid { get; set; }
        public double? ChangeToInventory { get; set; }
        public double? ChangeToAccountReceivables { get; set; }
        public double? SalePurchaseOfStock { get; set; }
        public double? OtherCashflowsFromFinancingActivities { get; set; }
        public double? ChangeToNetincome { get; set; }
        public double? CapitalExpenditures { get; set; }
        public double? ChangeReceivables { get; set; }
        public double? CashFlowsOtherOperating { get; set; }
        public double? ExchangeRateChanges { get; set; }
        public double? CashAndCashEquivalentsChanges { get; set; }
        public double? ChangeInWorkingCapital { get; set; }
        public double? OtherNonCashItems { get; set; }
        public double? FreeCashFlow { get; set; }

    }
}
