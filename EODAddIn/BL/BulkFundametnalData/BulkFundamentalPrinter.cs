using EOD.Model;
using EOD.Model.BulkFundamental;
using EOD.Model.OptionsData;
using EODAddIn.Program;
using EODAddIn.Utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using static EODAddIn.Utils.ExcelUtils;
using EODAddIn.BL.BulkFundametnalData;

namespace EODAddIn.BL.BulkFundamental
{
    public class BulkFundamentalPrinter
    {
        public static async void PrintBulkFundamentals(List<string> tickers)
        {
            try
            {
                SetNonInteractive();
                List<(string, string)> tickerAndExchanges = new List<(string, string)>();
                List<string> exchanges = new List<string>();
                List<string> TickerForBulk = new List<string>();
                Dictionary<string, BulkFundamentalData> data = null;
                foreach (string ticker in tickers)
                {
                    string[] subs = ticker.Split('.');
                    tickerAndExchanges.Add((ticker, subs[1]));
                    exchanges.Add(subs[1]);
                }
                foreach (string exchange in exchanges)
                {
                    foreach ((string, string) tickerAndExchange in tickerAndExchanges)
                    {
                        if (tickerAndExchange.Item2 == exchange)
                        {
                            TickerForBulk.Add(tickerAndExchange.Item1);
                        }
                    }
                    data = await BulkFundamentalAPI.GetBulkData(exchange, tickers, 0, 500);
                    for (int i = 0; i < data.Count; i++)
                    {
                        BulkFundamentalData symbol = data[i.ToString()];
                        string nameSheet = $"{symbol.General.Code},{symbol.General.Exchange}-Bulk fundamental";

                        Excel.Worksheet sh = AddSheet(nameSheet);

                        int row = 1;

                        row = PrintBulkFundamentalsGeneral(symbol, sh.Cells[row, 1]);
                        row++;
                        row = PrintBulkFundamentalsHighlights(symbol, sh.Cells[row, 1]);
                        row++;
                        row = PrintBulkFundamentalsValuation(symbol, sh.Cells[row, 1]);
                        row++;
                        row = PrintBulkFundamentalTechnicals(symbol, sh.Cells[row, 1]);
                        row++;
                        row = PrintBulkFundamentalEarnings(symbol, sh.Cells[row, 1]);
                        row++;
                        row = PrintBulkFundamentalFinancials(symbol, sh.Cells[row, 1]);
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                _xlsApp.Interactive = true;
            }
        }

        private static int PrintBulkFundamentalsGeneral(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "General";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Code";
            sh.Cells[row, column + 1] = data.General.Code;

            sh.Cells[row, column + 2] = "Type";
            sh.Cells[row, column + 3] = data.General.Type;
            row++;

            sh.Cells[row, column] = "Name";
            sh.Cells[row, column + 1] = data.General.Name;

            sh.Cells[row, column + 2] = "Exchange";
            sh.Cells[row, column + 3] = data.General.Exchange;
            row++;

            sh.Cells[row, column] = "Currency";
            sh.Cells[row, column + 1] = data.General.CurrencyCode;
            sh.Cells[row, column + 2] = data.General.CurrencySymbol;
            row++;

            sh.Cells[row, column] = "Country";
            sh.Cells[row, column + 1] = data.General.CountryName;
            sh.Cells[row, column + 2] = data.General.CountryISO;
            row++;

            sh.Cells[row, column] = "Sector";
            sh.Cells[row, column + 1] = data.General.Sector;
            row++;

            sh.Cells[row, column] = "Industry";
            sh.Cells[row, column + 1] = data.General.Industry;
            row++;

            sh.Cells[row, column] = "Employees";
            sh.Cells[row, column + 1] = data.General.FullTimeEmployees;
            row++;

            sh.Cells[row, column] = "Description";
            sh.Cells[row, column + 1] = data.General.Description;
            row++;

            return row;
        }

        private static int PrintBulkFundamentalsHighlights(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Highlights";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Market capitalization";
            sh.Cells[row, column + 1] = data.Highlights.MarketCapitalization;
            row++;

            sh.Cells[row, column] = "EBITDA";
            sh.Cells[row, column + 1] = data.Highlights.EBITDA;
            row++;

            sh.Cells[row, column] = "PE Ratio";
            sh.Cells[row, column + 1] = data.Highlights.PERatio;
            row++;

            sh.Cells[row, column] = "PEG Ratio";
            sh.Cells[row, column + 1] = data.Highlights.PEGRatio;
            row++;

            sh.Cells[row, column] = "WallStreet Target Price";
            sh.Cells[row, column + 1] = data.Highlights.WallStreetTargetPrice;
            row++;

            sh.Cells[row, column] = "Book Value";
            sh.Cells[row, column + 1] = data.Highlights.BookValue;
            row++;

            sh.Cells[row, column] = "Dividend Share";
            sh.Cells[row, column + 1] = data.Highlights.DividendShare;
            row++;

            sh.Cells[row, column] = "Dividend Yield";
            sh.Cells[row, column + 1] = data.Highlights.DividendYield;
            row++;

            sh.Cells[row, column] = "Earnings Share";
            sh.Cells[row, column + 1] = data.Highlights.EarningsShare;
            row++;

            sh.Cells[row, column] = "EPS Estimate Current Year";
            sh.Cells[row, column + 1] = data.Highlights.EPSEstimateCurrentYear;
            row++;

            sh.Cells[row, column] = "EPS Estimate Next Year";
            sh.Cells[row, column + 1] = data.Highlights.EPSEstimateNextYear;
            row++;

            sh.Cells[row, column] = "EPS Estimate Next Quarter";
            sh.Cells[row, column + 1] = data.Highlights.EPSEstimateNextQuarter;
            row++;

            sh.Cells[row, column] = "Most Recent Quarter";
            sh.Cells[row, column + 1] = data.Highlights.MostRecentQuarter;
            row++;

            sh.Cells[row, column] = "Profit Margin";
            sh.Cells[row, column + 1] = data.Highlights.ProfitMargin;
            row++;

            sh.Cells[row, column] = "Operating Margin TTM";
            sh.Cells[row, column + 1] = data.Highlights.OperatingMarginTTM;
            row++;

            sh.Cells[row, column] = "Return On Assets TTM";
            sh.Cells[row, column + 1] = data.Highlights.ReturnOnAssetsTTM;
            row++;

            sh.Cells[row, column] = "Return On Equity TTM";
            sh.Cells[row, column + 1] = data.Highlights.ReturnOnEquityTTM;
            row++;

            sh.Cells[row, column] = "Revenue TTM";
            sh.Cells[row, column + 1] = data.Highlights.RevenueTTM;
            row++;

            sh.Cells[row, column] = "Revenue PerShare TTM";
            sh.Cells[row, column + 1] = data.Highlights.RevenuePerShareTTM;
            row++;

            sh.Cells[row, column] = "Quarterly Revenue Growth YOY";
            sh.Cells[row, column + 1] = data.Highlights.QuarterlyRevenueGrowthYOY;
            row++;

            sh.Cells[row, column] = "Gross Profit TTM";
            sh.Cells[row, column + 1] = data.Highlights.GrossProfitTTM;
            row++;

            sh.Cells[row, column] = "Diluted Eps TTM";
            sh.Cells[row, column + 1] = data.Highlights.DilutedEpsTTM;
            row++;

            sh.Cells[row, column] = "Quarterly Earnings Growth YOY";
            sh.Cells[row, column + 1] = data.Highlights.QuarterlyEarningsGrowthYOY;
            row++;

            return row;
        }

        private static int PrintBulkFundamentalsValuation(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Valuation";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Trailing PE";
            sh.Cells[row, column + 1] = data.Valuation.TrailingPE;
            row++;

            sh.Cells[row, column] = "Forward PE";
            sh.Cells[row, column + 1] = data.Valuation.ForwardPE;
            row++;

            sh.Cells[row, column] = "Price Sales TTM";
            sh.Cells[row, column + 1] = data.Valuation.PriceSalesTTM;
            row++;

            sh.Cells[row, column] = "Price Book MRQ";
            sh.Cells[row, column + 1] = data.Valuation.PriceBookMRQ;
            row++;

            sh.Cells[row, column] = "Enterprise Value Revenue";
            sh.Cells[row, column + 1] = data.Valuation.EnterpriseValueRevenue;
            row++;

            sh.Cells[row, column] = "Enterprise Value Ebitda";
            sh.Cells[row, column + 1] = data.Valuation.EnterpriseValueEbitda;
            row++;

            return row;
        }

        private static int PrintBulkFundamentalTechnicals(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Valuation";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Beta";
            sh.Cells[row, column + 1] = data.Technicals.Beta;
            row++;

            sh.Cells[row, column] = "52 Week High";
            sh.Cells[row, column + 1] = data.Technicals.Week52High;
            row++;

            sh.Cells[row, column] = "52WeekLow";
            sh.Cells[row, column + 1] = data.Technicals.Week52Low;
            row++;

            sh.Cells[row, column] = "50 Day MA";
            sh.Cells[row, column + 1] = data.Technicals.Day50MA;
            row++;

            sh.Cells[row, column] = "200 Day MA";
            sh.Cells[row, column + 1] = data.Technicals.Day200MA;
            row++;

            sh.Cells[row, column] = "Shares Short";
            sh.Cells[row, column + 1] = data.Technicals.SharesShort;
            row++;

            sh.Cells[row, column] = "Shares Short Prior Month";
            sh.Cells[row, column + 1] = data.Technicals.SharesShortPriorMonth;
            row++;

            sh.Cells[row, column] = "Short Ratio";
            sh.Cells[row, column + 1] = data.Technicals.ShortRatio;
            row++;

            sh.Cells[row, column] = "Short Percent";
            sh.Cells[row, column + 1] = data.Technicals.ShortPercent;
            row++;

            return row;
        }

        private static int PrintBulkFundamentalSplitsDividents(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Splits Dividents";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Forward Annual Dividend Rate";
            sh.Cells[row, column + 1] = data.SplitsDividends.ForwardAnnualDividendRate;
            row++;

            sh.Cells[row, column] = "Forward Annual Dividend Yield";
            sh.Cells[row, column + 1] = data.SplitsDividends.ForwardAnnualDividendYield;
            row++;

            sh.Cells[row, column] = "Payout Ratio";
            sh.Cells[row, column + 1] = data.SplitsDividends.PayoutRatio;
            row++;

            sh.Cells[row, column] = "Dividend Date";
            sh.Cells[row, column + 1] = data.SplitsDividends.DividendDate;
            row++;

            sh.Cells[row, column] = "Ex Dividend Date";
            sh.Cells[row, column + 1] = data.SplitsDividends.ExDividendDate;
            row++;

            sh.Cells[row, column] = "Last Split Factor";
            sh.Cells[row, column + 1] = data.SplitsDividends.LastSplitFactor;
            row++;

            sh.Cells[row, column] = "Last Split Date";
            sh.Cells[row, column + 1] = data.SplitsDividends.LastSplitDate;
            row++;

            return row;
        }

        private static int PrintBulkFundamentalEarnings(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Earnings";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Date";
            sh.Cells[row, column + 1] = "EPS Actual";
            sh.Cells[row, column + 2] = "EPS Estimate";
            sh.Cells[row, column + 3] = "EPS Difference";
            sh.Cells[row, column + 4] = "Surprise Percent";
            row++;

            for (int i = 0; i < 4; i++)
            {
                string key = "Last_" + i.ToString();
                sh.Cells[row, column] = data.Earnings[key].Date;
                sh.Cells[row, column + 1] = data.Earnings[key].EpsActual;
                sh.Cells[row, column + 2] = data.Earnings[key].EpsEstimate;
                sh.Cells[row, column + 4] = data.Earnings[key].EpsDifference;
                sh.Cells[row, column + 5] = data.Earnings[key].SurprisePercent;
                row++;
            }

            return row;
        }

        private static int PrintBulkFundamentalFinancials(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Financials";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            row = PrintBulkFundamentalBalanceSheet(data, sh.Cells[row, 1]);

            row = PrintBulkFundamentalCashFlow(data, sh.Cells[row, 1]);

            row = PrintBulkFundamentalIncomeStatement(data, sh.Cells[row, 1]);

            return row;
        }
        private static int PrintBulkFundamentalBalanceSheet(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Balance Sheet";
            sh.Cells[row, column].Font.Bold = true;
            sh.Cells[row, column + 1] = data.Financials.Balance_Sheet.Currency_symbol;
            row++;

            EOD.Model.BulkFundamental.Balance_SheetData model = new EOD.Model.BulkFundamental.Balance_SheetData();
            PropertyInfo[] properties = model.GetType().GetProperties();

            foreach (var prop in properties)
            {
                sh.Cells[row, column] = prop.Name;
                column++;
            }
            row++;

            int countValues = 8;
            object[,] val = new object[countValues, properties.Length];
            int i = 0;

            var item = data.Financials.Balance_Sheet.Quarterly_last_0;
            val = FillRowBalanceSheet(val, item, properties, i);
            i++;
            item = data.Financials.Balance_Sheet.Quarterly_last_1;
            val = FillRowBalanceSheet(val, item, properties, i);
            i++;
            item = data.Financials.Balance_Sheet.Quarterly_last_2;
            val = FillRowBalanceSheet(val, item, properties, i);
            i++;
            item = data.Financials.Balance_Sheet.Quarterly_last_3;
            val = FillRowBalanceSheet(val, item, properties, i);
            i++;
            item = data.Financials.Balance_Sheet.Yearly_last_0;
            val = FillRowBalanceSheet(val, item, properties, i);
            i++;
            item = data.Financials.Balance_Sheet.Yearly_last_1;
            val = FillRowBalanceSheet(val, item, properties, i);
            i++;
            item = data.Financials.Balance_Sheet.Yearly_last_2;
            val = FillRowBalanceSheet(val, item, properties, i);
            i++;
            item = data.Financials.Balance_Sheet.Yearly_last_3;
            val = FillRowBalanceSheet(val, item, properties, i);

            sh.Range[sh.Cells[row, range.Column], sh.Cells[row + countValues - 1, properties.Length]].Value = val;

            return row + countValues;
        }
        private static int PrintBulkFundamentalCashFlow(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Cash Flow";
            sh.Cells[row, column].Font.Bold = true;
            sh.Cells[row, column + 1] = data.Financials.Cash_Flow.Currency_symbol;
            row++;

            EOD.Model.BulkFundamental.Cash_FlowData model = new EOD.Model.BulkFundamental.Cash_FlowData();
            PropertyInfo[] properties = model.GetType().GetProperties();

            foreach (var prop in properties)
            {
                sh.Cells[row, column] = prop.Name;
                column++;
            }
            row++;

            int countValues = 8;
            object[,] val = new object[countValues, properties.Length];
            int i = 0;

            var item = data.Financials.Cash_Flow.Quarterly_last_0;
            val = FillRowCashFlow(val, item, properties, i);
            i++;
            item = data.Financials.Cash_Flow.Quarterly_last_1;
            val = FillRowCashFlow(val, item, properties, i);
            i++;
            item = data.Financials.Cash_Flow.Quarterly_last_2;
            val = FillRowCashFlow(val, item, properties, i);
            i++;
            item = data.Financials.Cash_Flow.Quarterly_last_3;
            val = FillRowCashFlow(val, item, properties, i);
            i++;
            item = data.Financials.Cash_Flow.Yearly_last_0;
            val = FillRowCashFlow(val, item, properties, i);
            i++;
            item = data.Financials.Cash_Flow.Yearly_last_1;
            val = FillRowCashFlow(val, item, properties, i);
            i++;
            item = data.Financials.Cash_Flow.Yearly_last_2;
            val = FillRowCashFlow(val, item, properties, i);
            i++;
            item = data.Financials.Cash_Flow.Yearly_last_3;
            val = FillRowCashFlow(val, item, properties, i);

            sh.Range[sh.Cells[row, range.Column], sh.Cells[row + countValues - 1, properties.Length]].Value = val;

            return row + countValues;
        }
        private static int PrintBulkFundamentalIncomeStatement(BulkFundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Income Statement";
            sh.Cells[row, column].Font.Bold = true;
            sh.Cells[row, column + 1] = data.Financials.Cash_Flow.Currency_symbol;
            row++;

            EOD.Model.BulkFundamental.Income_StatementData model = new EOD.Model.BulkFundamental.Income_StatementData();
            PropertyInfo[] properties = model.GetType().GetProperties();

            foreach (var prop in properties)
            {
                sh.Cells[row, column] = prop.Name;
                column++;
            }
            row++;

            int countValues = 8;
            object[,] val = new object[countValues, properties.Length];
            int i = 0;

            var item = data.Financials.Income_Statement.Quarterly_last_0;
            val = FillRowIncomeStatement(val, item, properties, i);
            i++;
            item = data.Financials.Income_Statement.Quarterly_last_1;
            val = FillRowIncomeStatement(val, item, properties, i);
            i++;
            item = data.Financials.Income_Statement.Quarterly_last_2;
            val = FillRowIncomeStatement(val, item, properties, i);
            i++;
            item = data.Financials.Income_Statement.Quarterly_last_3;
            val = FillRowIncomeStatement(val, item, properties, i);
            i++;
            item = data.Financials.Income_Statement.Yearly_last_0;
            val = FillRowIncomeStatement(val, item, properties, i);
            i++;
            item = data.Financials.Income_Statement.Yearly_last_1;
            val = FillRowIncomeStatement(val, item, properties, i);
            i++;
            item = data.Financials.Income_Statement.Yearly_last_2;
            val = FillRowIncomeStatement(val, item, properties, i);
            i++;
            item = data.Financials.Income_Statement.Yearly_last_3;
            val = FillRowIncomeStatement(val, item, properties, i);

            sh.Range[sh.Cells[row, range.Column], sh.Cells[row + countValues - 1, properties.Length]].Value = val;

            return row + countValues;
        }
    }
}
