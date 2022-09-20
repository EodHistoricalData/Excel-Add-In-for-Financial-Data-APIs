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
using EOD;
using EOD.Model.Fundamental;

namespace EODAddIn.BL.ETFPrinter
{
    public class ETFPrinter
    {
        /// <summary>
        /// Print ETF data
        /// </summary>
        /// <param name="data"></param>
        public static void PrintEtf(EOD.Model.Fundamental.FundamentalData data, string ticker)
        {
            try
            {
                SetNonInteractive();

                string nameSheet = $"{ticker}-Etfs";

                Excel.Worksheet sh = AddSheet(nameSheet);

                if (ExcelUtils.SheetExists(nameSheet))
                {
                    sh = Globals.ThisAddIn.Application.Worksheets[nameSheet];
                    int maxrow = ExcelUtils.RowsCount(sh);
                    sh.Range[$"A1:Z{maxrow}"].ClearContents();
                }
                else
                {
                    sh = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                    sh.Name = nameSheet;
                }

                int row = 1;
                int startGroup1 = 2;

                row = PrintEtfGeneral(data, sh.Cells[row, 1]);
                row++;

                sh.Rows[$"{startGroup1}:{row}"].Group();
                row++;

                startGroup1 = row + 1;
                row = PrintEtfTechnicals(data, sh.Cells[row, 1]);

                sh.Rows[$"{startGroup1}:{row}"].Group();
                row++;

                startGroup1 = row + 1;
                row = PrintEtfData(data, sh.Cells[row, 1]);

                sh.Rows[$"{startGroup1}:{row}"].Group();

                sh.Outline.AutomaticStyles = false;
                sh.Outline.SummaryRow = Excel.XlSummaryRow.xlSummaryAbove;

                sh.Outline.ShowLevels(1);
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
        /// <summary>
        /// Filling the worksheet with General data for ETF
        /// </summary>
        /// <param name="data"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int PrintEtfGeneral(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
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
            row++;

            sh.Cells[row, column] = "Exchange";
            sh.Cells[row, column + 1] = data.General.Exchange;
            row++;

            sh.Cells[row, column] = "Currency";
            sh.Cells[row, column + 1] = data.General.CurrencyCode;
            sh.Cells[row, column + 2] = data.General.CurrencySymbol;
            row++;

            sh.Cells[row, column] = "Country";
            sh.Cells[row, column + 1] = data.General.CountryName;
            row++;

            sh.Cells[row, column] = "Description";
            sh.Cells[row, column + 1] = data.General.Description;
            row++;

            sh.Cells[row, column] = "Category";
            sh.Cells[row, column + 1] = data.General.HomeCategory;

            return row;
        }
        /// <summary>
        /// Filling the worksheet with Technicals data for ETF
        /// </summary>
        /// <param name="data"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int PrintEtfTechnicals(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Technicals";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Beta";
            sh.Cells[row, column + 1] = data.Technicals.Beta;
            row++;

            sh.Cells[row, column] = "Week52High";
            sh.Cells[row, column + 1] = data.Technicals.Week52High;
            row++;

            sh.Cells[row, column] = "Week52Low";
            sh.Cells[row, column + 1] = data.Technicals.Week52Low;
            row++;

            sh.Cells[row, column] = "Day50MA ";
            sh.Cells[row, column + 1] = data.Technicals.Day50MA;
            row++;

            sh.Cells[row, column] = "Day200MA ";
            sh.Cells[row, column + 1] = data.Technicals.Day200MA;

            return row;
        }
        /// <summary>
        /// Filling the worksheet with ETF data
        /// </summary>
        /// <param name="data"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int PrintEtfData(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "ETF Data";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Company Name";
            sh.Cells[row, column + 1] = data.ETF_Data.Company_Name;
            row++;

            sh.Cells[row, column] = "Company URL";
            sh.Cells[row, column + 1] = data.ETF_Data.Company_URL;
            row++;

            sh.Cells[row, column] = "ETF URL";
            sh.Cells[row, column + 1] = data.ETF_Data.ETF_URL;
            row++;

            sh.Cells[row, column] = "Domicile";
            sh.Cells[row, column + 1] = data.ETF_Data.Domicile;
            row++;

            sh.Cells[row, column] = "Index Name";
            sh.Cells[row, column + 1] = data.ETF_Data.Index_Name;
            row++;

            sh.Cells[row, column] = "Yield";
            sh.Cells[row, column + 1] = data.ETF_Data.Yield;
            row++;

            sh.Cells[row, column] = "Dividend Paying Frequency";
            sh.Cells[row, column + 1] = data.ETF_Data.Dividend_Paying_Frequency;
            row++;

            sh.Cells[row, column] = "Inception Date";
            sh.Cells[row, column + 1] = data.ETF_Data.Inception_Date;
            row++;

            sh.Cells[row, column] = "Max Annual Mgmt Charge";
            sh.Cells[row, column + 1] = data.ETF_Data.Max_Annual_Mgmt_Charge;
            row++;

            sh.Cells[row, column] = "Ongoing Charge";
            sh.Cells[row, column + 1] = data.ETF_Data.Ongoing_Charge;
            row++;

            sh.Cells[row, column] = "Date Ongoing Charge";
            sh.Cells[row, column + 1] = data.ETF_Data.Date_Ongoing_Charge;
            row++;

            sh.Cells[row, column] = "Net Expense Ratio";
            sh.Cells[row, column + 1] = data.ETF_Data.NetExpenseRatio;
            row++;

            sh.Cells[row, column] = "Annual Holdings Turnover";
            sh.Cells[row, column + 1] = data.ETF_Data.AnnualHoldingsTurnover;
            row++;

            sh.Cells[row, column] = "Total Assets";
            sh.Cells[row, column + 1] = data.ETF_Data.TotalAssets;
            row++;

            sh.Cells[row, column] = "Average Mkt Cap Mil";
            sh.Cells[row, column + 1] = data.ETF_Data.Average_Mkt_Cap_Mil;
            row++;

            int startGroup2 = row + 1;
            row = PrintEtfMarketCap(data, sh.Cells[row, 1]);

            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;

            startGroup2 = row + 1;
            row = PrintEtfAssetAllocation(data, sh.Cells[row, 1]);

            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;

            startGroup2 = row + 1;
            row = PrintEtfWorldRegions(data, sh.Cells[row, 1]);

            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;

            startGroup2 = row + 1;
            row = PrintEtfSectorWeights(data, sh.Cells[row, 1]);

            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;

            startGroup2 = row + 1;
            row = PrintEtfFixedIncome(data, sh.Cells[row, 1]);

            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;

            startGroup2 = row + 1;
            row = PrintEtfHoldings(data.ETF_Data.Top_10_Holdings, data.ETF_Data.Holdings, sh.Cells[row, 1]);

            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;

            startGroup2 = row + 1;
            row = PrintEtfValuationsGrowth(data, sh.Cells[row, 1]);

            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;

            startGroup2 = row + 1;
            row = PrintEtfMorningStar(data, sh.Cells[row, 1]);

            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;

            startGroup2 = row + 1;
            row = PrintEtfPerformance(data, sh.Cells[row, 1]);

            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;

            startGroup2 = row + 1;
            sh.Cells[row, column] = "Holdings Count";
            sh.Cells[row, column + 1] = data.ETF_Data.Holdings_Count;

            row = PrintEtfHoldings(data.ETF_Data.Top_10_Holdings, data.ETF_Data.Holdings, sh.Cells[row, 1]);
            row++;

            sh.Rows[$"{startGroup2}:{row}"].Group();

            return row;
        }

        public static int PrintEtfMarketCap(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Market Capitalisation";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Mega";
            sh.Cells[row, column + 1] = "Big";
            sh.Cells[row, column + 2] = "Medium";
            sh.Cells[row, column + 3] = "Small";
            sh.Cells[row, column + 4] = "Micro";
            row++;

            sh.Cells[row, column] = data.ETF_Data.Market_Capitalisation.Mega;
            sh.Cells[row, column + 1] = data.ETF_Data.Market_Capitalisation.Big;
            sh.Cells[row, column + 2] = data.ETF_Data.Market_Capitalisation.Medium;
            sh.Cells[row, column + 3] = data.ETF_Data.Market_Capitalisation.Small;
            sh.Cells[row, column + 4] = data.ETF_Data.Market_Capitalisation.Micro;

            return row;
        }

        public static int PrintEtfAssetAllocation(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Asset Allocation";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column + 1] = "Long %";
            sh.Cells[row, column + 2] = "Short %";
            sh.Cells[row, column + 3] = "Net Assets %";
            row++;

            sh.Cells[row, column] = "Cash";
            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.Cash.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.Cash.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.Cash.NetAssetsPercent;
            row++;

            sh.Cells[row, column] = "Not Classified";

            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.NotClassified.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.NotClassified.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.NotClassified.NetAssetsPercent;
            row++;

            sh.Cells[row, column] = "Stock non-US";
            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.StocknonUS.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.StocknonUS.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.StocknonUS.NetAssetsPercent;
            row++;

            sh.Cells[row, column] = "Other";
            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.Other.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.Other.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.Other.NetAssetsPercent;
            row++;

            sh.Cells[row, column] = "Stock US";
            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.StockUS.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.StockUS.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.StockUS.NetAssetsPercent;
            row++;

            sh.Cells[row, column] = "Bond";
            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.Bond.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.Bond.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.Bond.NetAssetsPercent;

            return row;
        }

        public static int PrintEtfWorldRegions(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "World Regions";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column + 1] = "Equity %";
            sh.Cells[row, column + 2] = "Relative to Category";
            row++;

            sh.Cells[row, column] = "North America";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.NorthAmerica.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.NorthAmerica.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "United Kingdom";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.UnitedKingdom.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.UnitedKingdom.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Europe Developed";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.EuropeDeveloped.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.EuropeDeveloped.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Europe Emerging";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.EuropeEmerging.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.EuropeEmerging.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Africa/Middle East";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.AfricaMiddleEast.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.AfricaMiddleEast.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Japan";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.Japan.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.Japan.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Australasia";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.Australasia.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.Australasia.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Asia Developed";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.AsiaDeveloped.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.AsiaDeveloped.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Asia Emerging";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.AsiaEmerging.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.AsiaEmerging.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Latin America";
            sh.Cells[row, column + 1] = data.ETF_Data.World_Regions.LatinAmerica.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.World_Regions.LatinAmerica.RelativeToCategory;

            return row;
        }

        public static int PrintEtfSectorWeights(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Sector Weights";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column + 1] = "Equity %";
            sh.Cells[row, column + 2] = "Relative to Category";
            row++;

            sh.Cells[row, column] = "Basic Materials";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.BasicMaterials.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.BasicMaterials.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Consumer Cyclicals";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.ConsumerCyclisials.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.ConsumerCyclisials.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Financial Services";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.FinancialServices.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.FinancialServices.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Energy";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.Energy.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.Energy.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Industrials";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.Industrials.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.Industrials.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Technology";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.Technology.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.Technology.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Consumer Defensive";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.ConsumerDefencive.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.ConsumerDefencive.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Healthcare";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.Healthcare.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.Healthcare.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Utilities";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.Utilities.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.Utilities.RelativeToCategory;

            return row;
        }

        public static int PrintEtfFixedIncome(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Fixed Income";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column + 1] = "Fund %";
            sh.Cells[row, column + 2] = "Relative to Category";
            row++;

            sh.Cells[row, column] = "EffectiveDuration";
            sh.Cells[row+1, column] = "EffectiveMaturity";
            sh.Cells[row+2, column] = "EffectiveMaturity";
            sh.Cells[row+3, column] = "YieldToMaturity";
            foreach (KeyValuePair< string, EOD.Model.Fundamental.FixedIncomeData> item in data.ETF_Data.Fixed_Income)
            {
                sh.Cells[row, column + 1] = item.Value.FundPercent;
                sh.Cells[row, column + 2] = item.Value.RelativeToCategory;
                row++;
            }
            return row;
        }

        public static int PrintEtfValuationsGrowth(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Valuations Growth";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column + 1] = "Price/Prospective Earnings";
            sh.Cells[row, column + 2] = "Price/Book";
            sh.Cells[row, column + 3] = "Price/Sales";
            sh.Cells[row, column + 4] = "Price/Cash Flow";
            sh.Cells[row, column + 5] = "Dividend-Yield Factor";
            row++;

            sh.Cells[row, column] = "Valuations Rates Portfolio";
            sh.Cells[row, column + 1] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.PriceProspectiveEarnings;
            sh.Cells[row, column + 2] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.PriceBook;
            sh.Cells[row, column + 3] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.PriceSales;
            sh.Cells[row, column + 4] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.PriceCashFlow;
            sh.Cells[row, column + 5] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.DividendYieldFactor;
            row++;

            sh.Cells[row, column] = "Valuations Rates To Category";
            sh.Cells[row, column + 1] = data.ETF_Data.Valuations_Growth.Valuations_Rates_To_Category.PriceProspectiveEarnings;
            sh.Cells[row, column + 2] = data.ETF_Data.Valuations_Growth.Valuations_Rates_To_Category.PriceBook;
            sh.Cells[row, column + 3] = data.ETF_Data.Valuations_Growth.Valuations_Rates_To_Category.PriceSales;
            sh.Cells[row, column + 4] = data.ETF_Data.Valuations_Growth.Valuations_Rates_To_Category.PriceCashFlow;
            sh.Cells[row, column + 5] = data.ETF_Data.Valuations_Growth.Valuations_Rates_To_Category.DividendYieldFactor;
            row++;

            sh.Cells[row, column + 1] = "Long-Term Projected Earnings Growth";
            sh.Cells[row, column + 2] = "Historical Earnings Growth";
            sh.Cells[row, column + 3] = "Sales Growth";
            sh.Cells[row, column + 4] = "Cash-Flow Growth";
            sh.Cells[row, column + 5] = "Book-Value Growth";
            row++;

            sh.Cells[row, column] = "Growth Rates Portfolio";
            sh.Cells[row, column + 1] = data.ETF_Data.Valuations_Growth.Growth_Rates_Portfolio.LongTermProjectedEarningsGrowth;
            sh.Cells[row, column + 2] = data.ETF_Data.Valuations_Growth.Growth_Rates_Portfolio.HistoricalEarningsGrowth;
            sh.Cells[row, column + 3] = data.ETF_Data.Valuations_Growth.Growth_Rates_Portfolio.SalesGrowth;
            sh.Cells[row, column + 4] = data.ETF_Data.Valuations_Growth.Growth_Rates_Portfolio.CashFlowGrowth;
            sh.Cells[row, column + 5] = data.ETF_Data.Valuations_Growth.Growth_Rates_Portfolio.BookValueGrowth;
            row++;

            sh.Cells[row, column] = "Growth Rates To Category";
            sh.Cells[row, column + 1] = data.ETF_Data.Valuations_Growth.Growth_Rates_To_Category.LongTermProjectedEarningsGrowth;
            sh.Cells[row, column + 2] = data.ETF_Data.Valuations_Growth.Growth_Rates_To_Category.HistoricalEarningsGrowth;
            sh.Cells[row, column + 3] = data.ETF_Data.Valuations_Growth.Growth_Rates_To_Category.SalesGrowth;
            sh.Cells[row, column + 4] = data.ETF_Data.Valuations_Growth.Growth_Rates_To_Category.CashFlowGrowth;
            sh.Cells[row, column + 5] = data.ETF_Data.Valuations_Growth.Growth_Rates_To_Category.BookValueGrowth;

            return row;
        }

        public static int PrintEtfMorningStar(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Morning Star";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Ratio";
            sh.Cells[row, column + 1] = data.ETF_Data.MorningStar.Ratio;
            row++;

            sh.Cells[row, column] = "Category Benchmark";
            sh.Cells[row, column + 1] = data.ETF_Data.MorningStar.Category_Benchmark;
            row++;

            sh.Cells[row, column] = "Sustainability Ratio";
            sh.Cells[row, column + 1] = data.ETF_Data.MorningStar.Sustainability_Ratio;

            return row;
        }

        public static int PrintEtfPerformance(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Performance";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "1y_Volatility";
            sh.Cells[row, column + 1] = data.ETF_Data.Performance.Volatility1y;
            row++;

            sh.Cells[row, column] = "3y_Volatility";
            sh.Cells[row, column + 1] = data.ETF_Data.Performance.Volatility3y;
            row++;

            sh.Cells[row, column] = "3y_ExpReturn";
            sh.Cells[row, column + 1] = data.ETF_Data.Performance.ExpReturn3y;
            row++;

            sh.Cells[row, column] = "3y_SharpRatio";
            sh.Cells[row, column + 1] = data.ETF_Data.Performance.SharpRatio3y;
            row++;

            sh.Cells[row, column] = "Returns_YTD";
            sh.Cells[row, column + 1] = data.ETF_Data.Performance.Returns_YTD;
            row++;

            sh.Cells[row, column] = "Returns_1Y";
            sh.Cells[row, column + 1] = data.ETF_Data.Performance.Returns_1Y;
            row++;

            sh.Cells[row, column] = "Returns_3Y";
            sh.Cells[row, column + 1] = data.ETF_Data.Performance.Returns_3Y;
            row++;

            sh.Cells[row, column] = "Returns_5Y";
            sh.Cells[row, column + 1] = data.ETF_Data.Performance.Returns_5Y;
            row++;

            sh.Cells[row, column] = "Returns_10Y";
            sh.Cells[row, column + 1] = data.ETF_Data.Performance.Returns_10Y;

            return row;
        }
        private static int PrintEtfHoldings<T>(Dictionary<string, T> dataTable1,
    Dictionary<string, T> dataTable2,
    Excel.Range range,
    string dataTable1Name = "Top 10 holdings",
    string dataTable2Name = "Holdings")
    where T : class
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = $"{dataTable1Name}";
            sh.Cells[row, column].Font.Bold = true;
            row++;
            PrintTableHoldings(sh.Cells[row, column], dataTable1);
            row += dataTable1.Values.Count;

            sh.Cells[row, column] = $"{dataTable2Name}";
            sh.Cells[row, column].Font.Bold = true;
            row++;
            PrintTableHoldings(sh.Cells[row, column], dataTable2);
            row += dataTable2.Values.Count;

            return row;
        }
        /// <summary>
        /// Printing a list of holdings
        /// </summary>
        /// <typeparam name="T">Data type in the list</typeparam>
        /// <param name="range">target cell</param>
        /// <param name="data">List of holdings</param>
        private static void PrintTableHoldings<T>(Excel.Range range, Dictionary<string, T> data)
           where T : class, new()
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            T model = new T();
            PropertyInfo[] properties = model.GetType().GetProperties();

            foreach (var prop in properties)
            {
                sh.Cells[row, column] = prop.Name;
                column++;
            }
            row++;

            column = range.Column;
            int countValues = data.Values.Count;
            object[,] val = new object[countValues, properties.Length];
            int i = 0;
            foreach (T item in data.Values)
            {
                int j = 0;
                foreach (var prop in properties)
                {
                    val[i, j] = prop.GetValue(item);
                    j++;
                }
                row++;
                i++;
            }
            sh.Range[sh.Cells[range.Row, range.Column], sh.Cells[row - 2, column + properties.Length - 1]].Value = val;
        }

    }
}
