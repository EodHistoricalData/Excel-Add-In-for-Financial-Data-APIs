using EOD.Model;
using EOD.Model.BulkFundamental;
using EOD.Model.OptionsData;
using EODAddIn.Model;
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

namespace EODAddIn.BL.Screener
{
    public class ScreneerPrinter
    {
        static int screenerCounter = 1;
        static int rowGeneral = 3;
        static int rowGeneralTicker = 3;
        static int rowEarnings = 3;
        static int rowEarningsTicker = 3;
        static int rowBigTables = 3;
        static int rowBigTablesTicker = 3;
        private static Excel.Application _xlsApp = Globals.ThisAddIn.Application;
        private static bool CreateSheet = true;
        public static void PrintScreener(EOD.Model.Screener.StockMarkerScreener screener)
        {

            try
            {
                SetNonInteractive();
                Worksheet sh = new Worksheet();
                string nameSheet = "Screener " + Convert.ToString(screenerCounter);
                while (ExcelUtils.SheetExists(nameSheet))
                {
                    screenerCounter++;
                    nameSheet = "Screener " + Convert.ToString(screenerCounter);
                }
                sh = AddSheet(nameSheet);
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
                object[,] val = new object[screener.Data.Count + 1, 15];
                int i = 1;
                val[0, 0] = "Code";
                val[0, 1] = "Name";
                val[0, 2] = "Last day data date";
                val[0, 3] = "Adjusted Close";
                val[0, 4] = "Refund ID";
                val[0, 5] = "Exchange";
                val[0, 6] = "Currency symbol";
                val[0, 7] = "Market Capitalization";
                val[0, 8] = "Earnings Share";
                val[0, 9] = "Dividend yield";
                val[0, 10] = "Sector";
                val[0, 11] = "Industry";
                if (screener.Data.Count == 0)
                {
                    MessageBox.Show("No matches", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                foreach (var item in screener.Data)
                {
                    int j = 0;
                    val[i, j] = "'" + item.Code; j++;
                    val[i, j] = item.Name; j++;
                    val[i, j] = item.Last_Day_Data_Date; j++;
                    val[i, j] = item.Adjusted_Close; j++;
                    val[i, j] = item.Refund_1d; j++;
                    val[i, j] = item.Exchange; j++;
                    val[i, j] = item.Currency_Symbol; j++;
                    val[i, j] = item.Market_Capitalization; j++;
                    val[i, j] = item.Earnings_Share; j++;
                    val[i, j] = item.Dividend_Yield; j++;
                    val[i, j] = item.Sector; j++;
                    val[i, j] = item.Industry; j++;
                    i++;
                }
                sh.Range[sh.Cells[1, 1], sh.Cells[screener.Data.Count, 15]].Value = val;
                string endpoint = "L" + Convert.ToString(i - 1);
                MakeTable("A1", endpoint, sh, "Screener result", 9);
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
        private static bool CheckIsScreenerResult(Worksheet sh)
        {
            if (!(Globals.ThisAddIn.Application.ActiveSheet is Worksheet sh1))
            {
                MessageBox.Show(
                    "Choose worksheet with a screener results!",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
            string codeValue = Convert.ToString(sh.Cells[1, 1].Value);
            string exchangeValue = Convert.ToString(sh.Cells[1, 6].Value);
            if (String.IsNullOrEmpty(codeValue) | String.IsNullOrEmpty(exchangeValue))
            {
                MessageBox.Show(
                    "Choose worksheet with a screener results!",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
            return true;
        }
        private static Worksheet CreateGeneralWorksheet(string sheetName)
        {
            Worksheet sh = new Worksheet();
            sh = AddSheet("General data for " + sheetName);
            int columns = 1;
            sh.Cells[1, 1] = "Highlights";
            sh.Cells[1, 1].Font.Bold = true;
            sh.Cells[2, columns] = "Ticker"; columns++;
            sh.Cells[2, columns] = "Code"; columns++;
            sh.Cells[2, columns] = "Type"; columns++;
            sh.Cells[2, columns] = "Name"; columns++;
            sh.Cells[2, columns] = "Currency Code"; columns++;
            sh.Cells[2, columns] = "Currency Name"; columns++;
            sh.Cells[2, columns] = "Sector"; columns++;
            sh.Cells[2, columns] = "Industry"; columns++;
            sh.Cells[2, columns] = "Employees"; columns++;
            sh.Cells[2, columns] = "Description"; columns++;
            sh.Cells[2, columns] = "Exchange"; columns++;
            sh.Cells[2, columns] = "Market capitalization"; columns++;
            sh.Cells[2, columns] = "EBITDA"; columns++;
            sh.Cells[2, columns] = "PE Ratio"; columns++;
            sh.Cells[2, columns] = "PEG Ratio"; columns++;
            sh.Cells[2, columns] = "WallStreet Target Price"; columns++;
            sh.Cells[2, columns] = "Book Value"; columns++;
            sh.Cells[2, columns] = "Dividend Share"; columns++;
            sh.Cells[2, columns] = "Dividend Yield"; columns++;
            sh.Cells[2, columns] = "Earnings Share"; columns++;
            sh.Cells[2, columns] = "EPS Estimate Current Year"; columns++;
            sh.Cells[2, columns] = "EPS Estimate Next Year"; columns++;
            sh.Cells[2, columns] = "EPS Estimate Next Quarter"; columns++;
            sh.Cells[2, columns] = "Most Recent Quarter"; columns++;
            sh.Cells[2, columns] = "Profit Margin"; columns++;
            sh.Cells[2, columns] = "Operating Margin TTM"; columns++;
            sh.Cells[2, columns] = "Return On Assets TTM"; columns++;
            sh.Cells[2, columns] = "Return On Equity TTM"; columns++;
            sh.Cells[2, columns] = "Revenue TTM"; columns++;
            sh.Cells[2, columns] = "Revenue Per Share TTM"; columns++;
            sh.Cells[2, columns] = "Quarterly Revenue Growth YOY"; columns++;
            sh.Cells[2, columns] = "Gross Profit TTM"; columns++;
            sh.Cells[2, columns] = "Diluted Eps TTM"; columns++;
            sh.Cells[2, columns] = "Quarterly Earnings Growth YOY"; columns++;
            sh.Cells[2, columns] = "Trailing PE"; columns++;
            sh.Cells[2, columns] = "Forward PE"; columns++;
            sh.Cells[2, columns] = "Price Sales TTM"; columns++;
            sh.Cells[2, columns] = "Price Book MRQ"; columns++;
            sh.Cells[2, columns] = "Enterprise Value Revenue"; columns++;
            sh.Cells[2, columns] = "Enterprise Value Ebitda"; columns++;
            sh.Cells[2, columns] = "Beta"; columns++;
            sh.Cells[2, columns] = "52 Week High"; columns++;
            sh.Cells[2, columns] = "52 Week Low"; columns++;
            sh.Cells[2, columns] = "50 Day MA"; columns++;
            sh.Cells[2, columns] = "200 Day MA"; columns++;
            sh.Cells[2, columns] = "Shares Short"; columns++;
            sh.Cells[2, columns] = "Shares Short Prior Month"; columns++;
            sh.Cells[2, columns] = "Short Ratio"; columns++;
            sh.Cells[2, columns] = "Short Percent"; columns++;
            return sh;
        }
        private static Worksheet CreateEarningsWorksheet(string sheetName)
        {
            Worksheet sh = new Worksheet();
            sh = AddSheet("Earnings data for " + sheetName);
            int column = 1;
            int row = 1;
            sh.Cells[row, column] = "Earnings";
            sh.Cells[row, column].Font.Bold = true;
            row++;
            sh.Cells[row, column] = "Ticker";
            sh.Cells[row, column + 1] = "Date";
            sh.Cells[row, column + 2] = "EPS Actual";
            sh.Cells[row, column + 3] = "EPS Estimate";
            sh.Cells[row, column + 4] = "EPS Difference";
            sh.Cells[row, column + 5] = "Surprise Percent";
            row++;
            return sh;
        }
        private static Worksheet CreateBalanceWorksheet(string sheetName)
        {
            Worksheet sh = new Worksheet();
            sh = AddSheet("Balance data for " + sheetName);
            int column = 1;
            int row = 1;
            sh.Cells[row, column] = "Balance Sheet";
            sh.Cells[row, column].Font.Bold = true;
            return sh;
        }
        private static Worksheet CreateCashFlowWorksheet(string sheetName)
        {
            Worksheet sh = new Worksheet();
            sh = AddSheet("Cash Flow data for " + sheetName);
            int column = 1;
            int row = 1;
            sh.Cells[row, column] = "Cash FLow Sheet";
            sh.Cells[row, column].Font.Bold = true;
            return sh;
        }
        private static Worksheet CreateIncomeStatementWorksheet(string sheetName)
        {
            Worksheet sh = new Worksheet();
            sh = AddSheet("Income data for " + sheetName);
            int column = 1;
            int row = 1;
            sh.Cells[row, column] = "Income Statement Sheet";
            sh.Cells[row, column].Font.Bold = true;
            return sh;
        }
        private static List<(string, string)> GetTickersAndExchangesFromScreener(Worksheet sh)
        {
            List<(string, string)> tickers = new List<(string, string)>();
            int i = 2;
            string ticker = Convert.ToString(sh.Cells[i, 1].Value) + "." + Convert.ToString(sh.Cells[i, 6].Value);
            string excahnge = Convert.ToString(sh.Cells[i, 6].Value);
            while (ticker != ".")
            {
                i++;
                tickers.Add((ticker, excahnge));
                ticker = Convert.ToString(sh.Cells[i, 1].Value) + "." + Convert.ToString(sh.Cells[i, 6].Value);
                excahnge = Convert.ToString(sh.Cells[i, 6].Value);
            }
            return tickers;
        }
        private static List<string> GetExchangesFromScreener(Worksheet sh)
        {
            List<string> exchanges = new List<string>();
            int i = 2;
            string cellValue = Convert.ToString(sh.Cells[i, 6].Value);
            while (cellValue != null)
            {
                i++;
                if (!exchanges.Contains(cellValue))
                {
                    exchanges.Add(cellValue);
                }
                cellValue = Convert.ToString(sh.Cells[i, 6].Value);

            }
            return exchanges;
        }
       
        public static async void PrintScreenerBulk()
        {
            Worksheet shGeneral = new Worksheet();
            Worksheet shEarnings = new Worksheet();
            Worksheet shBalance = new Worksheet();
            Worksheet shCashFlow = new Worksheet();
            Worksheet shIncomeStatement = new Worksheet();
            shGeneral = Globals.ThisAddIn.Application.ActiveSheet;
            if (!CheckIsScreenerResult(shGeneral))
            {
                return;
            }
            string screenerSheetName = shGeneral.Name;

            List<(string, string)> tickersAndExchanges = GetTickersAndExchangesFromScreener(shGeneral);
            List<string> exchanges = GetExchangesFromScreener(shGeneral);
            shGeneral = CreateGeneralWorksheet(screenerSheetName);
            shEarnings = CreateEarningsWorksheet(screenerSheetName);
            shBalance = CreateBalanceWorksheet(screenerSheetName);
            shCashFlow = CreateCashFlowWorksheet(screenerSheetName);
            shIncomeStatement = CreateIncomeStatementWorksheet(screenerSheetName);
            List<string> tickers = new List<string>();
            int offset = 0;
            Dictionary<string, BulkFundamentalData> res;
            int tickersCount = 0;
            foreach (string exchange in exchanges)
            {
                foreach ((string, string) tickerAndExchange in tickersAndExchanges)
                {
                    if (tickerAndExchange.Item2 == exchange)
                    {
                        tickers.Add(tickerAndExchange.Item1);
                    }
                }
                tickersCount = tickers.Count;
                while (tickersCount > 500)
                {
                    res = await GetBulkFundamental.GetBulkData(exchange, tickers, offset, 500);
                    offset += 500;
                    tickersCount--;
                }
                res = await GetBulkFundamental.GetBulkData(exchange, tickers, offset, 500);
                PrintBulkFundamentalForScreener(res, tickers, shGeneral, shEarnings, shBalance, shCashFlow, shIncomeStatement);
                tickers.Clear();
            }
            MakeTable("A2", "AW" + Convert.ToString(rowGeneral), shGeneral, "Highlights", 9);
            MakeTable("A2", "F" + Convert.ToString(rowEarningsTicker), shEarnings, "Earnings", 9);
            MakeTable("A2", "AB" + Convert.ToString(rowBigTables), shBalance, "Balance", 9);
            MakeTable("A2", "S" + Convert.ToString(rowBigTables), shCashFlow, "Cash Flow", 9);
            MakeTable("A2", "Y" + Convert.ToString(rowBigTables), shIncomeStatement, "Income Statement", 9);
            rowGeneralTicker = 3;
            rowEarningsTicker = 3;
            rowGeneral = 3;
            rowEarnings = 3;
            rowBigTables = 3;
            rowBigTablesTicker = 3;
        }
        private static void PrintBulkFundamentalForScreener(Dictionary<string, BulkFundamentalData> data, List<string> tickers, Worksheet shGeneral, Worksheet shEarnings, Worksheet shBalance, Worksheet shCashFlow, Worksheet shIncomeStatement)
        {
            try
            {
                SetNonInteractive();
                for (int i = 0; i < tickers.Count; i++)
                {
                    shGeneral.Cells[rowGeneralTicker, 1] = tickers[i];
                    rowGeneralTicker++;
                    for (int j = 0; j < 4; j++)
                    {
                        shEarnings.Cells[rowEarningsTicker, 1] = tickers[i]; rowEarningsTicker++;
                    }
                    for (int k = 0; k < 8; k++)
                    {
                        shBalance.Cells[rowBigTablesTicker, 1] = tickers[i];
                        shCashFlow.Cells[rowBigTablesTicker, 1] = tickers[i];
                        shIncomeStatement.Cells[rowBigTablesTicker, 1] = tickers[i];
                        rowBigTablesTicker++;
                    }
                }
                int columns;
                for (int i = 0; i < data.Count; i++)
                {
                    columns = 2;
                    BulkFundamentalData symbol = data[i.ToString()];
                    columns = PrintScreenerBulkGeneral(symbol, shGeneral, columns);
                    columns = PrintScreenerBulkHighlights(symbol, shGeneral, columns);
                    columns = PrintScreenerBulkValuation(symbol, shGeneral, columns);
                    columns = PrintScreenerBulkTechnicals(symbol, shGeneral, columns);
                    PrintEarningsResultForScreener(symbol, shEarnings);
                    PrintBalanceResultForScreener(symbol, shBalance);
                    PrintCashFlowForScreener(symbol, shCashFlow);
                    PrintIncomeStatementForScreener(symbol, shIncomeStatement);
                    rowGeneral++;
                    rowBigTables += 8;
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
        private static int PrintScreenerBulkGeneral(BulkFundamentalData data, Worksheet sh, int column)
        {
            sh.Cells[rowGeneral, column] = data.General.Code; column++;
            sh.Cells[rowGeneral, column] = data.General.Type; column++;
            sh.Cells[rowGeneral, column] = data.General.Name; column++;
            sh.Cells[rowGeneral, column] = data.General.CurrencyCode; column++;
            sh.Cells[rowGeneral, column] = data.General.CountryName; column++;
            sh.Cells[rowGeneral, column] = data.General.Sector; column++;
            sh.Cells[rowGeneral, column] = data.General.Industry; column++;
            sh.Cells[rowGeneral, column] = data.General.FullTimeEmployees; column++;
            sh.Cells[rowGeneral, column] = data.General.Description; column++;
            sh.Cells[rowGeneral, column] = data.General.Exchange; column++;
            return column;
        }
        private static int PrintScreenerBulkHighlights(BulkFundamentalData data, Worksheet sh, int column)
        {
            sh.Cells[rowGeneral, column] = data.Highlights.MarketCapitalization;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.EBITDA;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.PERatio;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.PEGRatio;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.WallStreetTargetPrice;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.BookValue;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.DividendShare;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.DividendYield;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.EarningsShare;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.EPSEstimateCurrentYear;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.EPSEstimateNextYear;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.EPSEstimateNextQuarter;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.MostRecentQuarter;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.ProfitMargin;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.OperatingMarginTTM;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.ReturnOnAssetsTTM;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.ReturnOnEquityTTM;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.RevenueTTM;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.RevenuePerShareTTM;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.QuarterlyRevenueGrowthYOY;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.GrossProfitTTM;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.DilutedEpsTTM;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.QuarterlyEarningsGrowthYOY;
            column++;
            return column;
        }
        private static int PrintScreenerBulkValuation(BulkFundamentalData data, Worksheet sh, int column)
        {
            sh.Cells[1, column] = "Valuation";
            sh.Cells[1, column].Font.Bold = true;
            sh.Cells[rowGeneral, column] = data.Valuation.TrailingPE;
            column++;
            sh.Cells[rowGeneral, column] = data.Valuation.ForwardPE;
            column++;
            sh.Cells[rowGeneral, column] = data.Valuation.PriceSalesTTM;
            column++;
            sh.Cells[rowGeneral, column] = data.Valuation.PriceBookMRQ;
            column++;
            sh.Cells[rowGeneral, column] = data.Valuation.EnterpriseValueRevenue;
            column++;
            sh.Cells[rowGeneral, column] = data.Valuation.EnterpriseValueEbitda;
            column++;
            return column;
        }
        private static int PrintScreenerBulkTechnicals(BulkFundamentalData data, Worksheet sh, int column)
        {
            sh.Cells[1, column] = "Technicals";
            sh.Cells[column].Font.Bold = true;
            sh.Cells[rowGeneral, column] = data.Technicals.Beta;
            column++;
            sh.Cells[rowGeneral, column] = data.Technicals.Week52High;
            column++;
            sh.Cells[rowGeneral, column] = data.Technicals.Week52Low;
            column++;
            sh.Cells[rowGeneral, column] = data.Technicals.Day50MA;
            column++;
            sh.Cells[rowGeneral, column] = data.Technicals.Day200MA;
            column++;
            sh.Cells[rowGeneral, column] = data.Technicals.SharesShort;
            column++;
            sh.Cells[rowGeneral, column] = data.Technicals.SharesShortPriorMonth;
            column++;
            sh.Cells[rowGeneral, column] = data.Technicals.ShortPercent;
            column++;
            sh.Cells[rowGeneral, column] = data.Technicals.ShortRatio;
            column++;
            return column;
        }
        private static void PrintEarningsResultForScreener(BulkFundamentalData data, Worksheet sh)
        {
            int column = 2;
            for (int i = 0; i < 4; i++)
            {
                string key = "Last_" + i.ToString();
                sh.Cells[rowEarnings, column] = data.Earnings[key].Date;
                sh.Cells[rowEarnings, column + 1] = data.Earnings[key].EpsActual;
                sh.Cells[rowEarnings, column + 2] = data.Earnings[key].EpsEstimate;
                sh.Cells[rowEarnings, column + 3] = data.Earnings[key].EpsDifference;
                sh.Cells[rowEarnings, column + 4] = data.Earnings[key].SurprisePercent;
                rowEarnings++;
            }
        }
        private static void PrintBalanceResultForScreener(BulkFundamentalData data, Worksheet sh)
        {
            int row = 1;
            int column = 2;
            sh.Cells[row, column + 1] = data.Financials.Balance_Sheet.Currency_symbol; row++;
            EOD.Model.BulkFundamental.Balance_SheetData model = new EOD.Model.BulkFundamental.Balance_SheetData();
            PropertyInfo[] properties = model.GetType().GetProperties();
            sh.Cells[2, 1] = "Tickers";
            foreach (var prop in properties)
            {
                sh.Cells[row, column] = prop.Name;
                column++;
            }
            row = rowBigTables;
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

            sh.Range[sh.Cells[rowBigTables, 2], sh.Cells[rowBigTables + countValues - 1, properties.Length + 1]].Value = val;
            rowBigTables = row;
        }
        private static void PrintCashFlowForScreener(BulkFundamentalData data, Worksheet sh)
        {
            int row = 1;
            int column = 2;
            sh.Cells[row, column + 1] = data.Financials.Cash_Flow.Currency_symbol; row++;
            EOD.Model.BulkFundamental.Cash_FlowData model = new EOD.Model.BulkFundamental.Cash_FlowData();
            PropertyInfo[] properties = model.GetType().GetProperties();
            sh.Cells[2, 1] = "Tickers";
            foreach (var prop in properties)
            {
                sh.Cells[row, column] = prop.Name;
                column++;
            }
            row = rowBigTables;
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

            sh.Range[sh.Cells[rowBigTables, 2], sh.Cells[rowBigTables + countValues - 1, properties.Length + 1]].Value = val;
            rowBigTables = row;
        }
        private static void PrintIncomeStatementForScreener(BulkFundamentalData data, Worksheet sh)
        {
            int row = 1;
            int column = 2;
            sh.Cells[row, column + 1] = data.Financials.Income_Statement.Currency_symbol; row++;
            EOD.Model.BulkFundamental.Income_StatementData model = new EOD.Model.BulkFundamental.Income_StatementData();
            PropertyInfo[] properties = model.GetType().GetProperties();
            sh.Cells[2, 1] = "Tickers";
            foreach (var prop in properties)
            {
                sh.Cells[row, column] = prop.Name;
                column++;
            }
            row = rowBigTables;
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

            sh.Range[sh.Cells[rowBigTables, 2], sh.Cells[rowBigTables + countValues - 1, properties.Length + 1]].Value = val;
            rowBigTables = row;
        }
        public static void PrintScreenerHistorical(DateTime from, DateTime to, string period)
        {
            try
            {
                SetNonInteractive();
                int row = 3;
                Worksheet sh = new Worksheet();
                sh = Globals.ThisAddIn.Application.ActiveSheet;
                if (!CheckIsScreenerResult(sh))
                {
                    return;
                }
                string screenerSheetName = sh.Name;

                List<(string, string)> tickers = GetTickersAndExchangesFromScreener(sh);
                sh = CreateScreenerHictoricalWorksheet(sh.Name);
                foreach ((string, string) ticker in tickers)
                {
                    List<EndOfDay> res = APIEOD.GetEOD(ticker.Item1, from, to, period);
                    foreach (EndOfDay item in res)
                    {
                        sh.Cells[row, 1] = ticker.Item1;
                        sh.Cells[row, 2] = item.Date;
                        sh.Cells[row, 3] = item.Open;
                        sh.Cells[row, 4] = item.High;
                        sh.Cells[row, 5] = item.Low;
                        sh.Cells[row, 6] = item.Close;
                        sh.Cells[row, 7] = item.Adjusted_open;
                        sh.Cells[row, 8] = item.Adjusted_high;
                        sh.Cells[row, 9] = item.Adjusted_low;
                        sh.Cells[row, 10] = item.Adjusted_close;
                        sh.Cells[row, 11] = item.Volume;
                        row++;
                    }
                }
                MakeTable("A2", "K" + Convert.ToString(row), sh, sh.Name, 9);
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
        private static Worksheet CreateScreenerHictoricalWorksheet(string sheetName)
        {
            Worksheet sh = new Worksheet();
            sh = AddSheet("Historical data for " + sheetName);
            int column = 1;
            int row = 1;
            sh.Cells[row, column] = "Hictorical data";
            sh.Cells[row, column].Font.Bold = true; row++;
            sh.Cells[row, column] = "Ticker"; column++;
            sh.Cells[row, column] = "Date"; column++;
            sh.Cells[row, column] = "Open"; column++;
            sh.Cells[row, column] = "High"; column++;
            sh.Cells[row, column] = "Low"; column++;
            sh.Cells[row, column] = "Close"; column++;
            sh.Cells[row, column] = "Adjusted open"; column++;
            sh.Cells[row, column] = "Adjusted high"; column++;
            sh.Cells[row, column] = "Adjusted low"; column++;
            sh.Cells[row, column] = "Adjusted close"; column++;
            sh.Cells[row, column] = "Volume"; column++;
            return sh;
        }
        public static void PrintScreenerIntraday(DateTime from, DateTime to, string interval)
        {
            try
            {
                SetNonInteractive();
                int row = 3;
                Worksheet sh = new Worksheet();
                sh = Globals.ThisAddIn.Application.ActiveSheet;
                if (!CheckIsScreenerResult(sh))
                {
                    return;
                }
                string screenerSheetName = sh.Name;

                List<(string, string)> tickers = GetTickersAndExchangesFromScreener(sh);
                sh = CreateScreenerIntradayWorksheet(sh.Name);
                foreach ((string, string) ticker in tickers)
                {
                    List<Model.Intraday> res = APIEOD.GetIntraday(ticker.Item1, from, to, interval);
                    foreach (Intraday item in res)
                    {
                        sh.Cells[row, 1] = ticker.Item1;
                        sh.Cells[row, 2] = item.DateTime;
                        sh.Cells[row, 3] = item.Gmtoffset;
                        sh.Cells[row, 4] = item.DateTime;
                        sh.Cells[row, 5] = item.Open;
                        sh.Cells[row, 6] = item.High;
                        sh.Cells[row, 7] = item.Low;
                        sh.Cells[row, 8] = item.Close;
                        sh.Cells[row, 9] = item.Volume;
                        sh.Cells[row, 10] = item.Timestamp;
                        row++;
                    }
                }
                MakeTable("A2", "J" + Convert.ToString(row), sh, sh.Name, 9);
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
        private static Worksheet CreateScreenerIntradayWorksheet(string sheetName)
        {
            Worksheet sh = new Worksheet();
            sh = AddSheet("Intraday data for " + sheetName);
            int column = 1;
            int row = 1;
            sh.Cells[row, column] = "Intraday data";
            sh.Cells[row, column].Font.Bold = true; row++;
            sh.Cells[row, column] = "Ticker"; column++;
            sh.Cells[row, column] = "DateTime"; column++;
            sh.Cells[row, column] = "Gmtoffset"; column++;
            sh.Cells[row, column] = "DateTime"; column++;
            sh.Cells[row, column] = "Open"; column++;
            sh.Cells[row, column] = "High"; column++;
            sh.Cells[row, column] = "Low"; column++;
            sh.Cells[row, column] = "Close"; column++;
            sh.Cells[row, column] = "Volume"; column++;
            sh.Cells[row, column] = "Timestamp"; column++;
            return sh;
        }
    }
}
