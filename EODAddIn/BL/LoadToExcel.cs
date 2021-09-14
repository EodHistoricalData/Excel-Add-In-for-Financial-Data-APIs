using EODAddIn.Model;
using EODAddIn.Utils;

using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.BL
{
    public class LoadToExcel
    {
        public static void LoadEndOfDay(List<EndOfDay> endOfDays, string ticker, string period)
        {
            string nameSheet = $"{ticker}-{period}";
            Excel.Worksheet worksheet;
            if (ExcelUtils.SheetExists(nameSheet))
            {
                worksheet = Globals.ThisAddIn.Application.Worksheets[nameSheet];
                worksheet.Cells.Clear();
            }
            else
            {
                worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                worksheet.Name = nameSheet;
            }

            int r = 2;
            worksheet.Cells[r, 1] = "Date";
            worksheet.Cells[r, 2] = "Open";
            worksheet.Cells[r, 3] = "High";
            worksheet.Cells[r, 4] = "Low";
            worksheet.Cells[r, 5] = "Close";
            worksheet.Cells[r, 6] = "Adjusted_open";
            worksheet.Cells[r, 7] = "Adjusted_high";
            worksheet.Cells[r, 8] = "Adjusted_lowe";
            worksheet.Cells[r, 9] = "Adjusted_close";
            worksheet.Cells[r, 10] = "Volume";

            foreach (EndOfDay item in endOfDays)
            {
                r++;
                worksheet.Cells[r, 1] = item.Date;
                worksheet.Cells[r, 2] = item.Open;
                worksheet.Cells[r, 3] = item.High;
                worksheet.Cells[r, 4] = item.Low;
                worksheet.Cells[r, 5] = item.Close;
                worksheet.Cells[r, 6] = item.Adjusted_open;
                worksheet.Cells[r, 7] = item.Adjusted_high;
                worksheet.Cells[r, 8] = item.Adjusted_low;
                worksheet.Cells[r, 9] = item.Adjusted_close;
                worksheet.Cells[r, 10] = item.Volume;
            }
        }

        public static void LoadFundamental(FundamentalData data)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;

            int r = 1;
            int startGroup1 = 2;
            // General
            sh.Cells[r, 1] = "General";
            sh.Cells[r, 1].Font.Bold = true;
            r++;


            sh.Cells[r, 1] = "Code";
            sh.Cells[r, 2] = data.General.Code;

            sh.Cells[r, 3] = "Type";
            sh.Cells[r, 4] = data.General.Type;

            r++;

            sh.Cells[r, 1] = "Name";
            sh.Cells[r, 2] = data.General.Name;

            sh.Cells[r, 3] = "Exchange";
            sh.Cells[r, 4] = data.General.Exchange;

            r++;

            sh.Cells[r, 1] = "Currency";
            sh.Cells[r, 2] = data.General.CurrencyCode;
            sh.Cells[r, 3] = data.General.CurrencySymbol;
            r++;

            sh.Cells[r, 1] = "Sector";
            sh.Cells[r, 2] = data.General.Sector;
            r++;

            sh.Cells[r, 1] = "Industry";
            sh.Cells[r, 2] = data.General.Industry;
            r++;

            sh.Cells[r, 1] = "Employees";
            sh.Cells[r, 2] = data.General.FullTimeEmployees;
            r++;

            sh.Cells[r, 1] = "Description";
            sh.Cells[r, 2] = data.General.Description;

            r++;

            sh.Rows[$"{startGroup1}:{r}"].Group();

            r++;

            // Highlights
            sh.Cells[r, 1] = "Highlights";
            sh.Cells[r, 1].Font.Bold = true;
            r++;
            startGroup1 = r;
            sh.Cells[r, 1] = "Market Cap";
            sh.Cells[r, 2] = data.Highlights.MarketCapitalization;

            sh.Cells[r, 3] = "EBITDA";
            sh.Cells[r, 4] = data.Highlights.EBITDA;

            r++;

            sh.Cells[r, 1] = "PE Ratio";
            sh.Cells[r, 2] = data.Highlights.PERatio;

            sh.Cells[r, 3] = "PEG Ratio";
            sh.Cells[r, 4] = data.Highlights.PEGRatio;

            r++;

            sh.Cells[r, 1] = "Earning Share";
            sh.Cells[r, 2] = data.Highlights.EarningsShare;

            r++;

            sh.Cells[r, 1] = "Dividend Share";
            sh.Cells[r, 2] = data.Highlights.DividendShare;

            sh.Cells[r, 3] = "Dividend Yield";
            sh.Cells[r, 4] = data.Highlights.DividendYield;

            r++;

            sh.Cells[r, 1] = "EPS Estimate"; r++;

            sh.Cells[r, 1] = "Current Year";
            sh.Cells[r, 2] = data.Highlights.EPSEstimateCurrentYear;

            r++;

            sh.Cells[r, 1] = "Next Year";
            sh.Cells[r, 2] = data.Highlights.EPSEstimateNextYear;

            r++;

            sh.Cells[r, 1] = "Next Quarter";
            sh.Cells[r, 2] = data.Highlights.EPSEstimateNextQuarter;

            sh.Rows[$"{startGroup1}:{r}"].Group();

            r++;

            // Balance Sheet
            sh.Cells[r, 1] = "Balance Sheet";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            startGroup1 = r;


            sh.Cells[r, 1] = "Quarterly";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            int startGroup2 = r;
            Balance_SheetData balance_SheetData = new Balance_SheetData();

            int c = 1;
            System.Reflection.PropertyInfo[] properties = balance_SheetData.GetType().GetProperties();
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            int countValues = data.Financials.Balance_Sheet.Quarterly.Values.Count;
            object[,] val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;

                foreach (Balance_SheetData item in data.Financials.Balance_Sheet.Quarterly.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
            r += countValues;

            sh.Rows[$"{startGroup2}:{r}"].Group();

            r++;
            sh.Cells[r, 1] = "Yearly";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            startGroup2 = r;
            c = 1;
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            countValues = data.Financials.Balance_Sheet.Yearly.Values.Count;
            val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;
                foreach (Balance_SheetData item in data.Financials.Balance_Sheet.Yearly.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
            r += countValues;
            sh.Rows[$"{startGroup2}:{r}"].Group();

            r++;
            sh.Rows[$"{startGroup1}:{r}"].Group();
            r++;
            // Income_Statement
            sh.Cells[r, 1] = "Income Statement";
            sh.Cells[r, 1].Font.Bold = true;

            r++;

            startGroup1 = r;
            sh.Cells[r, 1] = "Quarterly";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            startGroup2 = r;
            Income_StatementData income_StatementData = new Income_StatementData();

            c = 1;
            properties = income_StatementData.GetType().GetProperties();
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            countValues = data.Financials.Income_Statement.Quarterly.Values.Count;
            val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;

                foreach (Income_StatementData item in data.Financials.Income_Statement.Quarterly.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
            r += countValues;

            sh.Rows[$"{startGroup2}:{r}"].Group();

            r++;
            sh.Cells[r, 1] = "Yearly";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            startGroup2 = r;
            c = 1;
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            countValues = data.Financials.Income_Statement.Yearly.Values.Count;
            val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;
                foreach (Income_StatementData item in data.Financials.Income_Statement.Yearly.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
            r += countValues;
            sh.Rows[$"{startGroup2}:{r}"].Group();
            r++;
            sh.Rows[$"{startGroup1}:{r}"].Group();
            r++;

            // Earnings 
            sh.Cells[r, 1] = "Earnings";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            startGroup1 = r;

            sh.Cells[r, 1] = "History";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            startGroup2 = r;
            EarningsHistoryData earningsHistoryData = new EarningsHistoryData();

            c = 1;
            properties = earningsHistoryData.GetType().GetProperties();
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            countValues = data.Earnings.History.Values.Count;
            val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;

                foreach (EarningsHistoryData item in data.Earnings.History.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
            r += countValues;

            sh.Rows[$"{startGroup2}:{r}"].Group();

            r++;
            sh.Cells[r, 1] = "Trend";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            startGroup2 = r;
            EarningsTrendData earningsTrendData = new EarningsTrendData();

            c = 1;
            properties = earningsTrendData.GetType().GetProperties();
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            countValues = data.Earnings.Trend.Values.Count;
            val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;

                foreach (EarningsTrendData item in data.Earnings.Trend.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
            r += countValues;
            sh.Rows[$"{startGroup2}:{r}"].Group();
            r++;
            sh.Rows[$"{startGroup1}:{r}"].Group();

            sh.Outline.AutomaticStyles = false;
            sh.Outline.SummaryRow = Excel.XlSummaryRow.xlSummaryAbove;

            sh.Outline.ShowLevels(1);
        }

        public static void LoadFundamentalGeneral(FundamentalData data)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;

            int r = Globals.ThisAddIn.Application.ActiveCell.Row;

            // General
            sh.Cells[r, 1] = "General";
            sh.Cells[r, 1].Font.Bold = true;
            r++;


            sh.Cells[r, 1] = "Code";
            sh.Cells[r, 2] = data.General.Code;

            sh.Cells[r, 3] = "Type";
            sh.Cells[r, 4] = data.General.Type;

            r++;

            sh.Cells[r, 1] = "Name";
            sh.Cells[r, 2] = data.General.Name;

            sh.Cells[r, 3] = "Exchange";
            sh.Cells[r, 4] = data.General.Exchange;

            r++;

            sh.Cells[r, 1] = "Currency";
            sh.Cells[r, 2] = data.General.CurrencyCode;
            sh.Cells[r, 3] = data.General.CurrencySymbol;
            r++;

            sh.Cells[r, 1] = "Sector";
            sh.Cells[r, 2] = data.General.Sector;
            r++;

            sh.Cells[r, 1] = "Industry";
            sh.Cells[r, 2] = data.General.Industry;
            r++;

            sh.Cells[r, 1] = "Employees";
            sh.Cells[r, 2] = data.General.FullTimeEmployees;
            r++;

            sh.Cells[r, 1] = "Description";
            sh.Cells[r, 2] = data.General.Description;
        }

        public static void LoadFundamentalHighlights(FundamentalData data)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;

            int r = Globals.ThisAddIn.Application.ActiveCell.Row;



            // Highlights
            sh.Cells[r, 1] = "Highlights";
            sh.Cells[r, 1].Font.Bold = true;
            r++;
     
            sh.Cells[r, 1] = "Market Cap";
            sh.Cells[r, 2] = data.Highlights.MarketCapitalization;

            sh.Cells[r, 3] = "EBITDA";
            sh.Cells[r, 4] = data.Highlights.EBITDA;

            r++;

            sh.Cells[r, 1] = "PE Ratio";
            sh.Cells[r, 2] = data.Highlights.PERatio;

            sh.Cells[r, 3] = "PEG Ratio";
            sh.Cells[r, 4] = data.Highlights.PEGRatio;

            r++;

            sh.Cells[r, 1] = "Earning Share";
            sh.Cells[r, 2] = data.Highlights.EarningsShare;

            r++;

            sh.Cells[r, 1] = "Dividend Share";
            sh.Cells[r, 2] = data.Highlights.DividendShare;

            sh.Cells[r, 3] = "Dividend Yield";
            sh.Cells[r, 4] = data.Highlights.DividendYield;

            r++;

            sh.Cells[r, 1] = "EPS Estimate"; r++;

            sh.Cells[r, 1] = "Current Year";
            sh.Cells[r, 2] = data.Highlights.EPSEstimateCurrentYear;

            r++;

            sh.Cells[r, 1] = "Next Year";
            sh.Cells[r, 2] = data.Highlights.EPSEstimateNextYear;

            r++;

            sh.Cells[r, 1] = "Next Quarter";
            sh.Cells[r, 2] = data.Highlights.EPSEstimateNextQuarter;

        }

        public static void LoadFundamentalEarnings(FundamentalData data)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;

            int r = Globals.ThisAddIn.Application.ActiveCell.Row;
            int c;
            int countValues;
            


            // Earnings 
            sh.Cells[r, 1] = "Earnings";
            sh.Cells[r, 1].Font.Bold = true;


            sh.Cells[r, 1] = "History";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            EarningsHistoryData earningsHistoryData = new EarningsHistoryData();

            c = 1;
            System.Reflection.PropertyInfo[] properties = earningsHistoryData.GetType().GetProperties();
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            countValues = data.Earnings.History.Values.Count;
            object[,] val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;

                foreach (EarningsHistoryData item in data.Earnings.History.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
            r += countValues;

            sh.Cells[r, 1] = "Trend";
            sh.Cells[r, 1].Font.Bold = true;

            r++;

            EarningsTrendData earningsTrendData = new EarningsTrendData();

            c = 1;
            properties = earningsTrendData.GetType().GetProperties();
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            countValues = data.Earnings.Trend.Values.Count;
            val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;

                foreach (EarningsTrendData item in data.Earnings.Trend.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
    
        }

        public static void LoadFundamentalIncomeStatement(FundamentalData data)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;

            int r = Globals.ThisAddIn.Application.ActiveCell.Row;

            // Income_Statement
            sh.Cells[r, 1] = "Income Statement";
            sh.Cells[r, 1].Font.Bold = true;

            r++;

            sh.Cells[r, 1] = "Quarterly";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
            Income_StatementData income_StatementData = new Income_StatementData();

            int c = 1;
            System.Reflection.PropertyInfo[] properties = income_StatementData.GetType().GetProperties();
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            int countValues = data.Financials.Income_Statement.Quarterly.Values.Count;
            object[,] val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;

                foreach (Income_StatementData item in data.Financials.Income_Statement.Quarterly.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
            r += countValues;

            sh.Cells[r, 1] = "Yearly";
            sh.Cells[r, 1].Font.Bold = true;

            r++;

            c = 1;
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            countValues = data.Financials.Income_Statement.Yearly.Values.Count;
            val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;
                foreach (Income_StatementData item in data.Financials.Income_Statement.Yearly.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;

        }

        public static void LoadFundamentalBalanceSheet(FundamentalData data)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;

            int r = Globals.ThisAddIn.Application.ActiveCell.Row;

            // Balance Sheet
            sh.Cells[r, 1] = "Balance Sheet";
            sh.Cells[r, 1].Font.Bold = true;

            r++;


            sh.Cells[r, 1] = "Quarterly";
            sh.Cells[r, 1].Font.Bold = true;

            r++;
 
            Balance_SheetData balance_SheetData = new Balance_SheetData();

            int c = 1;
            System.Reflection.PropertyInfo[] properties = balance_SheetData.GetType().GetProperties();
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            int countValues = data.Financials.Balance_Sheet.Quarterly.Values.Count;
            object[,] val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;

                foreach (Balance_SheetData item in data.Financials.Balance_Sheet.Quarterly.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;
            r += countValues;

            sh.Cells[r, 1] = "Yearly";
            sh.Cells[r, 1].Font.Bold = true;

            r++;

            c = 1;
            foreach (var prop in properties)
            {
                sh.Cells[r, c] = prop.Name;
                c++;
            }

            r++;

            c = 1;
            countValues = data.Financials.Balance_Sheet.Yearly.Values.Count;
            val = new object[countValues, properties.Length];
            foreach (var prop in properties)
            {
                int i = 0;
                foreach (Balance_SheetData item in data.Financials.Balance_Sheet.Yearly.Values)
                {
                    val[i, c - 1] = prop.GetValue(item);
                    i++;
                }

                c++;
            }
            sh.Range[sh.Cells[r, 1], sh.Cells[r + countValues - 1, c - 1]].Value = val;

        }
    }
}
