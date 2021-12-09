using EODAddIn.Model;
using EODAddIn.Utils;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;

using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.BL
{
    public class LoadToExcel
    {
        /// <summary>
        /// Загрузка данных на конец дня на лист Excel
        /// </summary>
        /// <param name="endOfDays">Список с данными</param>
        /// <param name="ticker">Тикер</param>
        /// <param name="period">Период</param>
        /// <param name="chart">Необходимость построения диаграммы</param>
        public static void PrintEndOfDay(List<EndOfDay> endOfDays, string ticker, string period, bool chart)
        {
            bool createSheet = true;
            string nameSheet = $"{ticker}-{period}";
            Excel.Worksheet worksheet;

            if (ExcelUtils.SheetExists(nameSheet))
            {
                worksheet = Globals.ThisAddIn.Application.Worksheets[nameSheet];
                int maxrow = ExcelUtils.RowsCount(worksheet);
                worksheet.Range[$"A1:J{maxrow}"].ClearContents();
                createSheet = false;
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

            try
            {
                ExcelUtils.OnStart();
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
            catch (Exception)
            {
                throw;
            }
            finally
            {
                ExcelUtils.OnEnd();
            }

            if (!createSheet) return;
            if (!chart) return;

            worksheet.Range["A2:E3"].Select();

            Excel.Shape shp = worksheet.Shapes.AddChart2(-1, Excel.XlChartType.xlStockOHLC);
            Excel.Chart chrt = shp.Chart;

            chrt.ChartGroups(1).UpBars.Format.Fill.ForeColor.RGB = Color.FromArgb(0, 176, 80);
            chrt.ChartGroups(1).DownBars.Format.Fill.ForeColor.RGB = Color.FromArgb(255, 0, 0);

            worksheet.Cells[2, 13].Value = DateTime.Today.AddDays(-90);
            worksheet.Cells[3, 13].Value = DateTime.Today.AddDays(-1);

            worksheet.Cells[2, 12].Value = "Start";
            worksheet.Cells[3, 12].Value = "End";

            worksheet.Range["A:A"].EntireColumn.AutoFit();
            worksheet.Range["M:M"].EntireColumn.AutoFit();
            worksheet.Range["L:L"].EntireColumn.AutoFit();

            Excel.Range rng = worksheet.Range["Q1"];
            string formula;

            rng.FormulaR1C1 = $"=IFERROR(OFFSET('{worksheet.Name}'!R2C1,MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)-2,1,MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C1:C1,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)+1,1),1)";
            formula = rng.FormulaR1C1Local;
            worksheet.Names.Add("_open", RefersToR1C1Local: formula);

            rng.FormulaR1C1 = $"=IFERROR(OFFSET('{worksheet.Name}'!R2C1,MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)-2,2,MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C1:C1,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)+1,1),1)";
            formula = rng.FormulaR1C1Local;
            worksheet.Names.Add("_high", RefersToR1C1Local: formula);

            rng.FormulaR1C1 = $"=IFERROR(OFFSET('{worksheet.Name}'!R2C1,MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)-2,3,MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C1:C1,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)+1,1),1)";
            formula = rng.FormulaR1C1Local;
            worksheet.Names.Add("_low", RefersToR1C1Local: formula);

            rng.FormulaR1C1 = $"=IFERROR(OFFSET('{worksheet.Name}'!R2C1,MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)-2,4,MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C1:C1,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)+1,1),1)";
            formula = rng.FormulaR1C1Local;
            worksheet.Names.Add("_close", RefersToR1C1Local: formula);

            rng.FormulaR1C1 = $"=OFFSET('{worksheet.Name}'!R2C1,IFERROR(MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)-2,0),0,IFERROR(MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C1:C1,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C1:C1,1)+1,1),1)";
            formula = rng.FormulaR1C1Local;
            worksheet.Names.Add("_period", RefersToR1C1Local: formula);

            rng.ClearContents();

            chrt.FullSeriesCollection(1).Values = $"='{worksheet.Name}'!_open";
            chrt.FullSeriesCollection(2).Values = $"='{worksheet.Name}'!_high";
            chrt.FullSeriesCollection(3).Values = $"='{worksheet.Name}'!_low";
            chrt.FullSeriesCollection(4).Values = $"='{worksheet.Name}'!_close";
            chrt.FullSeriesCollection(1).XValues = $"='{worksheet.Name}'!_period";

            chrt.FullSeriesCollection(4).Trendlines().Add();
            chrt.FullSeriesCollection(4).Trendlines(1).Type = Excel.XlTrendlineType.xlMovingAvg;
            chrt.FullSeriesCollection(4).Trendlines(1).Period = 2;

            shp.Left = (float)worksheet.Cells[5, 12].Left;
            shp.Top = (float)worksheet.Cells[5, 12].Top;
            shp.Height = 340.157480315f;
            shp.Width = 680.3149606299f;

            chrt.ChartTitle.Caption = worksheet.Name;
        }

        /// <summary>
        /// Загрузка всех фундаментальных данных на лист Excel
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalAll(FundamentalData data)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;

            int row = 1;
            int startGroup1 = 2;

            row = PrintFundamentalGeneral(data, sh.Cells[row, 1]);
            row++;

            sh.Rows[$"{startGroup1}:{row}"].Group();
            row++;

            startGroup1 = row+1;
            row = PrintFundamentalHighlights(data, sh.Cells[row, 1]);
            
            sh.Rows[$"{startGroup1}:{row}"].Group();
            row++;

            row = PrintFundamentalData("Balance Sheet", data.Financials.Balance_Sheet.Quarterly, data.Financials.Balance_Sheet.Yearly, sh.Cells[row, 1]);
            row++;

            row = PrintFundamentalData("Income Statement", data.Financials.Income_Statement.Quarterly, data.Financials.Income_Statement.Yearly, sh.Cells[row, 1]);
            row++;

            row = PrintFundamentalData("Cash Flow", data.Financials.Cash_Flow.Quarterly, data.Financials.Cash_Flow.Yearly, sh.Cells[row, 1]);
            row++;

            PrintFundamentalData("Earnings", data.Earnings.History, data.Earnings.Trend, sh.Cells[row, 1], "History", "Trend");


            sh.Outline.AutomaticStyles = false;
            sh.Outline.SummaryRow = Excel.XlSummaryRow.xlSummaryAbove;

            sh.Outline.ShowLevels(1);
        }

        /// <summary>
        /// Выводит фундаментальные General данные на лист в активную ячейку
        /// </summary>
        /// <param name="data">Фундаментальные данные</param>
        public static void PrintFundamentalGeneral(FundamentalData data)
        {
            PrintFundamentalGeneral(data, Globals.ThisAddIn.Application.ActiveCell);
        }

        /// <summary>
        /// Выводит фундаментальные Highlights данные на лист в активную ячейку
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalHighlights(FundamentalData data)
        {
            PrintFundamentalHighlights(data, Globals.ThisAddIn.Application.ActiveCell);
        }

        /// <summary>
        /// Выводит фундаментальные Earnings данные на лист в активную ячейку
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalEarnings(FundamentalData data)
        {
            PrintFundamentalData("Earnings", data.Earnings.History, data.Earnings.Trend, Globals.ThisAddIn.Application.ActiveCell, "History", "Trend");
        }

        /// <summary>
        /// Выводит фундаментальные Cash Flow данные на лист в активную ячейку
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalCashFlow(FundamentalData data)
        {
            PrintFundamentalData("Cash Flow", data.Financials.Cash_Flow.Quarterly, data.Financials.Cash_Flow.Yearly, Globals.ThisAddIn.Application.ActiveCell);
        }

        /// <summary>
        /// Выводит фундаментальные Balance Sheet данные на лист в активную ячейку
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalBalanceSheet(FundamentalData data)
        {
            PrintFundamentalData("Balance Sheet", data.Financials.Balance_Sheet.Quarterly, data.Financials.Balance_Sheet.Yearly, Globals.ThisAddIn.Application.ActiveCell);
        }

        /// <summary>
        /// Выводит фундаментальные Income Statement данные на лист в активную ячейку
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalIncomeStatement(FundamentalData data)
        {
            PrintFundamentalData("Income Statement", data.Financials.Income_Statement.Quarterly, data.Financials.Income_Statement.Yearly, Globals.ThisAddIn.Application.ActiveCell);
        }



        /// <summary>
        /// Выводит фундаментальные Highlights данные на лист 
        /// </summary>
        /// <param name="data">Фундаментальные данные</param>
        /// <param name="range">Ячейка с которой начинается печать</param>
        /// <returns>Номер последней задействованной строки</returns>
        private static int PrintFundamentalHighlights(FundamentalData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Highlights";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Market Cap";
            sh.Cells[row, column+1] = data.Highlights.MarketCapitalization;

            sh.Cells[row, column+2] = "EBITDA";
            sh.Cells[row, column+3] = data.Highlights.EBITDA;
            row++;

            sh.Cells[row, column] = "PE Ratio";
            sh.Cells[row, column+1] = data.Highlights.PERatio;

            sh.Cells[row, column+2] = "PEG Ratio";
            sh.Cells[row, column+3] = data.Highlights.PEGRatio;
            row++;

            sh.Cells[row, column] = "Earning Share";
            sh.Cells[row, column+1] = data.Highlights.EarningsShare;
            row++;

            sh.Cells[row, column] = "Dividend Share";
            sh.Cells[row, column+1] = data.Highlights.DividendShare;

            sh.Cells[row, column+2] = "Dividend Yield";
            sh.Cells[row, column+3] = data.Highlights.DividendYield;
            row++;

            sh.Cells[row, column] = "EPS Estimate"; 
            row++;

            sh.Cells[row, column] = "Current Year";
            sh.Cells[row, column+1] = data.Highlights.EPSEstimateCurrentYear;

            row++;

            sh.Cells[row, column] = "Next Year";
            sh.Cells[row, column+1] = data.Highlights.EPSEstimateNextYear;

            row++;

            sh.Cells[row, column] = "Next Quarter";
            sh.Cells[row, column+1] = data.Highlights.EPSEstimateNextQuarter;

            return row;
        }
        
        /// <summary>
        /// Выводит фундаментальные General данные на лист
        /// </summary>
        /// <param name="data">Фундаментальные данные</param>
        /// <param name="range">Ячейка с которой начинается печать</param>
        /// <returns>Номер последней задействованной строки</returns>
        private static int PrintFundamentalGeneral(FundamentalData data, Excel.Range range)
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

            return row;
        }

        private static int PrintFundamentalData<T, U>(string nameData, 
                                                    Dictionary<DateTime, T> dataTable1, 
                                                    Dictionary<DateTime, U> dataTable2, 
                                                    Excel.Range range, 
                                                    string dataTable1Name = "Quarterly", 
                                                    string dataTable2Name = "Yearly")
             where T : class
             where U : class
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = $"{nameData}";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            int startGroup1 = row;
            int startGroup2 = row + 1;

            PrintTablePeriod(dataTable1Name, sh.Cells[row, column], dataTable1);
            row += dataTable1.Values.Count+1;
            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;
            startGroup2 = row + 1;

            PrintTablePeriod(dataTable2Name, sh.Cells[row, column], dataTable2);
            row += dataTable2.Values.Count+1;
            sh.Rows[$"{startGroup2}:{row}"].Group();
            sh.Rows[$"{startGroup1}:{row}"].Group();

            return row;
        }


        /// <summary>
        /// Печать таблицы с данными по периоду
        /// </summary>
        /// <typeparam name="T">Тип данных в таблице</typeparam>
        /// <param name="periodName">Название периода</param>
        /// <param name="range">Целевая ячейка</param>
        /// <param name="data">Данные таблицы</param>
        /// <param name="properties">Список свойств</param>
        private static void PrintTablePeriod<T>(string periodName, Excel.Range range, Dictionary<DateTime, T> data)
            where T : class, new()
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = $"{periodName}";
            sh.Cells[row, column].Font.Bold = true;
            row++;

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
            int j = 0;
            foreach (var prop in properties)
            {
                int i = 0;
                foreach (T item in data.Values)
                {
                    val[i, j] = prop.GetValue(item);
                    i++;
                }
                column++;
                j++;
            }
            sh.Range[sh.Cells[row, range.Column], sh.Cells[row + countValues-1, column - 1]].Value = val;

        }
    }
}
