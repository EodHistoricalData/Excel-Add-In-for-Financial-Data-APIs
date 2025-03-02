﻿using EODAddIn.Utils;
using System;
using System.Collections.Generic;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using static EODAddIn.Utils.ExcelUtils;
using EOD.Model.Fundamental;
using Microsoft.Office.Interop.Excel;
using EOD.Model.BulkFundamental;
using System.Windows.Forms;

namespace EODAddIn.BL.FundamentalDataPrinter
{
    public class FundamentalDataPrinter
    {
        /// <summary>
        /// Print all Fundamenal data to worksheet
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalAll(FundamentalData data, string ticker)
        {
            try
            {
                SetNonInteractive();

                string nameSheet = GetWorksheetNewName($"{ticker}-Fundamental");

                Worksheet sh = AddSheet(nameSheet);

                if (SheetExists(nameSheet))
                {
                    sh = Globals.ThisAddIn.Application.Worksheets[nameSheet];
                    int maxrow = RowsCount(sh);
                    sh.Range[$"A1:Z{maxrow}"].ClearContents();
                }
                else
                {
                    sh = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                    sh.Name = nameSheet;
                }

                int row = 1;

                if (data.Statistics != null)
                {
                    // crypto
                    PrintCrypto(data, sh);
                    return;
                }

                int startGroup1 = 2;

                row = PrintFundamentalGeneral(data, sh.Cells[row, 1]);
                row++;
                row--;
                sh.Rows[$"{startGroup1}:{row}"].Group();
                row++;

                startGroup1 = row + 1;
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
                sh.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;
                //sh.UsedRange.EntireColumn.AutoFit();
                sh.UsedRange.ColumnWidth = 20.0;
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
        /// Displays fundamental data on the worksheet in the active cell
        /// </summary>
        /// <param name="data">fundamental data</param>
        public static void PrintFundamentalGeneral(EOD.Model.Fundamental.FundamentalData data)
        {
            PrintFundamentalGeneral(data, Globals.ThisAddIn.Application.ActiveCell);
        }

        /// <summary>
        ///  Displays Highlights data on the worksheet in the active cell
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalHighlights(EOD.Model.Fundamental.FundamentalData data)
        {
            PrintFundamentalHighlights(data, Globals.ThisAddIn.Application.ActiveCell);
        }

        /// <summary>
        ///Displays Earnings data on the worksheet in the active cell
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalEarnings(EOD.Model.Fundamental.FundamentalData data)
        {
            if (data.Earnings != null)
            {
                PrintFundamentalData("Earnings", data.Earnings.History, data.Earnings.Trend, Globals.ThisAddIn.Application.ActiveCell, "History", "Trend");
            }
            else
            {
                MessageBox.Show("There is no available data for the selected parameters.", "No data", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Displays Cash FLow data on the worksheet in the active cell
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalCashFlow(EOD.Model.Fundamental.FundamentalData data)
        {
            PrintFundamentalData("Cash Flow", data.Financials.Cash_Flow.Quarterly, data.Financials.Cash_Flow.Yearly, Globals.ThisAddIn.Application.ActiveCell);
        }

        /// <summary>
        ///Displays Balance Sheet data on the worksheet in the active cell
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalBalanceSheet(EOD.Model.Fundamental.FundamentalData data)
        {
            if (data.Financials != null)
            {
                PrintFundamentalData("Balance Sheet", data.Financials.Balance_Sheet.Quarterly, data.Financials.Balance_Sheet.Yearly, Globals.ThisAddIn.Application.ActiveCell);
            }
            else
            {
                MessageBox.Show("There is no available data for the selected parameters.", "No data", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Displays Income Statement data on the worksheet in the active cell
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalIncomeStatement(EOD.Model.Fundamental.FundamentalData data)
        {
            if (data.Financials != null)
            {
                PrintFundamentalData("Income Statement", data.Financials.Income_Statement.Quarterly, data.Financials.Income_Statement.Yearly, Globals.ThisAddIn.Application.ActiveCell);
            }
            else
            {
                MessageBox.Show("There is no available data for the selected parameters.", "No data", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Displays Highlights data on a worksheet
        /// </summary>
        /// <param name="data">Fundamental data</param>
        /// <param name="range">The cell where printing starts</param>
        /// <returns>The number of the last involved line</returns>
        private static int PrintFundamentalHighlights(EOD.Model.Fundamental.FundamentalData data, Excel.Range range)
        {
            if (data.Highlights == null)
            {

                MessageBox.Show("There is no available data for the selected parameters.", "No data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            sh.Cells[row, column] = "Highlights";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            sh.Cells[row, column] = "Market Cap";
            sh.Cells[row, column + 1] = data.Highlights.MarketCapitalization;

            sh.Cells[row, column + 2] = "EBITDA";
            sh.Cells[row, column + 3] = data.Highlights.EBITDA;
            row++;

            sh.Cells[row, column] = "PE Ratio";
            sh.Cells[row, column + 1] = data.Highlights.PERatio;

            sh.Cells[row, column + 2] = "PEG Ratio";
            sh.Cells[row, column + 3] = data.Highlights.PEGRatio;
            row++;

            sh.Cells[row, column] = "Earning Share";
            sh.Cells[row, column + 1] = data.Highlights.EarningsShare;
            row++;

            sh.Cells[row, column] = "Dividend Share";
            sh.Cells[row, column + 1] = data.Highlights.DividendShare;

            sh.Cells[row, column + 2] = "Dividend Yield";
            sh.Cells[row, column + 3] = data.Highlights.DividendYield;
            row++;

            sh.Cells[row, column] = "EPS Estimate";
            row++;

            sh.Cells[row, column] = "Current Year";
            sh.Cells[row, column + 1] = data.Highlights.EPSEstimateCurrentYear;

            row++;

            sh.Cells[row, column] = "Next Year";
            sh.Cells[row, column + 1] = data.Highlights.EPSEstimateNextYear;

            row++;

            sh.Cells[row, column] = "Next Quarter";
            sh.Cells[row, column + 1] = data.Highlights.EPSEstimateNextQuarter;

            return row;
        }

        /// <summary>
        ///  Displays General data on a worksheet
        /// </summary>
        /// <param name="data">Fundamental data</param>
        /// <param name="range">The cell where printing starts</param>
        /// <returns>The number of the last involved line</returns>
        private static int PrintFundamentalGeneral(FundamentalData data, Range range)
        {
            Worksheet sh = range.Parent;
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
                                                    Dictionary<string, T> dataTable1,
                                                    Dictionary<string, U> dataTable2,
                                                    Range range,
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
            row += dataTable1.Values.Count + 1;
            sh.Rows[$"{startGroup2}:{row}"].Group();
            row++;
            startGroup2 = row + 1;

            PrintTablePeriod(dataTable2Name, sh.Cells[row, column], dataTable2);
            row += dataTable2.Values.Count + 1;
            sh.Rows[$"{startGroup2}:{row}"].Group();
            sh.Rows[$"{startGroup1}:{row}"].Group();

            return row;
        }
        /// <summary>
        ///Print a table with data by period
        /// </summary>
        /// <typeparam name="T">Data type in the table</typeparam>
        /// <param name="periodName">Period Name</param>
        /// <param name="range">target cell</param>
        /// <param name="data">Table data</param>
        /// <param name="properties">Property List</param>
        private static void PrintTablePeriod<T>(string periodName, Range range, Dictionary<string, T> data)
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
            sh.Range[sh.Cells[row, range.Column], sh.Cells[row + countValues - 1, column - 1]].Value = val;
        }

        private static void PrintCrypto(FundamentalData data, Worksheet worksheet)
        {
            int row = 1;
            Type myType = data.GetType();
            IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

            foreach (PropertyInfo prop in props)
            {
                object propValue = prop.GetValue(data, null);
                switch (prop.Name)
                {
                    case "General":
                        row = Printer.PrintVerticalTable(prop.Name, propValue, worksheet.Cells[row, 1]);
                        row++;
                        break;
                    case "Statistics":
                        row = Printer.PrintVerticalTable(prop.Name, propValue, worksheet.Cells[row, 1]);
                        row++;
                        break;
                }
            }
        }


    }
}
