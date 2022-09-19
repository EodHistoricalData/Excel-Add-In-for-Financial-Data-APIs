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

namespace EODAddIn.BL
{
    public class LoadToExcel
    {
        private static Excel.Application _xlsApp = Globals.ThisAddIn.Application;
        private static bool CreateSheet = true;

        private static Excel.Worksheet AddSheet(string nameSheet)
        {
            Excel.Worksheet worksheet = null;
            try
            {
                if (ExcelUtils.SheetExists(nameSheet))
                {
                    worksheet = Globals.ThisAddIn.Application.Worksheets[nameSheet];
                    int maxrow = ExcelUtils.RowsCount(worksheet);
                    worksheet.Range[$"A1:AZ{maxrow}"].ClearContents();
                    CreateSheet = false;
                }
                else
                {
                    worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                    worksheet.Name = nameSheet;
                }
            }
            catch (Exception ex)
            {
                ErrorReport errorReport = new ErrorReport(ex);
                errorReport.MessageToUser("Can't create a worksheet.");
            }
            return worksheet;
        }

        /// <summary>
        /// Выключение интерактивности
        /// </summary>
        private static void SetNonInteractive()
        {
            while (_xlsApp.Interactive)
            {
                try
                {
                    _xlsApp.Interactive = false;
                }
                catch { }
            }
        }

        /// <summary>
        /// Загрузка данных на конец дня на лист Excel
        /// </summary>
        /// <param name="endOfDays">Список с данными</param>
        /// <param name="ticker">Тикер</param>
        /// <param name="period">Период</param>
        /// <param name="chart">Необходимость построения диаграммы</param>
        public static void PrintEndOfDay(List<EndOfDay> endOfDays, string ticker, string period, bool chart)
        {
            try
            {
                SetNonInteractive();

                string nameSheet = $"{ticker}-{period}";

                Excel.Worksheet worksheet = AddSheet(nameSheet);

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

                if (!CreateSheet) return;
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
        /// Загрузка данных в течении дня на лист Excel
        /// </summary>
        /// <param name="endOfDays">Список с данными</param>
        /// <param name="ticker">Тикер</param>
        /// <param name="period">Период</param>
        /// <param name="chart">Необходимость построения диаграммы</param>
        public static void PrintIntraday(List<Intraday> intraday, string ticker, string interval, bool chart)
        {
            try
            {
                SetNonInteractive();

                string nameSheet = $"{ticker}-{interval}";

                Excel.Worksheet worksheet = AddSheet(nameSheet);

                int r = 2;
                worksheet.Cells[r, 1] = "DateTime";
                worksheet.Cells[r, 2] = "Gmtoffset";
                worksheet.Cells[r, 3] = "DateTime";
                worksheet.Cells[r, 4] = "Open";
                worksheet.Cells[r, 5] = "High";
                worksheet.Cells[r, 6] = "Low";
                worksheet.Cells[r, 7] = "Close";
                worksheet.Cells[r, 8] = "Volume";
                worksheet.Cells[r, 9] = "Timestamp";
                try
                {
                    ExcelUtils.OnStart();
                    foreach (Intraday item in intraday)
                    {
                        r++;
                        worksheet.Cells[r, 1] = item.DateTime;
                        worksheet.Cells[r, 2] = item.Gmtoffset;
                        worksheet.Cells[r, 3] = item.DateTime;
                        worksheet.Cells[r, 4] = item.Open;
                        worksheet.Cells[r, 5] = item.High;
                        worksheet.Cells[r, 6] = item.Low;
                        worksheet.Cells[r, 7] = item.Close;
                        worksheet.Cells[r, 8] = item.Volume;
                        worksheet.Cells[r, 9] = item.Timestamp;
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

                if (!CreateSheet) return;
                if (!chart) return;

                worksheet.Range["C2:G3"].Select();

                Excel.Shape shp = worksheet.Shapes.AddChart2(-1, Excel.XlChartType.xlStockOHLC);
                Excel.Chart chrt = shp.Chart;

                chrt.ChartGroups(1).UpBars.Format.Fill.ForeColor.RGB = Color.FromArgb(0, 176, 80);
                chrt.ChartGroups(1).DownBars.Format.Fill.ForeColor.RGB = Color.FromArgb(255, 0, 0);

                worksheet.Cells[2, 13].Value = DateTime.Now.AddDays(-10);
                worksheet.Cells[3, 13].Value = DateTime.Now.AddDays(-1);

                worksheet.Range["I:I"].EntireColumn.Hidden = true;
                Excel.Range rng = worksheet.Range["Q1"];

                string formula;
                rng.FormulaR1C1 = $"=IFERROR(OFFSET('{worksheet.Name}'!R2C3,MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)-2,1,MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C3:C3,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)+1,1),1)";
                formula = rng.FormulaR1C1Local;
                worksheet.Names.Add("_open", RefersToR1C1Local: formula);

                rng.FormulaR1C1 = $"=IFERROR(OFFSET('{worksheet.Name}'!R2C3,MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)-2,2,MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C3:C3,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)+1,1),1)";
                formula = rng.FormulaR1C1Local;
                worksheet.Names.Add("_high", RefersToR1C1Local: formula);

                rng.FormulaR1C1 = $"=IFERROR(OFFSET('{worksheet.Name}'!R2C3,MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)-2,3,MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C3:C3,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)+1,1),1)";
                formula = rng.FormulaR1C1Local;
                worksheet.Names.Add("_low", RefersToR1C1Local: formula);

                rng.FormulaR1C1 = $"=IFERROR(OFFSET('{worksheet.Name}'!R2C3,MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)-2,4,MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C3:C3,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)+1,1),1)";
                formula = rng.FormulaR1C1Local;
                worksheet.Names.Add("_close", RefersToR1C1Local: formula);

                rng.FormulaR1C1 = $"=OFFSET('{worksheet.Name}'!R2C3,IFERROR(MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)-2,0),0,IFERROR(MATCH('{worksheet.Name}'!R3C13,'{worksheet.Name}'!C3:C3,1)-MATCH('{worksheet.Name}'!R2C13,'{worksheet.Name}'!C3:C3,1)+1,1),1)";
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

                shp.Left = (float)worksheet.Cells[5, 11].Left;
                shp.Top = (float)worksheet.Cells[5, 11].Top;
                shp.Height = 340.157480315f;
                shp.Width = 680.3149606299f;
                chrt.ChartTitle.Caption = worksheet.Name;


                int lastrow = ExcelUtils.RowsCount(worksheet);
                if (lastrow <= 2) return;

                Excel.Range timestampRng = worksheet.Range[$"I2:I{lastrow}"];
                Excel.Range timeRng = worksheet.Range[$"C2:C{lastrow}"];
                timeRng.Value = timestampRng.Value;

                worksheet.Cells[2, 11].Value = "Start";
                worksheet.Cells[3, 11].Value = "End";
                worksheet.Cells[1, 12].Value = "Date";
                worksheet.Cells[1, 13].Value = "Timestamp";

                Excel.Range firstTimeRng = worksheet.Range[$"A2:A{lastrow}"];
                worksheet.Cells[2, 12].Value = firstTimeRng.Cells[2, 1].Value;
                worksheet.Cells[3, 12].Value = firstTimeRng.Cells[firstTimeRng.Rows.Count, 1].Value;

                Excel.Range vlookupRng = worksheet.Range[$"A2:C{lastrow}"];
                string addressRng = vlookupRng.Address;

                string addressCell = worksheet.Cells[2, 12].Address[RowAbsolute: true, ColumnAbsolute: true];
                Excel.Range cellEdit = worksheet.Cells[2, 13];
                cellEdit.Formula = $"=IFERROR(VLOOKUP({addressCell},{addressRng},3,),1)";
                cellEdit.NumberFormat = "@";

                addressCell = worksheet.Cells[3, 12].Address[RowAbsolute: true, ColumnAbsolute: true];
                cellEdit = worksheet.Cells[3, 13];
                cellEdit.Formula = $"=IFERROR(VLOOKUP({addressCell},{addressRng},3,),1)";
                cellEdit.NumberFormat = "@";
                timeRng.NumberFormat = "@";

                worksheet.Range["A:C"].EntireColumn.AutoFit();
                worksheet.Range["L:M"].EntireColumn.AutoFit();

                shp.Left = (float)worksheet.Cells[5, 11].Left;
                shp.Top = (float)worksheet.Cells[5, 11].Top;
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
        /// Загрузка данных Опцонов на лист Excel
        /// </summary>
        /// <param name="data">Данные</param>
        public static void PrintOptions(EOD.Model.OptionsData.OptionsData data, string ticker)
        {
            try
            {
                SetNonInteractive();

                string nameSheet = $"{ticker}-Options";

                Excel.Worksheet sh = AddSheet(nameSheet);

                int row = 1;
                int column = 1;

                sh.Cells[row, column] = "Code";
                sh.Cells[row, column + 1] = data.Code;
                sh.Cells[row, column + 2] = "Exchange";
                sh.Cells[row, column + 3] = data.Exchange;
                row++;

                sh.Cells[row, column] = "Last trade date";
                sh.Cells[row, column + 1] = data.LastTradeDate;
                sh.Cells[row, column + 2] = "Last trade price";
                sh.Cells[row, column + 3] = data.LastTradePrice;
                row++;
                sh.Cells[row, column] = "Data";
                sh.Cells[row, column].Font.Bold = true;
                row++;

                try
                {
                    Globals.ThisAddIn.Application.ScreenUpdating = false;
                    Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
                    PrintOptionsData(data, sh.Cells[row, 1]);

                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                    Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                }

                sh.Outline.AutomaticStyles = false;
                sh.Outline.SummaryRow = Excel.XlSummaryRow.xlSummaryAbove;

                sh.Outline.ShowLevels(2);
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

        private static void PrintOptionsData(EOD.Model.OptionsData.OptionsData data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            int startGroup2 = row;
            sh.Cells[row, column] = "Options";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            for (int i = 0; i < data.Data.Count; i++)
            {
                sh.Cells[row, column] = "Expiration Date";
                sh.Cells[row + 1, column] = data.Data[i].ExpirationDate;
                column++;
                sh.Cells[row, column] = "Implied Volatility";
                sh.Cells[row + 1, column] = data.Data[i].ImpliedVolatility;
                column++;
                sh.Cells[row, column] = "Put Volume";
                sh.Cells[row + 1, column] = data.Data[i].PutVolume;
                column++;
                sh.Cells[row, column] = "Call Volume";
                sh.Cells[row + 1, column] = data.Data[i].CallVolume;
                column++;
                sh.Cells[row, column] = "Put Call Volume Ratio";
                sh.Cells[row + 1, column] = data.Data[i].PutCallVolumeRatio;
                column++;
                sh.Cells[row, column] = "Put Open Interest";
                sh.Cells[row + 1, column] = data.Data[i].PutOpenInterest;
                column++;
                sh.Cells[row, column] = "Call Open Interest";
                sh.Cells[row + 1, column] = data.Data[i].CallOpenInterest;
                column++;
                sh.Cells[row, column] = "Put Call Open Interest Ratio";
                sh.Cells[row + 1, column] = data.Data[i].PutCallOpenInterestRatio;
                column++;
                sh.Cells[row, column] = "Options Count";
                sh.Cells[row + 1, column] = data.Data[i].OptionsCount;
                row += 2;
                column = 1;
                row = PrintOptionsTable(data.Data[i].Options, sh.Cells[row, 1]);
                row++;
            }
            sh.Rows[$"{startGroup2}:{row}"].Group();
            return;
        }

        private static int PrintOptionsTable(Dictionary<string, List<EOD.Model.OptionsData.ContractData>> data, Excel.Range range)
        {
            Excel.Worksheet sh = range.Parent;
            int row = range.Row;
            int column = range.Column;

            int startGroup3 = row;
            sh.Cells[row, column] = "CALL";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            #region +++Title+++
            sh.Cells[row, column] = "Contract Name";
            column++;
            sh.Cells[row, column] = "Contract Size";
            column++;
            sh.Cells[row, column] = "Contract Period";
            column++;
            sh.Cells[row, column] = "Currency";
            column++;
            sh.Cells[row, column] = "Type";
            column++;
            sh.Cells[row, column] = "In The Money";
            column++;
            sh.Cells[row, column] = "Last Trade Date Time";
            column++;
            sh.Cells[row, column] = "Expiration Date";
            column++;
            sh.Cells[row, column] = "Strike";
            column++;
            sh.Cells[row, column] = "Last Price";
            column++;
            sh.Cells[row, column] = "Bid";
            column++;
            sh.Cells[row, column] = "Ask";
            column++;
            sh.Cells[row, column] = "Change";
            column++;
            sh.Cells[row, column] = "Change Percent";
            column++;
            sh.Cells[row, column] = "Volume";
            column++;
            sh.Cells[row, column] = "Open Interest";
            column++;
            sh.Cells[row, column] = "Implied Volatility";
            column++;
            sh.Cells[row, column] = "Delta";
            column++;
            sh.Cells[row, column] = "Gamma";
            column++;
            sh.Cells[row, column] = "Theta";
            column++;
            sh.Cells[row, column] = "Vega";
            column++;
            sh.Cells[row, column] = "Rho";
            column++;
            sh.Cells[row, column] = "Theoretical";
            column++;
            sh.Cells[row, column] = "Intrinsic Value";
            column++;
            sh.Cells[row, column] = "Time Value";
            column++;
            sh.Cells[row, column] = "Updated At";
            column++;
            sh.Cells[row, column] = "Days Before Expiration";
            #endregion

            row++;
            List<EOD.Model.OptionsData.ContractData> callContracts = data["CALL"];

            for (int j = 0; j < callContracts.Count; j++)
            {
                column = 1;
                sh.Cells[row, column] = callContracts[j].ContractName;
                column++;
                sh.Cells[row, column] = callContracts[j].ContractSize;
                column++;
                sh.Cells[row, column] = callContracts[j].ContractPeriod;
                column++;
                sh.Cells[row, column] = callContracts[j].Currency;
                column++;
                sh.Cells[row, column] = callContracts[j].Type;
                column++;
                sh.Cells[row, column] = callContracts[j].InTheMoney;
                column++;
                sh.Cells[row, column] = callContracts[j].LastTradeDateTime;
                column++;
                sh.Cells[row, column] = callContracts[j].ExpirationDate;
                column++;
                sh.Cells[row, column] = callContracts[j].Strike;
                column++;
                sh.Cells[row, column] = callContracts[j].LastPrice;
                column++;
                sh.Cells[row, column] = callContracts[j].Bid;
                column++;
                sh.Cells[row, column] = callContracts[j].Ask;
                column++;
                sh.Cells[row, column] = callContracts[j].Change;
                column++;
                sh.Cells[row, column] = callContracts[j].ChangePercent;
                column++;
                sh.Cells[row, column] = callContracts[j].Volume;
                column++;
                sh.Cells[row, column] = callContracts[j].OpenInterest;
                column++;
                sh.Cells[row, column] = callContracts[j].ImpliedVolatility;
                column++;
                sh.Cells[row, column] = callContracts[j].Delta;
                column++;
                sh.Cells[row, column] = callContracts[j].Gamma;
                column++;
                sh.Cells[row, column] = callContracts[j].Theta;
                column++;
                sh.Cells[row, column] = callContracts[j].Vega;
                column++;
                sh.Cells[row, column] = callContracts[j].Rho;
                column++;
                sh.Cells[row, column] = callContracts[j].Theoretical;
                column++;
                sh.Cells[row, column] = callContracts[j].IntrinsicValue;
                column++;
                sh.Cells[row, column] = callContracts[j].TimeValue;
                column++;
                sh.Cells[row, column] = callContracts[j].UpdatedAt;
                column++;
                sh.Cells[row, column] = callContracts[j].DaysBeforeExpiration;
                row++;
            }
            sh.Rows[$"{startGroup3}:{row}"].Group();

            column = 1;
            int startGroup4 = row + 1;
            sh.Cells[row, column] = "PUT";
            sh.Cells[row, column].Font.Bold = true;
            row++;

            #region +++Title+++
            column = 1;
            sh.Cells[row, column] = "Contract Name";
            column++;
            sh.Cells[row, column] = "Contract Size";
            column++;
            sh.Cells[row, column] = "Contract Period";
            column++;
            sh.Cells[row, column] = "Currency";
            column++;
            sh.Cells[row, column] = "Type";
            column++;
            sh.Cells[row, column] = "In The Money";
            column++;
            sh.Cells[row, column] = "Last Trade Date Time";
            column++;
            sh.Cells[row, column] = "Expiration Date";
            column++;
            sh.Cells[row, column] = "Strike";
            column++;
            sh.Cells[row, column] = "Last Price";
            column++;
            sh.Cells[row, column] = "Bid";
            column++;
            sh.Cells[row, column] = "Ask";
            column++;
            sh.Cells[row, column] = "Change";
            column++;
            sh.Cells[row, column] = "Change Percent";
            column++;
            sh.Cells[row, column] = "Volume";
            column++;
            sh.Cells[row, column] = "Open Interest";
            column++;
            sh.Cells[row, column] = "Implied Volatility";
            column++;
            sh.Cells[row, column] = "Delta";
            column++;
            sh.Cells[row, column] = "Gamma";
            column++;
            sh.Cells[row, column] = "Theta";
            column++;
            sh.Cells[row, column] = "Vega";
            column++;
            sh.Cells[row, column] = "Rho";
            column++;
            sh.Cells[row, column] = "Theoretical";
            column++;
            sh.Cells[row, column] = "Intrinsic Value";
            column++;
            sh.Cells[row, column] = "Time Value";
            column++;
            sh.Cells[row, column] = "Updated At";
            column++;
            sh.Cells[row, column] = "Days Before Expiration";
            #endregion

            row++;
            List<EOD.Model.OptionsData.ContractData> putContracts = data["PUT"];

            for (int j = 0; j < putContracts.Count; j++)
            {
                column = 1;
                sh.Cells[row, column] = putContracts[j].ContractName;
                column++;
                sh.Cells[row, column] = putContracts[j].ContractSize;
                column++;
                sh.Cells[row, column] = putContracts[j].ContractPeriod;
                column++;
                sh.Cells[row, column] = putContracts[j].Currency;
                column++;
                sh.Cells[row, column] = putContracts[j].Type;
                column++;
                sh.Cells[row, column] = putContracts[j].InTheMoney;
                column++;
                sh.Cells[row, column] = putContracts[j].LastTradeDateTime;
                column++;
                sh.Cells[row, column] = putContracts[j].ExpirationDate;
                column++;
                sh.Cells[row, column] = putContracts[j].Strike;
                column++;
                sh.Cells[row, column] = putContracts[j].LastPrice;
                column++;
                sh.Cells[row, column] = putContracts[j].Bid;
                column++;
                sh.Cells[row, column] = putContracts[j].Ask;
                column++;
                sh.Cells[row, column] = putContracts[j].Change;
                column++;
                sh.Cells[row, column] = putContracts[j].ChangePercent;
                column++;
                sh.Cells[row, column] = putContracts[j].Volume;
                column++;
                sh.Cells[row, column] = putContracts[j].OpenInterest;
                column++;
                sh.Cells[row, column] = putContracts[j].ImpliedVolatility;
                column++;
                sh.Cells[row, column] = putContracts[j].Delta;
                column++;
                sh.Cells[row, column] = putContracts[j].Gamma;
                column++;
                sh.Cells[row, column] = putContracts[j].Theta;
                column++;
                sh.Cells[row, column] = putContracts[j].Vega;
                column++;
                sh.Cells[row, column] = putContracts[j].Rho;
                column++;
                sh.Cells[row, column] = putContracts[j].Theoretical;
                column++;
                sh.Cells[row, column] = putContracts[j].IntrinsicValue;
                column++;
                sh.Cells[row, column] = putContracts[j].TimeValue;
                column++;
                sh.Cells[row, column] = putContracts[j].UpdatedAt;
                column++;
                sh.Cells[row, column] = putContracts[j].DaysBeforeExpiration;
                row++;
            }

            sh.Rows[$"{startGroup4}:{row + 1}"].Group();

            return row;
        }

        /// <summary>
        /// Загрузка всех данных ETF
        /// </summary>
        /// <param name="data"></param>
        public static void PrintEtf(FundamentalData data, string ticker)
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
        /// Заполнение листа данными General для ETF
        /// </summary>
        /// <param name="data"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int PrintEtfGeneral(FundamentalData data, Excel.Range range)
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
            sh.Cells[row, column + 1] = data.General.Category;

            return row;
        }
        /// <summary>
        /// Заполнение листа данными Technicals для ETF
        /// </summary>
        /// <param name="data"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int PrintEtfTechnicals(FundamentalData data, Excel.Range range)
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

            sh.Cells[row, column] = "52WeekHigh";
            sh.Cells[row, column + 1] = data.Technicals.WeekHigh52;
            row++;

            sh.Cells[row, column] = "52WeekLow";
            sh.Cells[row, column + 1] = data.Technicals.WeekLow52;
            row++;

            sh.Cells[row, column] = "50DayMA";
            sh.Cells[row, column + 1] = data.Technicals.DayMA50;
            row++;

            sh.Cells[row, column] = "200DayMA";
            sh.Cells[row, column + 1] = data.Technicals.DayMA200;

            return row;
        }
        /// <summary>
        /// Заполнение листа данными ETF
        /// </summary>
        /// <param name="data"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int PrintEtfData(FundamentalData data, Excel.Range range)
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

        public static int PrintEtfMarketCap(FundamentalData data, Excel.Range range)
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

        public static int PrintEtfAssetAllocation(FundamentalData data, Excel.Range range)
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
            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.StockNonUs.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.StockNonUs.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.StockNonUs.NetAssetsPercent;
            row++;

            sh.Cells[row, column] = "Other";
            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.Other.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.Other.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.Other.NetAssetsPercent;
            row++;

            sh.Cells[row, column] = "Stock US";
            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.StockUs.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.StockUs.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.StockUs.NetAssetsPercent;
            row++;

            sh.Cells[row, column] = "Bond";
            sh.Cells[row, column + 1] = data.ETF_Data.Asset_Allocation.Bond.LongPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Asset_Allocation.Bond.ShortPercent;
            sh.Cells[row, column + 3] = data.ETF_Data.Asset_Allocation.Bond.NetAssetsPercent;

            return row;
        }

        public static int PrintEtfWorldRegions(FundamentalData data, Excel.Range range)
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

        public static int PrintEtfSectorWeights(FundamentalData data, Excel.Range range)
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
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.ConsumerCyclicals.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.ConsumerCyclicals.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Financial Services";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.FinancialServices.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.FinancialServices.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Real Estate";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.RealEstate.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.RealEstate.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "Communication Services";
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.CommunicationServices.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.CommunicationServices.RelativeToCategory;
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
            sh.Cells[row, column + 1] = data.ETF_Data.Sector_Weights.ConsumerDefensive.EquityPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Sector_Weights.ConsumerDefensive.RelativeToCategory;
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

        public static int PrintEtfFixedIncome(FundamentalData data, Excel.Range range)
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
            sh.Cells[row, column + 1] = data.ETF_Data.Fixed_Income.EffectiveDuration.FundPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Fixed_Income.EffectiveDuration.RelativeToCategory;
            row++;

            //sh.Cells[row, column] = "ModifiedDuration";
            //sh.Cells[row, column + 1] = data.ETF_Data.Fixed_Income.ModifiedDuration.FundPercent;
            //sh.Cells[row, column + 2] = data.ETF_Data.Fixed_Income.ModifiedDuration.RelativeToCategory;
            //row++;

            sh.Cells[row, column] = "EffectiveMaturity";
            sh.Cells[row, column + 1] = data.ETF_Data.Fixed_Income.EffectiveMaturity.FundPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Fixed_Income.EffectiveMaturity.RelativeToCategory;
            row++;

            sh.Cells[row, column] = "YieldToMaturity";
            sh.Cells[row, column + 1] = data.ETF_Data.Fixed_Income.YieldToMaturity.FundPercent;
            sh.Cells[row, column + 2] = data.ETF_Data.Fixed_Income.YieldToMaturity.RelativeToCategory;

            return row;
        }

        public static int PrintEtfValuationsGrowth(FundamentalData data, Excel.Range range)
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
            sh.Cells[row, column + 1] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.PriceToProspectiveEarnings;
            sh.Cells[row, column + 2] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.PriceToBook;
            sh.Cells[row, column + 3] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.PriceToSales;
            sh.Cells[row, column + 4] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.PriceToCashFlow;
            sh.Cells[row, column + 5] = data.ETF_Data.Valuations_Growth.Valuations_Rates_Portfolio.DividendYieldFactor;
            row++;

            sh.Cells[row, column] = "Valuations Rates To Category";
            sh.Cells[row, column + 1] = data.ETF_Data.Valuations_Growth.Valuations_Rates_To_Category.PriceToProspectiveEarnings;
            sh.Cells[row, column + 2] = data.ETF_Data.Valuations_Growth.Valuations_Rates_To_Category.PriceToBook;
            sh.Cells[row, column + 3] = data.ETF_Data.Valuations_Growth.Valuations_Rates_To_Category.PriceToSales;
            sh.Cells[row, column + 4] = data.ETF_Data.Valuations_Growth.Valuations_Rates_To_Category.PriceToCashFlow;
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

        public static int PrintEtfMorningStar(FundamentalData data, Excel.Range range)
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

        public static int PrintEtfPerformance(FundamentalData data, Excel.Range range)
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

        /// <summary>
        /// Загрузка всех фундаментальных данных на лист Excel
        /// </summary>
        /// <param name="data"></param>
        public static void PrintFundamentalAll(FundamentalData data, string ticker)
        {
            try
            {
                SetNonInteractive();

                string nameSheet = $"{ticker}-Fundamentals";

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

                row = PrintFundamentalGeneral(data, sh.Cells[row, 1]);
                row++;

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
            sh.Range[sh.Cells[row, range.Column], sh.Cells[row + countValues - 1, column - 1]].Value = val;
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
        /// Печать списка холдингов
        /// </summary>
        /// <typeparam name="T">Тип данных в сиске</typeparam>
        /// <param name="range">Целевая ячейка</param>
        /// <param name="data">Список холдингов</param>
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

        #region +++Bulk Fundamentals API+++
        public static void PrintBulkFundamentals(Dictionary<string, BulkFundamentalData> data)
        {
            try
            {
                SetNonInteractive();

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

        private static object[,] FillRowBalanceSheet(object[,] table, EOD.Model.BulkFundamental.Balance_SheetData item, PropertyInfo[] properties, int i)
        {
            int j = 0;
            foreach (var prop in properties)
            {
                table[i, j] = prop.GetValue(item);
                j++;
            }
            return table;
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

        private static object[,] FillRowCashFlow(object[,] table, EOD.Model.BulkFundamental.Cash_FlowData item, PropertyInfo[] properties, int i)
        {
            int j = 0;
            foreach (var prop in properties)
            {
                table[i, j] = prop.GetValue(item);
                j++;
            }
            return table;
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

        private static object[,] FillRowIncomeStatement(object[,] table, EOD.Model.BulkFundamental.Income_StatementData item, PropertyInfo[] properties, int i)
        {
            int j = 0;
            foreach (var prop in properties)
            {
                table[i, j] = prop.GetValue(item);
                j++;
            }
            return table;
        }
        #endregion
        /// <summary>
        /// 
        /// </summary>
        /// <param name="screener"></param>
        static int screenerCounter = 1;
        static int rowGeneral = 3;
        static int rowGeneralTicker = 3;
        static int rowEarnings = 3;
        static int rowEarningsTicker = 3;
        static int rowBigTables = 3;
        static int rowBigTablesTicker = 3;

        public static void PrintScreener(EOD.Model.Screener.StockMarkerScreener screener)
        {
            try
            {
                SetNonInteractive();
                Worksheet sh=new Worksheet();
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
                object[,] val = new object[screener.Data.Count+1, 15];
                int i = 1;
                val[0,0] = "Code";
                val[0,1] = "Name";
                val[0,2] = "Last day data date";
                val[0,3] = "Adjusted Close";
                val[0,4] = "Refund ID";
                val[0,5] = "Exchange";
                val[0,6] = "Currency symbol";
                val[0,7] = "Market Capitalization";
                val[0,8] = "Earnings Share";
                val[0,9] = "Dividend yield";
                val[0,10] = "Sector";
                val[0,11] = "Industry";
                if (screener.Data.Count==0)
                {
                    MessageBox.Show("No matches", "Error",  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                foreach (var item in screener.Data)
                {
                    int j = 0;
                    val[i, j] = item.Code; j++;
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
                string endpoint = "L"+Convert.ToString(i-1);
                MakeTable("A1", endpoint, sh,"Screener result", 9);            }
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
        #region +++Screener Bulk Printer+++
        private static Worksheet CreateGeneralWorksheet(string sheetName)
        {
            Worksheet sh = new Worksheet();
            sh = AddSheet("General data for "+ sheetName);
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
            sh.Cells[row, column+1] = "Date";
            sh.Cells[row, column + 2] = "EPS Actual";
            sh.Cells[row, column + 3] = "EPS Estimate";
            sh.Cells[row, column + 4] = "EPS Difference";
            sh.Cells[row, column + 5] = "Surprise Percent";
            row++;
            return sh;
        }
        private static  Worksheet CreateBalanceWorksheet(string sheetName)
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
            List<(string,string)> tickers = new List<(string, string)>();
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
            while (cellValue!=null)
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
        private static void MakeTable (string start, string end, Worksheet sh, string tableName, int tableStyle)
        {
            var range = sh.get_Range(start, end);
            ListObject tbl = sh.ListObjects.AddEx(
                SourceType: XlListObjectSourceType.xlSrcRange,
                Source: range
                );
            tbl.Name = tableName;
            switch (tableStyle)
            {
                case 1:
                    tbl.TableStyle = "TableStyleLight1";
                    break;
                case 2:
                    tbl.TableStyle = "TableStyleLight2";
                    break;
                case 3:
                    tbl.TableStyle = "TableStyleLight3";
                    break;
                case 4:
                    tbl.TableStyle = "TableStyleLight4";
                    break;
                case 5:
                    tbl.TableStyle = "TableStyleLight5";
                    break;
                case 6:
                    tbl.TableStyle = "TableStyleLight6";
                    break;
                case 7:
                    tbl.TableStyle = "TableStyleLight7";
                    break;
                case 8:
                    tbl.TableStyle = "TableStyleLight8";
                    break;
                default:
                    tbl.TableStyle = "TableStyleLight9";
                    break;
            }
        }
        public  static void PrintScreenerBulk ()
        {
            Worksheet shGeneral = new Worksheet();
            Worksheet shEarnings = new Worksheet();
            Worksheet shBalance = new Worksheet();
            Worksheet shCashFlow = new Worksheet();
            Worksheet shIncomeStatement = new Worksheet();
            shGeneral = Globals.ThisAddIn.Application.ActiveSheet;
           if(!CheckIsScreenerResult(shGeneral))
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
            Dictionary<string,BulkFundamentalData> res;
            foreach (string exchange in exchanges)
            {
                foreach ((string,string) tickerAndExchange in tickersAndExchanges)
                {
                    if (tickerAndExchange.Item2==exchange)
                    {
                        tickers.Add(tickerAndExchange.Item1);
                    }
                }
                res = GetBulkFundamental.GetBulkData(exchange, tickers, offset, 500).Result;
                PrintBulkFundamentalForScreener(res, tickers,shGeneral, shEarnings, shBalance,shCashFlow, shIncomeStatement);
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
                for (int i = 0; i<tickers.Count; i++)
                {
                    shGeneral.Cells[rowGeneralTicker, 1] = tickers[i];
                    rowGeneralTicker++;
                    for (int j = 0; j <4; j++)
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
                    columns=PrintScreenerBulkGeneral(symbol, shGeneral, columns);
                    columns=PrintScreenerBulkHighlights(symbol, shGeneral, columns);
                    columns=PrintScreenerBulkValuation(symbol, shGeneral, columns);
                    columns=PrintScreenerBulkTechnicals(symbol, shGeneral, columns);
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
            sh.Cells[rowGeneral, column ] = data.Highlights.MarketCapitalization;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.EBITDA;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.PERatio;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.PEGRatio;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.WallStreetTargetPrice;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.BookValue;
            column++;

            sh.Cells[rowGeneral, column] = data.Highlights.DividendShare;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.DividendYield;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.EarningsShare;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.EPSEstimateCurrentYear;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.EPSEstimateNextYear;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.EPSEstimateNextQuarter;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.MostRecentQuarter;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.ProfitMargin;
            column++;

            sh.Cells[rowGeneral, column ] = data.Highlights.OperatingMarginTTM;
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
            sh.Cells[1,column].Font.Bold = true;
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

            sh.Range[sh.Cells[rowBigTables, 2], sh.Cells[rowBigTables + countValues -1, properties.Length+1]].Value = val;
            rowBigTables = row;
        }
        private static void PrintCashFlowForScreener (BulkFundamentalData data, Worksheet sh)
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
        #endregion


        public static void PrintScreenerHistorical(DateTime from, DateTime to, string period)
        {

            int row = 3;
            Worksheet sh = new Worksheet();
            sh = Globals.ThisAddIn.Application.ActiveSheet;
            if (!CheckIsScreenerResult(sh))
            {
                return;
            }
            string screenerSheetName = sh.Name;

            List < (string, string) > tickers= GetTickersAndExchangesFromScreener(sh);
            sh=CreateScreenerHictoricalWorksheet(sh.Name);
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
            MakeTable("A2", "K" + Convert.ToString(row), sh,sh.Name, 9);
        }
        private static Worksheet CreateScreenerHictoricalWorksheet(string sheetName)
        {
            Worksheet sh = new Worksheet();
            sh = AddSheet("Historical data for " + sheetName);
            int column = 1;
            int row = 1;
            sh.Cells[row, column] = "Hictorical data"; 
            sh.Cells[row, column].Font.Bold = true; row++;
            sh.Cells[row, column] = "Ticker";column++;
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
    }
}