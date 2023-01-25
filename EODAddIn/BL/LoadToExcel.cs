using EODAddIn.Model;
using EODAddIn.Program;
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
                errorReport.MessageToUser("Не получилось создать лист.");
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
        public static void PrintIntraday(List<Intraday> intraday, string ticker, string interval, bool chart, int period)
        {
            try
            {
                SetNonInteractive();

                string nameSheet;
                if (period == 0)
                {
                    nameSheet = $"{ticker}-{interval}";
                }
                else
                {
                    nameSheet = $"{ticker}-{period}m";
                }

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
    }
}
