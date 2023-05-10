using EOD.Model;
using EODAddIn.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using static EODAddIn.Utils.ExcelUtils;
using EODAddIn.BL.Screener;

namespace EODAddIn.BL.HistoricalPrinter
{
    public class HistoricalPrinter
    {
        /// <summary>
        ///Print end of day data to a worksheet
        /// </summary>
        /// <param name="endOfDays">List with data</param>
        /// <param name="ticker">Ticker</param>
        /// <param name="period">Period</param>
        /// <param name="chart"> necessity of chart</param>
        public static int PrintEndOfDay(List<EOD.Model.HistoricalStockPrice> endOfDays, string ticker, string period, bool chart, bool isCreateTable)
        {
            try
            {
                SetNonInteractive();
                string nameSheet = $"{ticker}-{period}";
                int r = 2;

                Excel.Worksheet worksheet = AddSheet(nameSheet);
                worksheet.Cells[r-1, 1] = "Historical Data";
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
                    foreach (HistoricalStockPrice item in endOfDays)
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
                if (isCreateTable)
                {
                    ExcelUtils.MakeTable("A2", "J" + r.ToString(), worksheet, "Intraday", 9);
                }
                if (!CreateSheet) return r;
                if (!chart) return r;

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
                return r;
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

        public static int PrintEndOfDaySummary(List<HistoricalStockPrice> res, string ticker, string period, int row)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
            int c = 1;
            int r = 1;
            sh.Cells[r, c] = "Hictorical data";
            sh.Cells[r, c].Font.Bold = true; r++;
            sh.Cells[r, c] = "Ticker"; c++;
            sh.Cells[r, c] = "Date"; c++;
            sh.Cells[r, c] = "Open"; c++;
            sh.Cells[r, c] = "High"; c++;
            sh.Cells[r, c] = "Low"; c++;
            sh.Cells[r, c] = "Close"; c++;
            sh.Cells[r, c] = "Adjusted open"; c++;
            sh.Cells[r, c] = "Adjusted high"; c++;
            sh.Cells[r, c] = "Adjusted low"; c++;
            sh.Cells[r, c] = "Adjusted close"; c++;
            sh.Cells[r, c] = "Volume"; c++;
            foreach (HistoricalStockPrice item in res)
            {
                sh.Cells[row, 1] = ticker;
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
            return row;
        }
    }
}
