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
    }
}
