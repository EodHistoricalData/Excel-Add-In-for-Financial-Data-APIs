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

namespace EODAddIn.BL.IntradayPrinter
{
    public class IntradayPrinter
    {

        /// <summary>
        /// Print intraday data to a worksheet
        /// </summary>
        /// <param name="endOfDays">List of data</param>
        /// <param name="ticker">Ticker</param>
        /// <param name="period">period</param>
        /// <param name="chart">necessity of chart</param>
        public static void PrintIntraday(List<EOD.Model.IntradayHistoricalStockPrice> intraday, string ticker, string interval, bool chart)
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
                    foreach (EOD.Model.IntradayHistoricalStockPrice item in intraday)
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
    }
}
