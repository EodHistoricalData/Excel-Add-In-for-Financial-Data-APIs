using EOD.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

using static EODAddIn.Utils.ExcelUtils;
using Excel = Microsoft.Office.Interop.Excel;

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
        public static int PrintEndOfDay(List<HistoricalStockPrice> endOfDays, string ticker, string period, bool chart, bool isCreateTable)
        {
            try
            {
                SetNonInteractive();

                string nameSheet = GetWorksheetNewName($"{ticker} EOD {period} {endOfDays.First().Date:d}");

                int r = 1;

                Worksheet worksheet = AddSheet(nameSheet);

                object[,] table = new object[endOfDays.Count + 1, 10];
                table[0, 0] = "Date";
                table[0, 1] = "Open";
                table[0, 2] = "High";
                table[0, 3] = "Low";
                table[0, 4] = "Close";
                table[0, 5] = "Adjusted_open";
                table[0, 6] = "Adjusted_high";
                table[0, 7] = "Adjusted_lowe";
                table[0, 8] = "Adjusted_close";
                table[0, 9] = "Volume";
                try
                {
                    OnStart();
                    int i = 0;
                    foreach (HistoricalStockPrice item in endOfDays)
                    {
                        i++;
                        table[i, 0] = item.Date;
                        table[i, 1] = item.Open;
                        table[i, 2] = item.High;
                        table[i, 3] = item.Low;
                        table[i, 4] = item.Close;
                        table[i, 5] = item.Adjusted_open;
                        table[i, 6] = item.Adjusted_high;
                        table[i, 7] = item.Adjusted_low;
                        table[i, 8] = item.Adjusted_close;
                        table[i, 9] = item.Volume;
                    }
                    r = Printer.PrintHorisontalTable(worksheet.Cells[1, 1], table);

                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    OnEnd();
                }
                if (isCreateTable)
                {
                    MakeTable("A1", "J" + (r - 1).ToString(), worksheet, "EOD", 9);
                }

                if (!CreateSheet) return r;

                if (!chart) return r;

                DrawCharts(worksheet, r);

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

        public static int PrintEndOfDaySummary(List<HistoricalStockPrice> res, string ticker, string period, int row, Worksheet worksheet)
        {
            if (row == 2)
            {
                worksheet.Cells[1, 1] = "Ticker";
                worksheet.Cells[1, 2] = "Date";
                worksheet.Cells[1, 3] = "Open";
                worksheet.Cells[1, 4] = "High";
                worksheet.Cells[1, 5] = "Low";
                worksheet.Cells[1, 6] = "Close";
                worksheet.Cells[1, 7] = "Adjusted open";
                worksheet.Cells[1, 8] = "Adjusted high";
                worksheet.Cells[1, 9] = "Adjusted low";
                worksheet.Cells[1, 10] = "Adjusted close";
                worksheet.Cells[1, 11] = "Volume";
            }

            object[,] data = new object[res.Count, 11];

            int i = 0;
            foreach (HistoricalStockPrice item in res)
            {
                data[i, 0] = ticker;
                data[i, 1] = item.Date;
                data[i, 2] = item.Open;
                data[i, 3] = item.High;
                data[i, 4] = item.Low;
                data[i, 5] = item.Close;
                data[i, 6] = item.Adjusted_open;
                data[i, 7] = item.Adjusted_high;
                data[i, 8] = item.Adjusted_low;
                data[i, 9] = item.Adjusted_close;
                data[i, 10] = item.Volume;
                i++;
            }
            worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row + i - 1, 11]] = data;
            worksheet.UsedRange.EntireColumn.AutoFit();

            return row + i;
        }

        private static void DrawCharts(Worksheet worksheet, int row)
        {
            int lastCol = worksheet.UsedRange.Columns.Count;
            var charts = worksheet.ChartObjects();
            var chartObject = charts.Add(60, 10, 300, 300);
            Chart chart = chartObject.Chart;
            chart.ChartType = XlChartType.xlLine;
            Range range = worksheet.get_Range((Range)worksheet.Cells[1, 10], (Range)worksheet.Cells[row, lastCol]);
            chart.SetSourceData(range);
            range = worksheet.get_Range((Range)worksheet.Cells[2, 1], (Range)worksheet.Cells[row, 1]);
            chart.FullSeriesCollection(1).XValues = range;
            chartObject.Left = (float)worksheet.Cells[1, lastCol + 1].Left;
            chartObject.Top = (float)worksheet.Cells[1, lastCol + 1].Top + 340.157480315f;
            chartObject.Height = 340.157480315f;
            chartObject.Width = 680.3149606299f;

            worksheet.Range["A2:E3"].Select();
            chartObject = worksheet.Shapes.AddChart2(-1, XlChartType.xlStockOHLC);
            chart = chartObject.Chart;
            chart.ChartGroups(1).UpBars.Format.Fill.ForeColor.RGB = Color.FromArgb(0, 176, 80);
            chart.ChartGroups(1).DownBars.Format.Fill.ForeColor.RGB = Color.FromArgb(255, 0, 0);
            range = worksheet.get_Range((Range)worksheet.Cells[1, 1], (Range)worksheet.Cells[row - 1, 5]);
            chart.SetSourceData(range);

            chartObject.Left = (float)worksheet.Cells[1, lastCol + 1].Left;
            chartObject.Top = (float)worksheet.Cells[1, lastCol + 1].Top;
            chartObject.Height = 340.157480315f;
            chartObject.Width = 680.3149606299f;
            chart.ChartTitle.Caption = worksheet.Name;
        }
    }
}
