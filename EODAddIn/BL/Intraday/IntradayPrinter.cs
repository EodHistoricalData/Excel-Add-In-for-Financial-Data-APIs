using EOD.Model;
using EODAddIn.Utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using static EODAddIn.Utils.ExcelUtils;
using EODAddIn.BL.Screener;
using System.Linq;
using System.Windows.Forms;

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
        public static int PrintIntraday(List<EOD.Model.IntradayHistoricalStockPrice> intraday, string ticker, string interval, bool chart, bool isCreateTable)
        {
            try
            {
                SetNonInteractive();

                string nameSheet = GetWorksheetNewName($"{ticker}  Intraday {interval}");
                Worksheet worksheet = AddSheet(nameSheet);

                object[,] data = new object[intraday.Count + 1, 8];
                int r = 1;
                data[0, 0] = "Timestamp";
                data[0, 1] = "Gmtoffset";
                data[0, 2] = "DateTime";
                data[0, 3] = "Open";
                data[0, 4] = "High";
                data[0, 5] = "Low";
                data[0, 6] = "Close";
                data[0, 7] = "Volume";
                int i = 0;
                foreach (IntradayHistoricalStockPrice item in intraday)
                {
                    r++;
                    i++;
                    data[i, 0] = item.Timestamp;
                    data[i, 1] = item.Gmtoffset;
                    data[i, 2] = item.DateTime;
                    data[i, 3] = item.Open;
                    data[i, 4] = item.High;
                    data[i, 5] = item.Low;
                    data[i, 6] = item.Close;
                    data[i, 7] = item.Volume;
                }
                worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[r, 8]].Value = data;
                worksheet.Range["A:H"].EntireColumn.AutoFit();

                if (isCreateTable)
                {
                    MakeTable("A1", "H" + r.ToString(), worksheet, "Intraday", 9);
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
        public static int PrintIntradaySummary(List<IntradayHistoricalStockPrice> res, string ticker, int row, Worksheet worksheet)
        {
            if (row == 2)
            {
                worksheet.Cells[1, 1] = "Ticker";
                worksheet.Cells[1, 2] = "TimeStamp";
                worksheet.Cells[1, 3] = "Gmtoffset";
                worksheet.Cells[1, 4] = "DateTime";
                worksheet.Cells[1, 5] = "Open";
                worksheet.Cells[1, 6] = "High";
                worksheet.Cells[1, 7] = "Low";
                worksheet.Cells[1, 8] = "Close";
                worksheet.Cells[1, 9] = "Volume";
            }

            object[,] data = new object[res.Count, 9];

            int i = 0;
            foreach (IntradayHistoricalStockPrice item in res)
            {
                data[i, 0] = ticker;
                data[i, 1] = item.Timestamp;
                data[i, 2] = item.Gmtoffset;
                data[i, 3] = item.DateTime;
                data[i, 4] = item.Open;
                data[i, 5] = item.High;
                data[i, 6] = item.Low;
                data[i, 7] = item.Close;
                data[i, 8] = item.Volume;
                i++;
            }

            worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row + i - 1, 9]] = data;
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
            Range range = worksheet.get_Range((Range)worksheet.Cells[1, 8], (Range)worksheet.Cells[row, lastCol]);
            chart.SetSourceData(range);
            range = worksheet.get_Range((Range)worksheet.Cells[1, 3], (Range)worksheet.Cells[row, 3]);
            chart.FullSeriesCollection(1).XValues = range;
            chart.DisplayBlanksAs = XlDisplayBlanksAs.xlInterpolated;
            chartObject.Left = (float)worksheet.Cells[1, lastCol + 1].Left;
            chartObject.Top = (float)worksheet.Cells[1, lastCol + 1].Top + 340.157480315f;
            chartObject.Height = 340.157480315f;
            chartObject.Width = 680.3149606299f;

            worksheet.Range["C1:G2"].Select();
            chartObject = worksheet.Shapes.AddChart2(-1, XlChartType.xlStockOHLC);
            chart = chartObject.Chart;
            chart.ChartGroups(1).UpBars.Format.Fill.ForeColor.RGB = Color.FromArgb(0, 176, 80);
            chart.ChartGroups(1).DownBars.Format.Fill.ForeColor.RGB = Color.FromArgb(255, 0, 0);
            range = worksheet.get_Range((Range)worksheet.Cells[1, 3], (Range)worksheet.Cells[row - 1, 7]);
            chart.SetSourceData(range);
            chart.Axes(XlAxisType.xlCategory).CategoryType = XlCategoryType.xlCategoryScale;

            chartObject.Left = (float)worksheet.Cells[1, lastCol + 1].Left;
            chartObject.Top = (float)worksheet.Cells[1, lastCol + 1].Top;
            chartObject.Height = 340.157480315f;
            chartObject.Width = 680.3149606299f;
            chart.ChartTitle.Caption = worksheet.Name;
        }
    }
}
