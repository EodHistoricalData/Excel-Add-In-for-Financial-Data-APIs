using EODAddIn.Utils;

using EODHistoricalData.Wrapper.Model.TechnicalIndicators;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

using static EODAddIn.Utils.ExcelUtils;

namespace EODAddIn.BL.TechnicalIndicatorData
{
    internal class TechnicalsPrinter
    {
        public static int PrintTechnicals(List<TechnicalIndicator> data, string ticker, List<IndicatorParameters> parameters, bool chart, bool isCreateTable)
        {
            int row = 1;
            try
            {
                SetNonInteractive();
                OnStart();

                string function = parameters.First(x => x.Name == "function").Value;

                string nameSheet = GetWorksheetNewName($"{ticker} - {function}");

                Worksheet worksheet = AddSheet(nameSheet);

                worksheet.Cells[row, 1] = "Technical Indicator Data";
                worksheet.Range["A:A"].EntireColumn.AutoFit();
                row++;

                if (data.Count == 0)
                {
                    MessageBox.Show("There is no available data for the selected parameters.", "No data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                Type myType = data[data.Count - 1].GetType();
                IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());
                IList<PropertyInfo> propsFor = new List<PropertyInfo>(props);
                foreach (PropertyInfo prop in propsFor)
                {
                    object propValue = prop.GetValue(data[data.Count - 1], null);
                    if (propValue == null)
                    {
                        props.Remove(prop);
                        continue;
                    }
                }
                object[,] table = new object[data.Count + 1, props.Count];
                for (int j = 0; j < props.Count; j++)
                {
                    table[0, j] = props[j].Name.ToString();
                }
                for (int i = 1; i < data.Count + 1; i++)
                {
                    for (int j = 0; j < props.Count; j++)
                    {
                        table[i, j] = props[j].GetValue(data[i - 1], null);
                    }
                }

                row = Printer.PrintHorisontalTable(worksheet.Cells[row, 1], table);

                if (chart)
                {
                    bool candle = function == "splitadjusted" ? true : false;
                    DrawChart(worksheet, row, candle);
                }

                if (isCreateTable)
                {
                    char lastColumn = (char)('A' + worksheet.UsedRange.Columns.Count - 1);
                    MakeTable("A2", lastColumn.ToString() + row.ToString(), worksheet, function, 9);
                }

                return row;
            }
            catch
            {
                throw;
            }
            finally
            {
                OnEnd();
                _xlsApp.Interactive = true;
            }
        }

        private static void DrawChart(Worksheet worksheet, int row, bool candle)
        {
            int lastCol = worksheet.UsedRange.Columns.Count;
            if (!candle)
            {
                var charts = worksheet.ChartObjects();
                var chartObject = charts.Add(60, 10, 300, 300);
                Chart chart = chartObject.Chart;
                chart.ChartType = XlChartType.xlLine;
                Range range = worksheet.get_Range((Range)worksheet.Cells[2, 2], (Range)worksheet.Cells[row, lastCol]);
                chart.SetSourceData(range);
                range = worksheet.get_Range((Range)worksheet.Cells[2, 1], (Range)worksheet.Cells[row, 1]);
                chart.FullSeriesCollection(1).XValues = range;
                chartObject.Left = (float)worksheet.Cells[1, lastCol + 1].Left;
                chartObject.Top = (float)worksheet.Cells[1, lastCol + 1].Top;
                chartObject.Height = 340.157480315f;
                chartObject.Width = 680.3149606299f;
            }
            else
            {
                worksheet.Range["A2:E3"].Select();
                Shape chartObject = worksheet.Shapes.AddChart2(-1, XlChartType.xlStockOHLC);
                Chart chart = chartObject.Chart;
                chart.ChartGroups(1).UpBars.Format.Fill.ForeColor.RGB = Color.FromArgb(0, 176, 80);
                chart.ChartGroups(1).DownBars.Format.Fill.ForeColor.RGB = Color.FromArgb(255, 0, 0);
                Range range = worksheet.get_Range((Range)worksheet.Cells[2, 2], (Range)worksheet.Cells[row, 2]);
                chart.FullSeriesCollection(1).Values = range;
                range = worksheet.get_Range((Range)worksheet.Cells[2, 3], (Range)worksheet.Cells[row, 3]);
                chart.FullSeriesCollection(2).Values = range;
                range = worksheet.get_Range((Range)worksheet.Cells[2, 4], (Range)worksheet.Cells[row, 4]);
                chart.FullSeriesCollection(3).Values = range;
                range = worksheet.get_Range((Range)worksheet.Cells[2, 5], (Range)worksheet.Cells[row, 5]);
                chart.FullSeriesCollection(4).Values = range;
                range = worksheet.get_Range((Range)worksheet.Cells[2, 1], (Range)worksheet.Cells[row, 1]);
                chart.FullSeriesCollection(1).XValues = range;
                chartObject.Left = (float)worksheet.Cells[1, lastCol + 1].Left;
                chartObject.Top = (float)worksheet.Cells[1, lastCol + 1].Top;
                chartObject.Height = 340.157480315f;
                chartObject.Width = 680.3149606299f;
            }

        }

        public static int PrintTechnicalsSummary(List<TechnicalIndicator> data, string ticker, int row, List<IndicatorParameters> parameters, Worksheet sh)
        {
            try
            {
                SetNonInteractive();
                OnStart();

                if (row == 1)
                {
                    string function = parameters.First(x => x.Name == "function").Value;

                    row++;

                    Type myType = data[data.Count - 1].GetType();
                    IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());
                    IList<PropertyInfo> propsFor = new List<PropertyInfo>(props);
                    foreach (PropertyInfo prop in propsFor)
                    {
                        object propValue = prop.GetValue(data[data.Count - 1], null);
                        if (propValue == null)
                        {
                            props.Remove(prop);
                            continue;
                        }
                    }
                    object[,] table = new object[data.Count + 1, props.Count + 1];
                    for (int j = 0; j < props.Count + 1; j++)
                    {
                        if (j == 0)
                        {
                            table[0, j] = "Ticker";
                        }
                        else
                        {
                            table[0, j] = props[j - 1].Name.ToString();
                        }
                    }
                    for (int i = 1; i < data.Count + 1; i++)
                    {
                        for (int j = 0; j < props.Count + 1; j++)
                        {
                            if (j == 0)
                            {
                                table[i, j] = ticker;
                            }
                            else
                            {
                                table[i, j] = props[j - 1].GetValue(data[i - 1], null);
                            }
                        }
                    }
                    row = Printer.PrintHorisontalTable(sh.Cells[row, 1], table) - 1;
                }
                else
                {
                    Type myType = data[data.Count - 1].GetType();
                    IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());
                    IList<PropertyInfo> propsFor = new List<PropertyInfo>(props);
                    foreach (PropertyInfo prop in propsFor)
                    {
                        object propValue = prop.GetValue(data[data.Count - 1], null);
                        if (propValue == null)
                        {
                            props.Remove(prop);
                            continue;
                        }
                    }
                    object[,] table = new object[data.Count, props.Count + 1];
                    for (int i = 1; i < data.Count; i++)
                    {
                        for (int j = 0; j < props.Count + 1; j++)
                        {
                            if (j == 0)
                            {
                                table[i, j] = ticker;
                            }
                            else
                            {
                                table[i, j] = props[j - 1].GetValue(data[i - 1], null);
                            }
                        }
                    }
                    row = Printer.PrintHorisontalTable(sh.Cells[row, 1], table);
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                OnEnd();
                _xlsApp.Interactive = true;
            }
            return row;
        }
    }
}
