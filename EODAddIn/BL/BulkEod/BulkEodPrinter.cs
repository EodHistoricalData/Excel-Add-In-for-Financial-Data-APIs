using EODAddIn.Utils;

using System;
using System.Collections.Generic;
using System.Windows.Forms;

using static EODAddIn.Utils.ExcelUtils;

using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.BL.BulkEod
{
    internal class BulkEodPrinter
    {
        public static void PrintBulkEod(List<EOD.Model.Bulks.Bulk> data, string exchange, DateTime date, string tickers, string type)
        {
            if (data.Count == 0)
            {
                MessageBox.Show("There is no data to fill in the table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                type = type == null ? "end-of-day data" : type;
                string sheetName = GetWorksheetNewName(exchange + " Bulk (" + type + ")");
                Excel.Worksheet worksheet = AddSheet(sheetName);
                int row = 1;
                int column = 1;

                OnStart();

                switch (type)
                {
                    case "end-of-day data":
                        {
                            // header
                            worksheet.Cells[row, column] = "Code";
                            worksheet.Cells[row, column + 1] = "Exchange Short Name";
                            worksheet.Cells[row, column + 2] = "Date";
                            worksheet.Cells[row, column + 3] = "Open";
                            worksheet.Cells[row, column + 4] = "High";
                            worksheet.Cells[row, column + 5] = "Low";
                            worksheet.Cells[row, column + 6] = "Close";
                            worksheet.Cells[row, column + 7] = "Adjusted Close";
                            worksheet.Cells[row, column + 8] = "Volume";
                            row++;
                            // data
                            foreach (EOD.Model.Bulks.Bulk item in data)
                            {
                                worksheet.Cells[row, column] = item.Code;
                                worksheet.Cells[row, column + 1] = item.Exchange_Short_Name;
                                worksheet.Cells[row, column + 2] = item.Date;
                                worksheet.Cells[row, column + 3] = item.Open;
                                worksheet.Cells[row, column + 4] = item.High;
                                worksheet.Cells[row, column + 5] = item.Low;
                                worksheet.Cells[row, column + 6] = item.Close;
                                worksheet.Cells[row, column + 7] = item.Adjusted_Close;
                                worksheet.Cells[row, column + 8] = item.Volume;
                                row++;
                            }

                            worksheet.UsedRange.EntireColumn.AutoFit();
                            break;
                        }
                    case "splits":
                        {
                            // header
                            worksheet.Cells[row, column] = "Code";
                            worksheet.Cells[row, column + 1] = "Exchange";
                            worksheet.Cells[row, column + 2] = "Date";
                            worksheet.Cells[row, column + 3] = "Split";
                            row++;
                            // data
                            foreach (EOD.Model.Bulks.Bulk item in data)
                            {
                                worksheet.Cells[row, column] = item.Code;
                                worksheet.Cells[row, column + 1] = item.Exchange;
                                worksheet.Cells[row, column + 2] = item.Date;
                                worksheet.Cells[row, column + 3] = item.Split;
                                row++;
                            }
                            worksheet.UsedRange.EntireColumn.AutoFit();
                            break;
                        }
                    case "dividends":
                        {
                            // header
                            worksheet.Cells[row, column] = "Code";
                            worksheet.Cells[row, column + 1] = "Exchange";
                            worksheet.Cells[row, column + 2] = "Date";
                            worksheet.Cells[row, column + 3] = "Dividend";
                            worksheet.Cells[row, column + 4] = "Currency";
                            worksheet.Cells[row, column + 5] = "Declaration Date";
                            worksheet.Cells[row, column + 6] = "Record Date";
                            worksheet.Cells[row, column + 7] = "Payment Date";
                            worksheet.Cells[row, column + 8] = "Period";
                            worksheet.Cells[row, column + 9] = "Unadjusted Value";
                            row++;
                            // data
                            foreach (EOD.Model.Bulks.Bulk item in data)
                            {
                                worksheet.Cells[row, column] = item.Code;
                                worksheet.Cells[row, column + 1] = item.Exchange;
                                worksheet.Cells[row, column + 2] = item.Date;
                                worksheet.Cells[row, column + 3] = item.Dividend;
                                worksheet.Cells[row, column + 4] = item.Currency;
                                worksheet.Cells[row, column + 5] = item.DeclarationDate;
                                worksheet.Cells[row, column + 6] = item.RecordDate;
                                worksheet.Cells[row, column + 7] = item.PaymentDate;
                                worksheet.Cells[row, column + 8] = item.Period;
                                worksheet.Cells[row, column + 9] = item.UnadjustedValue;
                                row++;
                            }
                            worksheet.UsedRange.EntireColumn.AutoFit();
                            break;
                        }
                }
            }
            catch
            {

            }
            finally
            {
                OnEnd();
            }
        }
    }
}
