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

namespace EODAddIn.BL.OptionsPrinter
{
    public class OptionsPrinter
    {
        /// <summary>
        /// Print Options Data to a worksheet
        /// </summary>
        /// <param name="data">Данные</param>
        public static void PrintOptions(OptionsData data, string ticker)
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



    }
}
