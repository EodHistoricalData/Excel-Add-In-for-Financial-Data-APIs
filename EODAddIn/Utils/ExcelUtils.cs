using EODAddIn.Program;
using Microsoft.Office.Interop.Excel;
using System;
using System.Reflection;
using Form = System.Windows.Forms;
namespace EODAddIn.Utils
{
    static class ExcelUtils
    {
        public static Application _xlsApp = Globals.ThisAddIn.Application;
        public static bool CreateSheet=true;
        private static XlCalculation Calculation = XlCalculation.xlCalculationAutomatic;

        public static void OnStart()
        {
            Application app = Globals.ThisAddIn.Application;
            Calculation = app.Calculation;
            app.ScreenUpdating = false;
            app.Calculation = XlCalculation.xlCalculationManual;
        }

        public static void OnEnd()
        {
            Application app = Globals.ThisAddIn.Application;
            app.ScreenUpdating = true;
            app.Calculation = Calculation;
        }

        public static bool IsRange(string rangeAddress)
        {
            Application app = Globals.ThisAddIn.Application;
            try
            {
                Range range = app.Range[rangeAddress];
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool CheckSheetName(string name)
        {
            if (name == string.Empty) return false;
            if (name.Length > 31) return false;
            if (name.Contains("/")) return false;
            if (name.Contains(@"\")) return false;
            if (name.Contains(@"?")) return false;
            if (name.Contains(@"*")) return false;
            if (name.Contains(@":")) return false;
            if (name.Contains(@"[")) return false;
            if (name.Contains(@"]")) return false;
            if (name[0] == '\'') return false;
            if (name[name.Length - 1] == '\'') return false;
            if (name == "History") return false;
            return true;
        }

        public static bool SheetExists(string name)
        {
            try
            {
                Worksheet worksheet = Globals.ThisAddIn.Application.Worksheets[name];
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Counting the number of rows on a worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static int RowsCount(Worksheet worksheet)
        {
            return worksheet.UsedRange.Row - 1 + worksheet.UsedRange.Rows.Count;
        }

        /// <summary>
        /// Counting the number of columns on a worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static int ColumnsCount(Worksheet worksheet)
        {
            return worksheet.UsedRange.Column - 1 + worksheet.UsedRange.Columns.Count;
        }

        public enum SelectionType
        {
            Range,
            Shape,
            None
        }

        /// <summary>
        /// Determining the type of the selected range
        /// </summary>
        /// <param name="selection">Object of Application.Selection</param>
        /// <returns>Range or Shape or Null</returns>
        public static SelectionType TypeName(dynamic selection)
        {
            if (selection is Range)
            {
                return SelectionType.Range;
            }
            if (selection is Shape)
            {
                return SelectionType.Shape;
            }
            return SelectionType.None;
        }

        /// <summary>
        /// Search in a selected range of cells with formulas
        /// </summary>
        /// <param name="range">Cell Range</param>
        /// <returns>Visible cells with formulas. Null if there are no such cells. It also throws an error message if range cannot be obtained for the reason COM</returns>
        public static Range GetVisibleFormulas(Range range)
        {
            Application application = Globals.ThisAddIn.Application;

            if (range.CountLarge == 1 || application.ActiveCell.MergeArea.CountLarge == range.CountLarge)
            {
                range = application.ActiveCell;
                if (!range.HasFormula)
                {
                    Form.MessageBox.Show("No formulas in selected range", "Select Range", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Information);
                    return null;
                }
            }
            else
            {
                try
                {
                    range = range.SpecialCells(XlCellType.xlCellTypeVisible);
                    range = range.SpecialCells(XlCellType.xlCellTypeFormulas);
                }
                catch (Exception ex)
                {
                    Form.MessageBox.Show(ex.Message, "Select Range", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Warning);
                    return null;
                }
            }
            return range;
        }

        /// <summary>
        /// Search in a selected range of cells with text
        /// </summary>
        /// <param name="range">Cell Range</param>
        /// <returns>Visible cells with text. Null если таких ячеек нет. Null if there are no such cells. It also throws an error message if range cannot be obtained for the reason COM</returns>
        public static Range GetVisibleConstants(Range range)
        {
            Application application = Globals.ThisAddIn.Application;

            if (range.CountLarge == 1 || application.ActiveCell.MergeArea.CountLarge == range.CountLarge)
            {
                range = application.ActiveCell;
                if (range.HasFormula)
                {
                    Form.MessageBox.Show("No text cells in specified range",
                            "Select Range", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Information);
                    return null;
                }
            }
            else
            {
                try
                {
                    range = range.SpecialCells(XlCellType.xlCellTypeVisible);
                    range = range.SpecialCells(XlCellType.xlCellTypeConstants);
                }
                catch (Exception ex)
                {
                    Form.MessageBox.Show(ex.Message, "Select Range", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Warning);
                    return null;
                }
            }
            return range;
        }

        /// <summary>
        /// Message about the impossibility to undo this action
        /// </summary>
        /// <returns>DialogResult Yes/No</returns>
        public static Form.DialogResult MessageConfirmAction()
        {
            return Form.MessageBox.Show("This action cannot be undone\nAre you sure you want to continue?", "Confirm", Form.MessageBoxButtons.YesNo, Form.MessageBoxIcon.Question);
        }

        /// <summary>
        ///Message that a range of cells is not selected
        /// </summary>
        /// <returns></returns>
        public static bool MessageIfNoRange()
        {
            Application application = Globals.ThisAddIn.Application;
            if (ExcelUtils.TypeName(application.Selection) != SelectionType.Range)
            {
                Form.MessageBox.Show("Select a range of cells",
                    "Select a range of cells", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Information);
                return true;
            }
            return false;
        }

        public static bool SheetDeleteNoMessage(Worksheet worksheet)
        {
            Application application = Globals.ThisAddIn.Application;
            try
            {
                application.DisplayAlerts = false;
                worksheet.Delete();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                application.DisplayAlerts = true;
            }
        }

        public static void MakeTable(string start, string end, Worksheet sh, string tableName, int tableStyle)
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

        public static void SetNonInteractive()
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

        public static Worksheet AddSheet(string nameSheet)
        {
            Worksheet worksheet = null;
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

        public static object[,] FillRowBalanceSheet(object[,] table, EOD.Model.BulkFundamental.Balance_SheetData item, PropertyInfo[] properties, int i)
        {
            int j = 0;
            foreach (var prop in properties)
            {
                table[i, j] = prop.GetValue(item);
                j++;
            }
            return table;
        }
        public static object[,] FillRowCashFlow(object[,] table, EOD.Model.BulkFundamental.Cash_FlowData item, PropertyInfo[] properties, int i)
        {
            int j = 0;
            foreach (var prop in properties)
            {
                table[i, j] = prop.GetValue(item);
                j++;
            }
            return table;
        }
        public static object[,] FillRowIncomeStatement(object[,] table, EOD.Model.BulkFundamental.Income_StatementData item, PropertyInfo[] properties, int i)
        {
            int j = 0;
            foreach (var prop in properties)
            {
                table[i, j] = prop.GetValue(item);
                j++;
            }
            return table;
        }
    }
}
