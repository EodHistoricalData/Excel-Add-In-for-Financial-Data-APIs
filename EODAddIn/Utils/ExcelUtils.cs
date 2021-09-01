using Microsoft.Office.Interop.Excel;

using System;

using Form = System.Windows.Forms;
namespace EODAddIn.Utils
{
    static class ExcelUtils
    {

        public static void OnStart()
        {
            Application app = Globals.ThisAddIn.Application;
            app.ScreenUpdating = false;
            app.Calculation = XlCalculation.xlCalculationManual;
        }

        public static void OnEnd()
        {
            Application app = Globals.ThisAddIn.Application;
            app.ScreenUpdating = true;
            app.Calculation = XlCalculation.xlCalculationAutomatic;
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
        /// Определение количества строк на листе
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static int RowsCount(Worksheet worksheet)
        {
            return worksheet.UsedRange.Row - 1 + worksheet.UsedRange.Rows.Count;
        }

        /// <summary>
        /// Определение количества столбцов на листе
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
        /// Определение типа выделенного диапазона
        /// </summary>
        /// <param name="selection">Объект Application.Selection</param>
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
        /// Ищет в указанном диапазоне ячеек все видимые ячейки с формулами
        /// </summary>
        /// <param name="range">Исходный диапазон ячеек</param>
        /// <returns>Видимые ячейки с формулами. Null если таких ячеек нет. Также выдает сообщение об ошибке если невозможно получить range по причине COM</returns>
        public static Range GetVisibleFormulas(Range range)
        {
            Application application = Globals.ThisAddIn.Application;

            if (range.CountLarge == 1 || application.ActiveCell.MergeArea.CountLarge == range.CountLarge)
            {
                range = application.ActiveCell;
                if (!range.HasFormula)
                {
                    Form.MessageBox.Show("Формулы в указанном диапазоне отсутствуют", "Выберите диапазон", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Information);
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
                    Form.MessageBox.Show(ex.Message, "Выберите диапазон", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Warning);
                    return null;
                }
            }
            return range;
        }

        /// <summary>
        /// Ищет в указанном диапазоне ячеек все видимые текстовые ячейки
        /// </summary>
        /// <param name="range">Исходный диапазон ячеек</param>
        /// <returns>Видимые текстовые ячейки. Null если таких ячеек нет. Также выдает сообщение об ошибке если невозможно получить range по причине COM</returns>
        public static Range GetVisibleConstants(Range range)
        {
            Application application = Globals.ThisAddIn.Application;

            if (range.CountLarge == 1 || application.ActiveCell.MergeArea.CountLarge == range.CountLarge)
            {
                range = application.ActiveCell;
                if (range.HasFormula)
                {
                    Form.MessageBox.Show("Текстовые ячейки в указанном диапазоне отсутствуют",
                            "Выберите диапазон", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Information);
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
                    Form.MessageBox.Show(ex.Message, "Выберите диапазон", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Warning);
                    return null;
                }
            }
            return range;
        }

        /// <summary>
        /// Сообщение о невозможности отменить данное действие
        /// </summary>
        /// <returns>DialogResult Yes/No</returns>
        public static Form.DialogResult MessageConfirmAction()
        {
            return Form.MessageBox.Show("Данное действие невозможно отменить\nУверены, что хотите продолжить", "Подтвердите действие", Form.MessageBoxButtons.YesNo, Form.MessageBoxIcon.Question);
        }

        /// <summary>
        /// Сообщение о том, что не выделен диапазон ячеек
        /// </summary>
        /// <returns></returns>
        public static bool MessageIfNoRange()
        {
            Application application = Globals.ThisAddIn.Application;
            if (ExcelUtils.TypeName(application.Selection) != SelectionType.Range)
            {
                Form.MessageBox.Show("Необходимо выделить диапазон ячеек",
                    "Выделите диапазон ячеек", Form.MessageBoxButtons.OK, Form.MessageBoxIcon.Information);
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
    }
}
