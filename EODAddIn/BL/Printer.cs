using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace EODAddIn.BL
{
    public class Printer
    {
        /// <summary>
        /// Print simple table
        /// </summary>
        /// <param name="header">Table name</param>
        /// <param name="value">Table values</param>
        /// <param name="firstCell">First cell of the table</param>
        /// <returns></returns>
        public static int PrintVerticalTable(string header, object value, Range firstCell)
        {
            Worksheet worksheet = firstCell.Parent;
            int row = firstCell.Row;
            int column = firstCell.Column;

            // print header
            worksheet.Cells[row, column] = header;
            worksheet.Cells[row, column].Font.Bold = true;
            row++;

            // collect table values
            Type myType = value.GetType();
            IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());
            List<(string, object)> pairs = new List<(string, object)>();
            foreach (PropertyInfo prop in props)
            {
                object propValue = prop.GetValue(value, null);
                if (propValue == null) continue;
                pairs.Add((prop.Name, propValue));
            }
            object[,] array = new object[pairs.Count, 2];
            for (int i = 0; i < pairs.Count; i++)
            {
                array[i, 0] = pairs[i].Item1;
                array[i, 1] = pairs[i].Item2;
            }
            // print table
            Range c1 = (Range)worksheet.Cells[row, 1];
            row += pairs.Count - 1;
            Range c2 = (Range)worksheet.Cells[row, 2];
            Range table = worksheet.get_Range(c1, c2);
            table.Value = array;
            table.EntireColumn.AutoFit();
            return row;
        }

        /// <summary>
        /// Prints 2-dimentional table on worksheet
        /// </summary>
        /// <param name="firstCell">first cell of table</param>
        /// <param name="array">values of table</param>
        /// <returns>Row after table</returns>
        public static int PrintHorisontalTable(Range firstCell, object[,] array)
        {
            int row = firstCell.Row;
            Worksheet worksheet = firstCell.Worksheet;
            int rows = array.GetUpperBound(0) + 1;
            int columns = array.Length / rows;
            Range lastCell = worksheet.Cells[row + rows - 1, columns];
            Range rng = worksheet.get_Range(firstCell, lastCell);
            rng.Value = array;
            rng.EntireColumn.AutoFit();
            return row + rows;
        }
    }
}
