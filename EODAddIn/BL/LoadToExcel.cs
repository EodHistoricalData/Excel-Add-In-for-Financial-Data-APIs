using EODAddIn.Model;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.BL
{
    public class LoadToExcel
    {
        public static void LoadEndOfDay(List<EndOfDay> endOfDays)
        {
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;

            int r = 2;
            worksheet.Cells[r, 1] = "Date";
            worksheet.Cells[r, 2] = "Open";
            worksheet.Cells[r, 3] = "High";
            worksheet.Cells[r, 4] = "Low";
            worksheet.Cells[r, 5] = "Close";
            worksheet.Cells[r, 6] = "Adjusted_close";
            worksheet.Cells[r, 7] = "Volume";

            foreach (EndOfDay item in endOfDays)
            {
                r++;
                worksheet.Cells[r, 1] = item.Date;
                worksheet.Cells[r, 2] = item.Open;
                worksheet.Cells[r, 3] = item.High;
                worksheet.Cells[r, 4] = item.Low;
                worksheet.Cells[r, 5] = item.Close;
                worksheet.Cells[r, 6] = item.Adjusted_close;
                worksheet.Cells[r, 7] = item.Volume;
            }
        }
    }
}
