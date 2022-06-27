using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class Balance_Sheet
    {
        public string Currency_symbol { get; set; }
        public Dictionary<DateTime, Balance_SheetData> Quarterly { get; set; }
        public Dictionary<DateTime, Balance_SheetData> Yearly { get; set; }
    }
}
