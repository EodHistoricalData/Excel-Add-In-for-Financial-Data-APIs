using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Model
{
    public class Income_Statement
    {
        public string Currency_symbol { get; set; }
        public Dictionary<DateTime, Income_StatementData> Quarterly { get; set; }
        public Dictionary<DateTime, Income_StatementData> Yearly { get; set; }
    }
}
