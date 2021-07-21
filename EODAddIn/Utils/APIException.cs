using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Utils
{
    public class APIException : Exception
    {
        public int Code { get; private set; }

        public APIException(int code)
        {
            Code = code;
        }
    }
}
