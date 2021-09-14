using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Utils
{
    public class APIException : Exception
    {
        public enum Error
        {
            Success = 200,
            Unauthenticated = 401,
            Forbidden = 403,
            NotFound = 404,
            
        }

        public int Code { get; private set; }

        public APIException(int code, string message) : base(message)
        {
            Code = code;
            
        }
    }
}
