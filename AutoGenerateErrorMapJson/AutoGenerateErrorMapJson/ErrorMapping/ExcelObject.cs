using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AutoGenerateErrorMapJson.ErrorMapping
{
    public class ExcelObject
    {
        public string ErrorCode { get; set; }

        public string HttpStatusCode
        { get; set; }

        public string ErrorCategory
        { get; set; }

        public string ErrorType
        { get; set; }

        public string SupplierError
        { get; set; }

    }
}
