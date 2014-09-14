using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class DirectoryFormatter : IFormatter
    {
        public void Format()
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Interop.Excel._Worksheet worksheet
        {
            get;
            set;
        }
    }
}
