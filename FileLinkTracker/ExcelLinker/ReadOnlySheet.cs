using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    public class ReadOnlySheet : Sheet
    {
        public ReadOnlySheet(Excel._Worksheet worksheet)
            : base(worksheet)
        {

        }

    }
}
