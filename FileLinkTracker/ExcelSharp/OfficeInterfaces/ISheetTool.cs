using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    public interface ISheetTool
    {
        Excel._Worksheet worksheet { get; }
    }
}
