using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharpTests
{
    interface ISheetCommand
    {
        public void ExecuteWithWorkbook();
        public void ExecuteWithSheet();
    }
}
