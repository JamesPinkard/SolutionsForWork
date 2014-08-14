using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLinker
{
    public class Workbook
    {
        public Excel._Worksheet ActiveSheet { get; private set; }
        public Excel.Range CurrentRange { get; private set; }
        private Excel._Workbook oWB;

        public Workbook( Excel._Workbook oWB)
        {
            this.oWB = oWB;
            this.ActiveSheet = (Excel._Worksheet)oWB.ActiveSheet;            
        }
    }
}
