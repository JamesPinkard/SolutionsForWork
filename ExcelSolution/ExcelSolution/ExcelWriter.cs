using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace ExcelSolution.Code
{
    class ExcelWriter
    {
        private int lastRow = 0;
        public Excel.Worksheet ExcelSheet { get; private set; }


        public ExcelWriter(Excel.Worksheet excelSheet)
        {
            lastRow = excelSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            ExcelSheet = excelSheet;
        }
    }
}
