using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace ExcelSolution.Code
{
    public class ExcelChartApp: IExcelApp
    {
        public bool IsOpen { get; private set; }
        private static Excel.Application MyApp = null;
        public static string DB_PATH = @"";
        private static Excel.Workbook MyBook = null;
        private static Excel.Worksheet MySheet = null;
        private static int lastRow=0;       
        
        public void InitializeExcel()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explict cast is not required here
            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            IsOpen = true;
        }
        public void CloseExcel()
        {
            MyApp.Quit();            
            this.IsOpen = false;
        }
    }
}
