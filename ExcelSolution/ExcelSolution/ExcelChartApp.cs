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
        private static Excel.Workbook MyBook = null;
        private static Excel.Worksheet MySheet = null;
             
        
        public void InitializeExcel()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Add();
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explict cast is not required here            
            IsOpen = true;
        }
        public void CloseExcel()
        {
            MyApp.Quit();            
            this.IsOpen = false;
        }
    }
}
