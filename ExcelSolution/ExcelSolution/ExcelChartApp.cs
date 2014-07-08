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
    public class ExcelChartApp
    {
        public bool Visible { get; private set; }
        public bool IsOpen { get; private set; }
        private static Excel.Application MyApp = null;
        
        public void InitializeExcel()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            this.Visible = MyApp.Visible;
        }
        public void CloseExcel()
        {
            MyApp.Quit();            
            this.IsOpen = false;
        }
    }
}
