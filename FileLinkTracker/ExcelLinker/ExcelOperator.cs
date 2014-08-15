using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelSharp
{
    public class ExcelOperator : IDisposable
    {
        public static Excel.Application oXL;
        public bool isExcelOpen { get; private set; }        
        private Excel._Workbook operatorWB;
        public int WorkbookCount
        {
            get { return oXL.Workbooks.Count; }
        }
        

        public Workbook OpenWorkbook(string path = null)
        {
            if (path != null)
            {
                operatorWB = oXL.Workbooks.Open(path);
                Workbook _workbook = new Workbook(operatorWB);                
                return _workbook;
            }

            else
            {
                operatorWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                Workbook _workbook = new Workbook(operatorWB);                
                return _workbook;
            }
        }

        public ExcelOperator()
        {
            isExcelOpen = false;
        }

        public void InitializeExcel()
        {
            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;
                isExcelOpen = true;            
            }

            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                Console.WriteLine(errorMessage);
            }
        }

        public void CloseExcel()
        {
            oXL.Quit();
            isExcelOpen = false;
        }

        public void CloseWorkbook()
        {
            oXL.Visible = true;
            oXL.DisplayAlerts = false;
            oXL.Workbooks.Close();
            oXL.Visible = false;
        }

        public void Dispose()
        {
            CloseWorkbook();
            CloseExcel();
        }
    }   
}
