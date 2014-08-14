using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelLinker
{
    public class ExcelOperator
    {
        public static Excel.Application oXL;
        public List<Workbook> opWorkbooks = new List<Workbook>();
        public bool isExcelOpen { get; private set; }
        private Excel._Workbook operatorWB;
        

        public void OpenWorkbook(string path = null)
        {
            if (path != null)
            {
                operatorWB = oXL.Workbooks.Open(path);
                Workbook _workbook = new Workbook(operatorWB);
                opWorkbooks.Add(_workbook);
            }

            operatorWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            Workbook newWorkbook = new Workbook(operatorWB);
            opWorkbooks.Add(newWorkbook);
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

                //Get a new workbook.
                operatorWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                OpenWorkbook();             
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
	}   
}
