using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSharp;

namespace ExcelSharpTests
{
    public class TestHelper
    {
        public void InitializeExcel(ExcelOperator testOperator, Workbook testWorkbook )
        {            
            testOperator.InitializeExcel();
            testWorkbook = initTestWorkbook(testOperator);
        }

        public void CloseExcel(ExcelOperator testOperator)
        {
            testOperator.CloseWorkbook();
            testOperator.CloseExcel();
        }

        public void ClearCells(Workbook testWorkbook)
        {
            testWorkbook.ActiveSheet.ClearCells();
        }

        private Workbook initTestWorkbook(ExcelOperator testOperator)
        {
            Workbook testWorkbook = testOperator.OpenWorkbook();
            return testWorkbook;
        }
    }
}
