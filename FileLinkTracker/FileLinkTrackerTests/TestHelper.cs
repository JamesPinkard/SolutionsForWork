using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSharp;

namespace ExcelSharpTests
{
    public static class TestHelper
    {
        public void InitializeWorksheet(ExcelOperator testOperator, Workbook testWorkbook )
        {
            testOperator = new ExcelOperator();
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
            testWorkbook.ActiveSheet.Cells.Clear();
        }

        private Workbook initTestWorkbook(ExcelOperator testOperator)
        {
            Workbook testWorkbook = testOperator.OpenWorkbook();
            return testWorkbook;
        }
    }
}
