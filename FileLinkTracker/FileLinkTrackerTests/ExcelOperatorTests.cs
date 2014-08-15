using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelSharpTests
{    
    public class ExcelOperatorTests
    {
        ExcelOperator testOperator;
        [SetUp]
        public void InitializeExcel()
        {
            testOperator = getInitializedOperator();
        }

        [Category("Slow Running")]
        [Test]
        public void ExcelOperator_Initialize_InitializesWorksheet()
        {            
            // Set Up Test Opened

            Assert.That(testOperator.isExcelOpen, Is.True);
        }

        [Category("Slow Running")]
        [Test]
        public void ExcelOperator_CloseExcel_CloseExcelApplication()
        {            
            testOperator.CloseExcel();

            Assert.That(testOperator.isExcelOpen, Is.False);
        }

        [Category("Slow Running")]
        [Test]
        public void ExcelOperator_NewWorkbook_OpensNewWorkbook()
        {
            testOperator.OpenWorkbook();

            Assert.That(testOperator.WorkbookCount, Is.EqualTo(1));
        }

        [TearDown]
        public void CloseExcelOperator()
        {
            if (testOperator.isExcelOpen)
	        {
		        testOperator.CloseExcel();
	        }
        }

        private static ExcelOperator getInitializedOperator()
        {
            ExcelOperator testOperator = new ExcelOperator();

            testOperator.InitializeExcel();
            return testOperator;
        }        
    }
}
