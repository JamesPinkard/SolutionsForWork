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
    
    [Category("Workbook")]
    [TestFixture]
    public class WorkbookTests
    {
        ExcelOperator testOperator;
        Workbook testWorkbook;

        #region Fixture Setup and Teardown
        [TestFixtureSetUp]
        public void InitializeExcel()
        {
            testOperator = new ExcelOperator();
            testOperator.InitializeExcel();
            testWorkbook = initTestWorkbook();
        }

        [TestFixtureTearDown]
        public void CloseExcel()
        {
            testOperator.CloseWorkbook();
            testOperator.CloseExcel();
        } 
        #endregion

        [Test]
        public void Workbook_AddSheet_AddsNewSheet()
        {          
            testWorkbook.AddSheet();

            Assert.That(testWorkbook.sheetCount, Is.EqualTo(2));
        }

        [Test]
        public void Workbook_RemoveSheet_RemovesSheetAtEnd()
        {
            testWorkbook.AddSheet();
            testWorkbook.RemoveSheet();

            Assert.That(testWorkbook.sheetCount, Is.EqualTo(2));
        }

        [Test]
        public void Workbook_SelectSheet_SetsActiveSheet()
        {
            testWorkbook.SelectSheet(1);

            Assert.That(testWorkbook.ActiveSheet, Is.SameAs(testWorkbook.Sheets[1]));
        }

        private Workbook initTestWorkbook()
        {
            Workbook testWorkbook = testOperator.OpenWorkbook();
            return testWorkbook;
        }
    }
}
