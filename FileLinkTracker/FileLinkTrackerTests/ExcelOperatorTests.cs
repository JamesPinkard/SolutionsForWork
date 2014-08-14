using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelLinker;
using Excel = Microsoft.Office.Interop.Excel;


namespace FileLinkTrackerTests
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

            Assert.That(testOperator.opWorkbooks[1], Is.InstanceOf(typeof(ExcelLinker.Workbook)));
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
    public class Workbook
    {
        ExcelOperator testOperator;

        #region Fixture Setup and Teardown
        [TestFixtureSetUp]
        public void InitializeExcel()
        {
            testOperator = new ExcelOperator();
            testOperator.InitializeExcel();
        }

        [TestFixtureTearDown]
        public void CloseExcel()
        {
            testOperator.CloseExcel();
        } 
        #endregion

        [Test]
        public void Workbook_AddSheet_AddsNewSheet()
        {
            // TODO
        }
    }
}
