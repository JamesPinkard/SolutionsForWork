using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;

namespace ExcelSharpTests
{
    [Category("Workbook Tests")]
    [TestFixture]
    class WorkbookTests
    {
        ExcelOperator testOperator;
        Workbook testWorkbook;

        #region Fixture Setup and Teardown
        [TestFixtureSetUp]
        public void InitializeWorkbook()
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

            Assert.That(testWorkbook.ActiveSheet, Is.SameAs(testWorkbook.Sheets[0]));
        }
        
        [Test]
        public void Workbook_GetSheet_ReturnsSheet()
        {
            Sheet testSheet = testWorkbook.GetSheet(1);
            Assert.That(testSheet.Index, Is.EqualTo(1));
        }

        private Workbook initTestWorkbook()
        {
            string path = Environment.CurrentDirectory;
            string wbPath = @"{0}\TestingFiles\TestingWorkbook.xlsx";
            Workbook testWorkbook = testOperator.OpenWorkbook(string.Format(wbPath, path));
            return testWorkbook;
        }
    }
}
