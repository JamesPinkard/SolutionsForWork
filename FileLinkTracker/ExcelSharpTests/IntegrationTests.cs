using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;
using System.IO;

namespace ExcelSharpTests
{
    [Category("Integration Test")]
    [TestFixture]
    public class IntegrationTests
    {
        ExcelOperator testOperator;
        Workbook testWorkbook;

        #region Setup and Teardown

        [TestFixtureSetUp]
        public void InitializeWorksheet()
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


        private Workbook initTestWorkbook()
        {
            string path = Environment.CurrentDirectory;
            string wbPath = @"{0}\TestingFiles\TestingWorkbook.xlsx";
            Workbook testWorkbook = testOperator.OpenWorkbook(string.Format(wbPath, path));
            return testWorkbook;
        }
        #endregion

        [Test]
        public void FactoryOfTableSheets_IntegrationTest()
        {
            IMaker tableSheetFactory = new TableSheetFactory(testWorkbook);
            int initialSheetCount = testWorkbook.sheetCount;

            // Should Not Return Anything; Should just add sheets to workbook
            tableSheetFactory.ExecuteMake(); //First Test
            Sheet tableSheet = testWorkbook.RecentlyAddedSheet;
            Assert.That(testWorkbook.sheetCount, Is.EqualTo(initialSheetCount + 1));
            Assert.That(tableSheet.Writer, Is.TypeOf(typeof(TableSheetWriter)));
            Assert.That(tableSheet.Tools, Is.TypeOf(typeof(TableSheetTool)));

            SheetFactory linkSheetFactory = new LinkSheetFactory(testWorkbook);            
            linkSheetFactory.ExecuteCopy();
            Sheet linkSheet = testWorkbook.RecentlyAddedSheet;

            Assert.That(testWorkbook.sheetCount, Is.EqualTo(initialSheetCount + 2));
            Assert.That(linkSheet.Writer, Is.TypeOf(typeof(LinkSheetWriter)));
            Assert.That(linkSheet.Tools, Is.TypeOf(typeof(LinkSheetTool)));
            Assert.That(linkSheet.wbPath, Is.EqualTo(tableSheet.wbPath));
        }        
    }
}
