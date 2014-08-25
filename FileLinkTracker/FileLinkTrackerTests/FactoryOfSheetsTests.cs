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
    [Category("Factory Sheet Tests")]
    [TestFixture]
    public class FactoryOfSheetsTests
    {
        ExcelOperator testOperator;
        Workbook testWorkbook;
        TestHelper helper;

        #region Setup and Teardown
        [TestFixtureSetUp]
        public void InitializeWorksheet()
        {
            helper.InitializeExcel(testOperator, testWorkbook);
        }

        [TestFixtureTearDown]
        public void CloseExcel()
        {
            helper.CloseExcel(testOperator);
        }

        [SetUp]
        public void ClearCells()
        {
            helper.ClearCells(testWorkbook);
        }
        #endregion        

        [Category("Integration Test")]
        [Test]
        public void FactoryOfTableSheets_IntegrationTest()
        {
            ISheetCommand tableSheetFactory = new TableSheetFactory(testWorkbook);            
            int initialSheetCount = testWorkbook.sheetCount;

            // Should Not Return Anything; Should just add sheets to workbook
            tableSheetFactory.ExecuteWithWorkbook(); //First Test
            Sheet tableSheet = testWorkbook.RecentlyAddedSheet;
            Assert.That(testWorkbook.sheetCount, Is.EqualTo(initialSheetCount + 1));            
            Assert.That(tableSheet.Writer, Is.TypeOf(typeof(TableSheetWriter)));
            Assert.That(tableSheet.Tools, Is.TypeOf(typeof(TableSheetTools)));
            
            SheetFactory linkSheetFactory = new LinkSheetFactory(testWorkbook);
            linkSheetFactory.SourceSheet = tableSheet;
            linkSheetFactory.ExecuteWithSheet();
            Sheet linkSheet = testWorkbook.RecentlyAddedSheet;

            Assert.That(testWorkbook.sheetCount, Is.EqualTo(initialSheetCount + 2));
            Assert.That(linkSheet.Writer, Is.TypeOf(typeof(LinkSheetWriter)));
            Assert.That(linkSheet.Tools, Is.TypeOf(typeof(LinkSheetTools)));
            Assert.That(linkSheet.wbPath, Is.EqualTo(tableSheet.wbPath));
        }
        
        [Category("Sheet Factory")]
        [Test]
        public void TableSheetFactory_ExecuteWithWorkbook_MakeNewTableSheet()
        {
            TableSheetFactory testFactory;
            testFactory = getTestFactory();

            testFactory.ExecuteWithWorkbook();

            Assert.That(testWorkbook.RecentlyAddedSheet.Writer,Is.InstanceOf(typeof(TableSheetWriter)));
        }

        private TableSheetFactory getTestFactory()
        {
            TableSheetFactory testFactory;
            testFactory = new TableSheetFactory(testWorkbook);
            return testFactory;
        }
        
        // Need to implement SheetHandler Class
        //      Needs to overwrite sheets, save, and open
    }
}
