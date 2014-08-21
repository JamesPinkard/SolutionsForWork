using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;

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
            helper.InitializeWorksheet(testOperator, testWorkbook);
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

        /* TODO
        [Test]
        public void FactoryOfTableSheets_IntegrationTest();
        ISheetFactory tableSheetFactory = new TableSheetFactory(testWorkbook);
        Sheet tableSheet = new TableSheet();
        tableSheet = tableSheetFactory.MakeSheet();


        SheetFactory linkSheetFactory = new LinkSheetFactory(testWorkbook);
        linkSheetFactory.Overwrite();
        */
    }
}
