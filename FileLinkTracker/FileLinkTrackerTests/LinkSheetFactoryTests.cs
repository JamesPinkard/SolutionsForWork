using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;

namespace ExcelSharpTests
{
    [Category("Link Sheet Tests")]
    [TestFixture]
    public class LinkSheetFactoryTests
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

        [SetUp]
        public void ClearCells()
        {
            testWorkbook.ActiveSheet.Cells.Clear();
        }

        private Workbook initTestWorkbook()
        {
            Workbook testWorkbook = testOperator.OpenWorkbook();
            return testWorkbook;
        }
        #endregion        

        /*
        [Test]
        public void LinkSheetFactory_Get
         */
    }
}
