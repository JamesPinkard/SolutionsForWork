using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;

namespace ExcelSharpTests
{
    [TestFixture]
    public abstract class BaseTest
    {
        protected ExcelOperator testOperator { get; set; }
        protected Workbook testWorkbook { get; set; }
        protected Sheet testSheet { get; set; }        

        [TestFixtureSetUp]
        public void InitializeWorksheet()
        {
            testOperator = new ExcelOperator();
            testOperator.InitializeExcel();
            testWorkbook = initTestWorkbook();
            testSheet = setupTestSheet();
        }

        [TestFixtureTearDown]
        public void CloseExcel()
        {
            testOperator.CloseWorkbook();
            testOperator.CloseExcel();
        }

        public Workbook initTestWorkbook()
        {
            string path = Environment.CurrentDirectory;
            string wbPath = @"{0}\TestingFiles\TestingWorkbook.xlsx";
            Workbook testWorkbook = testOperator.OpenWorkbook(string.Format(wbPath, path));
            return testWorkbook;
        }
        public Sheet setupTestSheet()
        {
            Sheet testSheet = testWorkbook.GetSheet(1);
            return testSheet;
        }
    }
}
