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
    public class SheetTests
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

        private Workbook initTestWorkbook()
        {
            Workbook testWorkbook = testOperator.OpenWorkbook();
            return testWorkbook;
        }
        #endregion

        /*
        [Test]
        public void Sheet_GetCells_ReturnsArrayOfValuesInCells()
        {
            
        }
        */
        /*
        [Test]
        public void Sheet_GetRange_ReturnsArrayOfValuesInRange()
        {

        }
        */
        
        /*
        [Test]
        public void Sheet_WriteHeaders_FirstRowContainsHeaders()
        {
            //TODO
        }
        */
    }
}
