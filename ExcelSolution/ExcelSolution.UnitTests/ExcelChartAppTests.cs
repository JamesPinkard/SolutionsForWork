using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSolution.Code;
using NUnit.Framework;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelSolution.UnitTests
{

    [TestFixture]
    public class ExcelChartAppTests
    {
        [Test]
        public void InitializeExcel_SampleData_ExcelIsNotVisible()
        {
            ExcelChartApp chartApplication = new ExcelChartApp();
            chartApplication.InitializeExcel();
            Assert.IsFalse(chartApplication.Visible);
        }

        [Test]
        public void CloseExcel_Default_AppIsNull()
        {
            ExcelChartApp chartApplication = new ExcelChartApp();
            chartApplication.CloseExcel();
            Assert.IsFalse(chartApplication.IsOpen);
        }
    }
    [TestFixture]
    public class WorkBookTests
    {
        [Test]
        public void InitializeBook_DBpath_SheetCountOfThree()
        {

        }
    }
}
