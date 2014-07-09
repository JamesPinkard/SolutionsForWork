using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSolution.Code;
using NUnit.Framework;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace ExcelSolution.UnitTests
{

    [TestFixture]
    public class ExcelChartAppTests
    {
        private ExcelChartApp chartApplication = new ExcelChartApp();

        [SetUp]
        public void OpenExcelApp()
        {
            ExcelChartApp chartApplication = new ExcelChartApp();
        }

        [Test]
        public void InitializeExcel_SampleData_ExcelIsNotOpen()
        {
            chartApplication.InitializeExcel();
            Assert.IsTrue(chartApplication.IsOpen);
        }

        [Test]
        public void CloseExcel_AppHasBeenInitialized_AppIsNull()
        {

            chartApplication.InitializeExcel();
            chartApplication.CloseExcel();
            Assert.IsFalse(chartApplication.IsOpen);
        }
        
        [TearDown]
        public void CloseExcelApp()
        {
            chartApplication.CloseExcel();
        }
    }
}
