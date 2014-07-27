using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using Reader.ProCode;
using NUnit.Framework;
using Moq;


namespace Reader.UnitTests
{
    [TestFixture]
    public class ExcelFielderTests
    {
        [Test]
        public void ExcelFielder_ReadExcelFileWithFieldsAtTopRow_ReturnTopRow()
        {
            var mockFieldManager = new Mock<IFieldManager>();
            var fielder = new ExcelFielder(mockFieldManager.Object);
            fielder.FileSource = @"C:\Users\jpinkard\Documents\GitHub\SolutionsForWork\AcumiChart\TestFiles\Excel_FieldInTopRowSample.xlsx";
            fielder.WorksheetSource = "NoDupFields";
            var excel = new ExcelQueryFactory(@"C:\Users\jpinkard\Documents\GitHub\SolutionsForWork\AcumiChart\TestFiles\Excel_FieldInTopRowSample.xlsx");
            var fields = excel.GetColumnNames("NoDupFields");

            
            fielder.Read();

            Assert.That(mockFieldManager.AvailableFields, Is.EqualTo(fields));
        }
    }
}
