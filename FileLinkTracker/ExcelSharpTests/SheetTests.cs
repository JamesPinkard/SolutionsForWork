using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ExcelSharp;


namespace ExcelSharpTests
{
    [Category("Sheet Tests")]
    [TestFixture]
    public class SheetTests
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
        public void Sheet_GetCells_ReturnsArrayOfValuesBlockOfCells()
        {
            setupActiveSheetArray();
            string[,] headerArray = setupHeaderArray();
            string[,] testArray = setupTestArray( 1, 3);
            Sheet testSheet = setupTestSheet();                       
            
            testArray = testSheet.GetCells(1, 3);

            Assert.That(testSheet.Name, Is.EqualTo("TestingSheet"));
            Assert.That(testArray, Is.EqualTo(headerArray));
        }

        [Test]
        public void Sheet_GetRange_ReturnsArrayOfValuesInRange()
        {
            setupActiveSheetArray();
            string[,] namesArray = setupNamesArray();
            string[,] testArray = setupTestArray(4, 3);
            Sheet testSheet = setupTestSheet();

            testArray = testSheet.GetRange("A1", "C4");

            Assert.That(testArray, Is.EqualTo(namesArray));
        }
        
        [Test]
        public void Sheet_Exists_ReturnTrueIfWorksheetSet()
        {
            Sheet testSheet = setupTestSheet();

            Assert.That(testSheet.Exists, Is.True);
            testWorkbook.AddSheet();
            testSheet = testWorkbook.GetSheet(2);

            testSheet.DeleteSheet();

            Assert.That(testSheet.Exists, Is.False);
        }
        // Check that string can return other values
         
        #region GetCellsTest Helper Methods
        /// <summary>
        /// Setup Worksheet(index 1) with test array values
        /// </summary>
        private void setupActiveSheetArray()
        {
            testWorkbook.SelectSheet(1);            
        }
        
       
        /// <summary>
        /// Set up first row of spreadsheet with header information.
        /// </summary>
        /// <returns></returns>
        private static string[,] setupHeaderArray()
        {
            string[,] headerArray = new string[1, 3];
            headerArray[0, 0] = "First Name";
            headerArray[0, 1] = "Last Name";
            headerArray[0, 2] = "Full Name";
            return headerArray;
        }

        /// <summary>
        /// Initialize new test array with bounds set as [rows, cols]
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="cols"></param>
        /// <returns></returns>
        private static string[,] setupTestArray(int rows, int cols)
        {
            string[,] testArray = new string[rows, cols];
            return testArray;
        }        
        
        /// <summary>
        /// Initialized new test sheet
        /// </summary>
        /// <returns></returns>
        private Sheet setupTestSheet()
        {
            Sheet testSheet = testWorkbook.GetSheet(1);
            return testSheet;
        }
        #endregion       
       
        #region GetRangeTest Helper Methods
        /// <summary>
        /// set up names array as 4 x 3 array
        /// with header and names.
        /// </summary>
        /// <returns></returns>
        private static string[,] setupNamesArray()
        {
            string[,] namesArray = new string[4, 3];
            namesArray[0, 0] = "First Name";
            namesArray[0, 1] = "Last Name";
            namesArray[0, 2] = "Full Name";
            namesArray[1, 0] = "James";
            namesArray[1, 1] = "Pinkard";
            namesArray[1, 2] = "James Pinkard";
            namesArray[2, 0] = "Laura";
            namesArray[2, 1] = "Pinkard";
            namesArray[2, 2] = "Laura Pinkard";
            namesArray[3, 0] = "Victor";
            namesArray[3, 1] = "Paniagua";
            namesArray[3, 2] = "Victor Paniagua";
            return namesArray;
        }          
        #endregion       
    }
}
