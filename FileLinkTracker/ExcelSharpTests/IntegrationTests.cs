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
            IOfficeMaker tableSheetFactory = new TableSheetFactory(testWorkbook);
            int initialSheetCount = testWorkbook.sheetCount;

            // Should Not Return Anything; Should just add sheets to workbook
            tableSheetFactory.ExecuteMake(); //First Test
            Sheet tableSheet = testWorkbook.RecentlyAddedSheet;
            
            Assert.That(testWorkbook.sheetCount, Is.EqualTo(initialSheetCount + 1));
            Assert.That(tableSheet.Writer, Is.TypeOf(typeof(AqtestTableWriter)));
            Assert.That(tableSheet.EmbedTool, Is.TypeOf(typeof(TableEmbedder)));

            SheetFactory linkSheetFactory = new LinkSheetFactory(testWorkbook, Directory.GetCurrentDirectory());            
            linkSheetFactory.ExecuteCopy();
            Sheet linkSheet = testWorkbook.RecentlyAddedSheet;

            Assert.That(testWorkbook.sheetCount, Is.EqualTo(initialSheetCount + 2));
            Assert.That(linkSheet.Writer, Is.TypeOf(typeof(DirectoryLinkWriter)));
            Assert.That(linkSheet.EmbedTool, Is.TypeOf(typeof(LinkEmbedder)));
            Assert.That(linkSheet.wbPath, Is.EqualTo(tableSheet.wbPath));
        }        

        [Test]
        public void ChangeSheetCommands_IntegrationTest()
        {
            // Set up Sheet
            IOfficeMaker tableSheetFactory = new TableSheetFactory(testWorkbook);
            IOfficeCommandable sourceSheet = tableSheetFactory.ExecuteMake();
            
            // Change Embed Tool Test
            OfficeCommand changeToLinkTool = new ChangeToLinkEmbedder(sourceSheet);
            changeToLinkTool.Execute();

            Assert.That(sourceSheet.isValid(), Is.True);
            Assert.That(sourceSheet.EmbedTool, Is.InstanceOf<LinkEmbedder>());
            
            // Change Writer Test
            OfficeCommand changeToLinkSheetWriter = new ChangeToLinkWriter(sourceSheet, Directory.GetCurrentDirectory());
            changeToLinkSheetWriter.Execute();

            Assert.That(sourceSheet.Writer, Is.InstanceOf<DirectoryLinkWriter>());

            // Change Formatter Test
            OfficeCommand changeToDirectoryFormatter = new ChangeToDirectoryFormatter(sourceSheet);
            changeToDirectoryFormatter.Execute();

            Assert.That(sourceSheet.Formatter, Is.InstanceOf<DirectoryFormatter>());

            // Change Remover Tool Test
            OfficeCommand changeToLinkRemover = new ChangeToLinkRemover(sourceSheet);
            changeToLinkRemover.Execute();

            Assert.That(sourceSheet.RemoveTool, Is.InstanceOf<LinkRemover>());
        }

        [Test]
        public void GenerateLinksCommand_IntegrationTest()
        {
            // Set up Sheet with Current Directory           
            LinkSheetFactory linkSheetFactory = new LinkSheetFactory(testWorkbook, Directory.GetCurrentDirectory());
            Sheet CurrentDirectorySheet = linkSheetFactory.ExecuteMake() as Sheet;

            // Set up Sheet with Solution Directory
            string solutionRelativePath = "../../../.";
            linkSheetFactory.DirectoryPath = solutionRelativePath;
            Sheet SolutionDirectorySheet = linkSheetFactory.ExecuteMake() as Sheet;

            DateTime testDate = new DateTime(2014, 5, 21);

            WorkbookCommand generateWorkbookLinks = new HyperlinksByDateCommand(testWorkbook, testDate);
            generateWorkbookLinks.Execute();

            WriterAssertions(CurrentDirectorySheet, testDate, Directory.GetCurrentDirectory());
            WriterAssertions(SolutionDirectorySheet, testDate, solutionRelativePath);
            
        }

        private static void WriterAssertions(Sheet testSheet, DateTime testDate, string pathToDirectory)
        {
            DirectoryLinkWriter sourceWriter = testSheet.Writer as DirectoryLinkWriter;
            Assert.That(testSheet.Writer, Is.InstanceOf<LinkWriter>());
            Assert.That(sourceWriter.LinkDate, Is.EqualTo(testDate));
            Assert.That(sourceWriter.DirectoryPath, Is.EqualTo(pathToDirectory));
        }
    }
}
