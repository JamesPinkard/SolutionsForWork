using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NUnit.Framework;
using ExcelSharp;

namespace ExcelSharpTests
{
    [Category("WriterTests")]
    [TestFixture]
    public class WriterTests : BaseTest
    {
        private LinkWriter testWriter;
        private DirectoryLinkWriter directoryWriter;

        private void SetupLinkSheet()
        {
            testSheet = (Sheet)testFactory.ExecuteMake();
        }
        private void SetupTestFactory()
        {
            testFactory = new LinkSheetFactory(testWorkbook, Directory.GetCurrentDirectory());
        }
        private void SetupLinkSheetAndFactory()
        {
            SetupTestFactory();
            SetupLinkSheet();
        }
                
        [Test]
        public void LinkWriter_RecursiveWrite_SourceIsCurrentDirectory()
        {
            SetupLinkSheetAndFactory();
            testWriter = (LinkWriter)testSheet.Writer;                              

            Assert.That(testWriter.Source, Is.EqualTo(Directory.GetCurrentDirectory()));
        }

        [Test]
        public void DirectoryLinkWriter_Write_SheetHasDirectoryAndSearchDate()
        {
            SetupLinkSheetAndFactory();
            directoryWriter = (DirectoryLinkWriter)testSheet.Writer;
            directoryWriter.Date = DateTime.Today;
            string strWriterDate = directoryWriter.Date.ToString();
            string[,] heading ={ { Directory.GetCurrentDirectory(), strWriterDate} };

            directoryWriter.Write();

            Assert.That(testSheet.GetRange("A1","B1"), Is.EqualTo(heading));
        }

        // TODO Error if link writer is set twice
        // Message: Use a sheet factory or change sheet command
        //  to initialize sheet properly.
    }
}
