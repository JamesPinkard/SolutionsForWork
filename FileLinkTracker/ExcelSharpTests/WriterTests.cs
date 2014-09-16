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
        public void DirectoryLinkWriter_GetSubdirectories_ReturnSubdirectories()
        {
            SetupLinkSheetAndFactory();
            directoryWriter = (DirectoryLinkWriter)testSheet.Writer;
            string[] heading = { Directory.GetCurrentDirectory(), directoryWriter.Date.ToString() };

            directoryWriter.Write();

            Assert.That(testSheet.GetCells(1, 2), Is.EqualTo(heading));
        }
    }
}
