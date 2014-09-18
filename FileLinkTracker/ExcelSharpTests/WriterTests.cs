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
        private Sheet linkSheet;
        private DirectoryInfo myTestDirectory;
        private int fileCount;
        private DirectoryInfo[] subDirectories;
        private IEnumerable<string> subFiles;
        private List<string> subFilesInSheet;
        private string subDirectoryInSheet;

        private void SetupLinkSheet()
        {
            linkSheet = (Sheet)testFactory.ExecuteMake();
        }
        private void SetupTestFactory()
        {
            // Current Directory is C:\Users\jpinkard\Documents\GitHub\SolutionsForWork\FileLinkTracker\ExcelSharpTests\bin\Debug
            testFactory = new LinkSheetFactory(testWorkbook, Directory.GetCurrentDirectory());
        }
        private void SetupLinkSheetAndFactory()
        {
            SetupTestFactory();
            SetupLinkSheet();
        }
        private DirectoryLinkWriter SetupDirectoryWriter()
        {
            SetupLinkSheetAndFactory();
            DirectoryLinkWriter dw = linkSheet.Writer as DirectoryLinkWriter;
            return dw;
        }
        private DirectoryLinkWriter SetupSolutionDirectoryWriter()
        {
            string relativeSolutionDirectory = "../../../.";
            testFactory = new LinkSheetFactory(testWorkbook, relativeSolutionDirectory);
            linkSheet = testFactory.ExecuteMake() as Sheet;
            return linkSheet.Writer as DirectoryLinkWriter;
        }
        
        [Test]
        public void LinkSheetFactory_LinkSheet_SourceIsCurrentDirectory()
        {
            SetupLinkSheetAndFactory();
            testWriter = (LinkWriter)linkSheet.Writer;                              

            Assert.That(testWriter.Source, Is.EqualTo(Directory.GetCurrentDirectory()));
        }

        [Test]
        public void DirectoryLinkWriter_SetDate_IsDateSetIsTrue()
        {
            directoryWriter = SetupSolutionDirectoryWriter();

            directoryWriter.LinkDate = DateTime.Today;

            Assert.That(directoryWriter.IsDateSet, Is.True);
        }

        [Test]
        public void DirectoryLinkWriter_WriteWithDate_SheetHasDirectoryAndSearchDate()
        {
            directoryWriter = SetupDirectoryWriter();
            directoryWriter.LinkDate = DateTime.Today;
            string strWriterDate = directoryWriter.LinkDate.ToString();
            string[,] heading ={ { Directory.GetCurrentDirectory(), strWriterDate} };

            directoryWriter.Write();

            Assert.That(linkSheet.GetRange("A1", "B1"), Is.EqualTo(heading));
        }

        [Test]
        public void DirectoryLinkWriter_WriteWithDate_SheetOnlyHasFilesCreatedOnDate()
        {
            directoryWriter = SetupSolutionDirectoryWriter();
            directoryWriter.LinkDate = new DateTime(2014, 8, 14);
            IEnumerable<string> fileNames = getFileNamesCreatedOnDate(directoryWriter.DirectoryPath);

            directoryWriter.Write();
            List<string> filesInSheet = getFilesInSheet(fileNames, 2);

            Assert.That(filesInSheet, Is.EquivalentTo(fileNames));
        }
            private IEnumerable<string> getFileNamesCreatedOnDate(string directoryPath)
            {

                FileInfo[] solutionFiles = getFilesInSolutionDirectory(directoryPath);
                IEnumerable<string> fileNames = from f in solutionFiles
                                                where f.CreationTime.Date == directoryWriter.LinkDate.Date
                                                select f.Name;
                return fileNames;
            }
        
        [Test]
        public void DirectoryLinkWriter_WriteWithNullDate_SheetHasFilesInDirectory()
        {            
            directoryWriter = SetupSolutionDirectoryWriter();
            IEnumerable<string> fileNames = getFileNames(directoryWriter.DirectoryPath);          
                       
            directoryWriter.Write();
            List<string> filesInSheet = getFilesInSheet(fileNames, 2);

            Assert.That(filesInSheet, Is.EquivalentTo(fileNames));
            Console.WriteLine(directoryWriter.DirectoryPath);

        }
        
            private IEnumerable<string> getFileNames(string directoryPath)
            {
                FileInfo[] solutionFiles = getFilesInSolutionDirectory(directoryPath);
                IEnumerable<string> fileNames = from f in solutionFiles
                                                select f.Name;
                return fileNames;
            }
            private static FileInfo[] getFilesInSolutionDirectory(string directoryPath)
            {
                DirectoryInfo solutionDirectory = new DirectoryInfo(directoryPath);
                FileInfo[] solutionFiles = solutionDirectory.GetFiles();
                return solutionFiles;
            }
            private List<string> getFilesInSheet(IEnumerable<string> fileNames, int startRow)
            {
                List<string> filesInSheet = new List<String>();
                filesInSheet = linkSheet.GetColumnRange("B", startRow, fileNames.Count<string>() + 1);
                return filesInSheet;
            }

        [Test]
        public void DirectoryLinkWriter_Write_SheetHasFirstSubdirectoryAfterFileList()
        {
            directoryWriter = SetupSolutionDirectoryWriter();
            SetupSubDirectoryTest();

            // Execute
            directoryWriter.Write();

            GetSubDirectoryAndSubFilesInSheet();
            

            Assert.That(subFilesInSheet, Is.EquivalentTo(subFiles));
            Assert.That(subDirectoryInSheet, Is.EqualTo(subDirectories[0].FullName));
            Console.WriteLine(directoryWriter.DirectoryPath);
        }

            private void GetSubDirectoryAndSubFilesInSheet()
        {
            // Get Files in sheet
            subFilesInSheet = getFilesInSheet(subFiles, fileCount + 2);
            subDirectoryInSheet = linkSheet.GetCell(fileCount + 1, 0);
        }

            private void SetupSubDirectoryTest()
            {
                // Get File Count
                fileCount = getDirectoryFileCount();

                // Get Sub Directories
                myTestDirectory = getSolutionDirectory();
                subDirectories = myTestDirectory.GetDirectories();
                subFiles = getFileNames(subDirectories[0].FullName);
            }
               
                private int getDirectoryFileCount()
            {            
                IEnumerable<string> fileNames = getFileNames(directoryWriter.DirectoryPath);
                int fileCount = fileNames.Count<string>();
                return fileCount;
            }
                private DirectoryInfo getSolutionDirectory()
                {
                    DirectoryInfo solutionDirectory = new DirectoryInfo(directoryWriter.DirectoryPath);
                    return solutionDirectory;
                }

        [Test]
        public void DirectoryLinkWriter_Write_SheetDirectoryIsHyperlink()
        {
            directoryWriter = SetupSolutionDirectoryWriter();
            myTestDirectory = getSolutionDirectory();

            directoryWriter.Write();
            string pathToHyperlink = directoryWriter.GetHyperlinkPath("A1");

            Assert.That(pathToHyperlink, Is.EqualTo(myTestDirectory.FullName));            
        }
        
        // TODO Test that sheet has sub directory hyperlinks
        // TODO Error if link writer is set twice
        // Message: Use a sheet factory or change sheet command
        //  to initialize sheet properly.
    }
}
