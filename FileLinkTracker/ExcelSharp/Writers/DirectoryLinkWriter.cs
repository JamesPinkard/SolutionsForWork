using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    public class DirectoryLinkWriter : LinkWriter
    {
        public string DirectoryPath { get; private set; }        
        private DirectoryInfo linkSourceDirectory;
        private int rowIndex = 2;
        private Excel.Range fileColumn;
        private Excel.Range directoryColumn;
                
        public DirectoryLinkWriter(string directoryPath)
        {
            linkSourceDirectory = new DirectoryInfo(directoryPath);
            DirectoryPath = directoryPath;
            base.Source = directoryPath;            
        }

        protected override void subWrite()
        {           
            fileColumn = worksheet.Columns[2];
            directoryColumn = worksheet.Columns[1];

            writeHeader();
            writeFileNamesToColumnB(linkSourceDirectory);

            DirectoryInfo[] subDirectories = linkSourceDirectory.GetDirectories();
            foreach (var sub in subDirectories)
            {
                writeDirectoryPathToColumnA(sub);
                writeFileNamesToColumnB(sub);
            }
        }

        // private helpers
        private void writeHeader()
        {
            Excel.Range header = worksheet.get_Range("A1:B1");
            header.Cells[1] = DirectoryPath;
            header.Cells[2] = LinkDate;
        }
        
        // If Null get all files in Directory
        private void writeFileNamesToColumnB(DirectoryInfo sub)
        {
            var subFiles = getFileNames(sub);
            writeFileNames(subFiles);
        }
        private IEnumerable<string> getFileNames(DirectoryInfo directory)
        {
            FileInfo[] linkFiles = directory.GetFiles();
            var fileNames = from f in linkFiles
                            select f.Name;
            return fileNames;
        }
        private void writeFileNames(IEnumerable<string> fileNames)
        {            
            
            foreach (var file in fileNames)
            {
                fileColumn.Cells[rowIndex] = file;
                rowIndex++;
            }
        }

        private void writeDirectoryPathToColumnA(DirectoryInfo sub)
        {
            directoryColumn.Cells[rowIndex] = sub.FullName;
            rowIndex++;
        }




    }
}
