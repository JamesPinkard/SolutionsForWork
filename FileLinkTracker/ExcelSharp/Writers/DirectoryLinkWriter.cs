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
                
        public DirectoryLinkWriter(string directoryPath)
        {
            linkSourceDirectory = new DirectoryInfo(directoryPath);
            DirectoryPath = directoryPath;
            base.Source = directoryPath;            
        }
        public string GetHyperlinkPath(string p)
        {
            Excel.Range hlinkRange = worksheet.get_Range(p);
            Excel.Hyperlink hlink = hlinkRange.Hyperlinks[1];
            string path = hlink.Address;
            return path;
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

        #region writeHeader
        private void writeHeader()
        {
            writePathAndDate();

            AddAbsolutePathHyperlink();
        }
        
            private void writePathAndDate()
        {
            Excel.Range header = worksheet.get_Range("A1:B1");
            header.Cells[1] = DirectoryPath;
            header.Cells[2] = LinkDate;
        }        

            private void AddAbsolutePathHyperlink()
        {
            string absoluteDirPath = Path.GetFullPath(DirectoryPath);
            Excel.Range dir = worksheet.get_Range("A1", "A1");
            worksheet.Hyperlinks.Add(dir, absoluteDirPath);
        }
        #endregion

        #region writeFileNamesToColumnB
        // If Null get all files in Directory
        private void writeFileNamesToColumnB(DirectoryInfo sub)
        {
            var subFiles = getFileNames(sub);
            writeFileNames(subFiles);
        }
        
        private IEnumerable<string> getFileNames(DirectoryInfo directory)
        {
            FileInfo[] linkFiles = directory.GetFiles();
            IEnumerable<string> fileNames;

            
            if (IsDateSet == false )
            {
                fileNames = from f in linkFiles
                              select f.Name;                
            }
            else
            {
                fileNames = from f in linkFiles
                            where f.CreationTime.Date == LinkDate.Date
                            select f.Name; 
            }
            
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
        #endregion
        
        private void writeDirectoryPathToColumnA(DirectoryInfo sub)
        {
            directoryColumn.Cells[rowIndex] = sub.FullName;
            rowIndex++;
        }

        private DirectoryInfo linkSourceDirectory;
        private int rowIndex = 2;
        private Excel.Range fileColumn;
        private Excel.Range directoryColumn;
    }
}
