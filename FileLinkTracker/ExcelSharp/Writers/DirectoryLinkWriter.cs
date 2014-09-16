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
        private DateTime date;

        public DateTime Date
        {
            get { return date; }
            set
            {
                date = value;
                base.Date = value;
            }
        }
        
        private DirectoryInfo linkSourceDirectory;
                
        public DirectoryLinkWriter(string directoryPath)
        {
            linkSourceDirectory = new DirectoryInfo(directoryPath);
            DirectoryPath = directoryPath;
            base.Source = directoryPath;
        }

        protected override void subWrite()
        {
            Excel.Range header = worksheet.get_Range("A1:B1");
            header.Cells[1] = DirectoryPath;            
            header.Cells[2] = date;

            

        }
    }
}
