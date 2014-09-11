using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelSharp
{
    public class DirectoryLinkWriter : LinkWriter
    {
        public string DirectoryPath { get; private set; }
        private DirectoryInfo linkSourceDirectory;
                
        public DirectoryLinkWriter(string directoryPath)
        {
            linkSourceDirectory = new DirectoryInfo(directoryPath);
            DirectoryPath = directoryPath;
        }
    }
}
