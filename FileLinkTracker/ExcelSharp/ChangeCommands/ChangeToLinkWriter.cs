using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public class ChangeToLinkWriter : ChangeWriterCommand
    {
        public string DirectoryPath { get; set; }
        protected override IOfficeWriter sourceWriter
        {
            get { return new DirectoryLinkWriter(DirectoryPath); }
        }

        public ChangeToLinkWriter(IOfficeCommandable receiver, string directoryPath)
            : base(receiver)
        {
            DirectoryPath = directoryPath;
        }
    }
}
