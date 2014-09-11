using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class LinkSheetFactory : SheetFactory
    {
        public string DirectoryPath { get; set; }
        public LinkSheetFactory(Workbook workbook, string directoryPath)
            : base(workbook)
        {
            DirectoryPath = directoryPath;
        }

        protected override IOfficeWriter MakeSheetWriter()
        {
            DirectoryLinkWriter linkWriter = new DirectoryLinkWriter(DirectoryPath);
            return linkWriter;
        }

        protected override IEmbedder MakeSheetTools()
        {
            LinkEmbedder linkTool = new LinkEmbedder();
            return linkTool;
        }
    }
}
