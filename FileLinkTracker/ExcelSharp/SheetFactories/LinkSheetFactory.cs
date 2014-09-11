using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class LinkSheetFactory : SheetFactory
    {
        public LinkSheetFactory(Workbook workbook) : base (workbook)
        {

        }

        protected override IOfficeWriter MakeSheetWriter()
        {
            LinkWriter linkWriter = new LinkWriter();
            return linkWriter;
        }

        protected override IEmbedder MakeSheetTools()
        {
            LinkEmbedder linkTool = new LinkEmbedder();
            return linkTool;
        }
    }
}
