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
            LinkSheetWriter linkWriter = new LinkSheetWriter();
            return linkWriter;
        }

        protected override IOfficeTool MakeSheetTools()
        {
            LinkSheetTool linkTool = new LinkSheetTool();
            return linkTool;
        }
    }
}
