using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class TableSheetFactory : SheetFactory
    {           
        public TableSheetFactory(Workbook workbook) : base(workbook) { }


        protected override IOfficeWriter MakeSheetWriter()
        {
            TableWriter tableWriter = new TableWriter();
            return tableWriter;
        }

        protected override IEmbedder MakeSheetTools()
        {
            TableEmbedder tableTools = new TableEmbedder();
            return tableTools;
        }
    }
}
