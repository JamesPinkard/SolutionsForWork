using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class TableSheetFactory : SheetFactory
    {           
        public TableSheetFactory(Workbook workbook) : base(workbook) { }


        protected override ISheetWriter MakeSheetWriter()
        {
            TableSheetWriter tableWriter = new TableSheetWriter();
            return tableWriter;
        }

        protected override ISheetTool MakeSheetTools()
        {
            TableSheetTool tableTools = new TableSheetTool();
            return tableTools;
        }
    }
}
