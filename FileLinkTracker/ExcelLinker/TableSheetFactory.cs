using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class TableSheetFactory : SheetFactory
    {   
        // TODO Implement Table Sheet Factory
        public TableSheetFactory(Workbook workbook) : base(workbook) { }


        protected override SheetWriter MakeSheetWriter()
        {
            TableSheetWriter tableWriter = new TableSheetWriter();
            return tableWriter;
        }

        protected override SheetTool MakeSheetTools()
        {
            TableSheetTool tableTools = new TableSheetTool();
            return tableTools;
        }
    }
}
