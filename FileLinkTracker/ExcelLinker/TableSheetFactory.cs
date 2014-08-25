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

        protected override SheetTools MakeSheetTools()
        {
            throw new NotImplementedException();
        }
    }
}
