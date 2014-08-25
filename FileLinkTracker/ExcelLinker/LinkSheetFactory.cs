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

        protected override SheetWriter MakeSheetWriter()
        {
            throw new NotImplementedException();
        }

        protected override SheetTools MakeSheetTools()
        {
            throw new NotImplementedException();
        }
    }
}
