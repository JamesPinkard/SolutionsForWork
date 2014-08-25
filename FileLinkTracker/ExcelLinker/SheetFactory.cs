using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class SheetFactory : ISheetCommand
    {
        private Workbook workbook;
        public Sheet SourceSheet { get; set; }

        public SheetFactory(Workbook workbook)
        {
            this.workbook = workbook;
        }

        public void ExecuteWithWorkbook()
        {
            SheetWriter sheetWriter = MakeSheetWriter();
            // rework Workbook design before proceding
            workbook.AddSheet();
            Sheet newSheet = workbook.RecentlyAddedSheet;
            newSheet.Writer = sheetWriter;
        }        
        protected abstract SheetWriter MakeSheetWriter();
        protected abstract SheetTools MakeSheetTools();

        public void ExecuteWithSheet()
        {
            // TODO Implement with template methods
        }
    }
}
