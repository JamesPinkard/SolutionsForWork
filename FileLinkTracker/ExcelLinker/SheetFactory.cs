using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class SheetFactory : ISheetCommand
    {
        public Sheet SourceSheet { get; set; }
        
        private Workbook workbook;
        private SheetWriter sheetWriter;
        private SheetTool sheetTools;

        public SheetFactory(Workbook workbook)
        {
            this.workbook = workbook;
        }

        protected abstract SheetWriter MakeSheetWriter();
        protected abstract SheetTool MakeSheetTools();
        
        public void ExecuteWithWorkbook()
        {
            MakeSheetUtilites();

            Sheet newSheet = MakeNewSheet();
            
            SetSheetUtilities(newSheet);
        }
        public void ExecuteWithSheet()
        {
            MakeSheetUtilites();
            
            Sheet copySheet = CopySheet();

            SetSheetUtilities(copySheet);
        }
        
        private void MakeSheetUtilites()
        {

            sheetWriter = MakeSheetWriter();
            sheetTools = MakeSheetTools();
        }
        private Sheet CopySheet()
        {
            workbook.CopySheet(SourceSheet);
            Sheet copySheet = workbook.RecentlyAddedSheet;
            return copySheet;
        }
        private Sheet MakeNewSheet()
        {
            workbook.AddSheet();
            Sheet newSheet = workbook.RecentlyAddedSheet;
            return newSheet;
        }
        private void SetSheetUtilities(Sheet newSheet)
        {

            newSheet.Writer = sheetWriter;
            newSheet.Tools = sheetTools;
        }
    }
}
