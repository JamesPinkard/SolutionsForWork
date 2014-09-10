using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class SheetFactory : IOfficeMaker
    {
        public Sheet SourceSheet { get { return workbook.ActiveSheet; } }
        
        private Workbook workbook;
        private IOfficeWriter sheetWriter;
        private IEmbedder sheetTools;

        public SheetFactory(Workbook workbook)
        {
            this.workbook = workbook;
        }
                
        protected abstract IOfficeWriter MakeSheetWriter();
        protected abstract IEmbedder MakeSheetTools();
        
        // IOfficeMaker methods
        public IOfficeCommandable ExecuteMake()
        {
            MakeSheetUtilites();

            Sheet newSheet = MakeNewSheet();
            
            SetSheetUtilities(newSheet);

            return (IOfficeCommandable)workbook.RecentlyAddedSheet;
        }
        public IOfficeCommandable ExecuteCopy()
        {
            MakeSheetUtilites();
            
            Sheet copySheet = CopySheet();

            SetSheetUtilities(copySheet);

            return (IOfficeCommandable)workbook.RecentlyAddedSheet;
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
            newSheet.EmbedTool = sheetTools;
        }
    }
}
