using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    public class Workbook
    {
        // TODO
        // Active Sheet should be a Sheet Class
        public Sheet ActiveSheet { get; private set; }
        public List<Sheet> Sheets
        {
            get { return sheets; }
        }
        public int sheetCount
        {
            get { return oWB.Sheets.Count; }
        }
        public Sheet RecentlyAddedSheet { get; private set; }

        private int sheetCounter;
        private Dictionary<Sheet, Excel._Worksheet> sheetDict = new Dictionary<Sheet,Excel._Worksheet>();        
        private List<Sheet> sheets = new List<Sheet>();
        private Excel._Workbook oWB;
        private Excel.Sheets workbookSheets;        
        
        public Workbook( Excel._Workbook oWB)
        {
            this.oWB = oWB;            
            this.workbookSheets = oWB.Sheets;
            foreach (Excel._Worksheet openedWorkSheet in workbookSheets)
            {
                sheetCounter += 1;
                Sheet openedSheet = new Sheet(openedWorkSheet);
                setIndex(openedSheet);
                Sheets.Add(openedSheet);
                sheetDict[openedSheet] = openedWorkSheet;                
            }
            this.ActiveSheet = Sheets[0];
        }            
        
        // TODO
        // Assign wbsheet to Sheet Class
        public void AddSheet()
        {
            sheetCounter += 1;
            workbookSheets.Add();
            Excel._Worksheet ws = workbookSheets[sheetCount - 1];
            Sheet addedSheet = new Sheet(ws);
            
            setIndex(addedSheet);
            sheets.Add(addedSheet);
            sheetDict[addedSheet] = ws;            
            RecentlyAddedSheet = addedSheet;
        }

        private void setIndex(Sheet addedSheet)
        {
            addedSheet.Index = sheetCounter;
        }

        // TODO
        // Remove Sheet and Makes Sheet Null
        public void RemoveSheet()
        {
            sheetCounter -= 1;            
            Sheet removedSheet = Sheets[sheetCount - 1];
            Excel._Worksheet removedWS = this.oWB.Sheets[sheetCount - 1];

            removedWS.Delete();
            Sheets.Remove(removedSheet);
            sheetDict.Remove(removedSheet);            
        }

        // TODO
        // Selects class Sheet as well as oWB Sheet
        public void SelectSheet(int sheetsIndex)
        {
            this.ActiveSheet = this.Sheets[sheetsIndex - 1];
            Excel._Worksheet activeSheet = sheetDict[ActiveSheet];
            activeSheet.Activate();
        }

        // TODO
        // Needs Re-thinking, the Sheet should be Null default
        // Error Should arise if it doesn't exist
        public Sheet GetSheet(int sheetIndex)
        {
            Sheet returnSheet = Sheets[sheetIndex - 1];
            return returnSheet;
        }
    }
}
