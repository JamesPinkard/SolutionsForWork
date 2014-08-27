using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    public class Workbook
    {
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

        public void AddSheet()
        {
            sheetCounter += 1;
            workbookSheets.Add();
            
            Excel._Worksheet ws = getEndSheet();
            initializeNewSheet(ws);
        }
        public void RemoveSheet()
        {
            sheetCounter -= 1;            
            Sheet removedSheet = Sheets[sheetCount - 1];
            Excel._Worksheet removedWS = this.oWB.Sheets[sheetCount - 1];

            removedWS.Delete();
            Sheets.Remove(removedSheet);
            sheetDict.Remove(removedSheet);            
        }
        public Sheet GetSheet(int sheetIndex)
        {
            Sheet returnSheet = Sheets[sheetIndex - 1];
            return returnSheet;
        }
        public void SelectSheet(int sheetsIndex)
        {
            ActiveSheet = Sheets[sheetsIndex - 1];
            Excel._Worksheet activeSheet = sheetDict[ActiveSheet];
            activeSheet.Activate();
        }
        public void CopySheet(Sheet originalSheet)
        {
            Excel._Worksheet oSheet = sheetDict[originalSheet];
            Excel._Worksheet endSheet = getEndSheet();                        
            sheetCounter += 1;

            // Copy Sheet to the last position in the workbook
            oSheet.Copy(Type.Missing,endSheet);

            Excel._Worksheet copySheet = getEndSheet();
            initializeNewSheet(copySheet);
        }

        private Excel._Worksheet getEndSheet()
        {
            Excel._Worksheet ws = workbookSheets[sheetCount];
            return ws;
        }
        private void initializeNewSheet(Excel._Worksheet ws)
        {
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
    }
}
        
                
                
                


