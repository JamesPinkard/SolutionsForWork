using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    public class Workbook
    {
        public Excel._Worksheet ActiveSheet { get; private set; }
        private Excel._Workbook oWB;
        public Excel.Sheets Sheets
        {
            get { return oWB.Sheets; }
        }
        public int sheetCount
        {
            get { return oWB.Sheets.Count; }
        }

        public Workbook( Excel._Workbook oWB)
        {
            this.oWB = oWB;
            this.ActiveSheet = (Excel._Worksheet)oWB.ActiveSheet;
        }        

        public void AddSheet()
        {
            oWB.Sheets.Add();
        }

        public void RemoveSheet()
        {
            int endSheetIndex = sheetCount;
            ActiveSheet = this.oWB.Sheets[sheetCount - 1];
            ActiveSheet.Delete();
        }

        public void SelectSheet(int sheetsIndex)
        {
            this.ActiveSheet = oWB.Sheets[sheetsIndex];
        }



        public Sheet GetSheet(int sheetIndex)
        {
            Sheet wbSheet = new Sheet(oWB.Sheets[sheetIndex]);
            return wbSheet;
        }
    }
}
