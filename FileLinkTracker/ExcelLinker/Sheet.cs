using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSharp
{
    public class Sheet
    {
        public int Index { get { return worksheet.Index; } }
        
        public Sheet(Excel._Worksheet worksheet )
        {
            this.worksheet = worksheet;
        }
                
        private Excel._Worksheet worksheet;        
    }
}
