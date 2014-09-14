using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public abstract class TableWriter : IOfficeWriter
    {
        public void Write()
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Interop.Excel._Worksheet worksheet
        {
            get { throw new NotImplementedException(); }
        }
    }
}
