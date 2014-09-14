using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    abstract public class NullableTableWriter : IOfficeWriter
    {

        // From Null object example in Agile Principles
        public static readonly NullableTableWriter NULL = new NullSheetWriter();

        private class NullSheetWriter : NullableTableWriter
        {

            public override void Write()
            {
                
            }

        }


        
        public virtual void Write()
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Interop.Excel._Worksheet worksheet
        {
            get { throw new NotImplementedException(); }
        }
    }
}
