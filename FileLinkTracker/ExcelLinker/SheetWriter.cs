using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    abstract public class SheetWriter
    {
        public abstract void WriteFields();
        public abstract void WriteTable();

        // From Null object example in Agile Principles
        public static readonly SheetWriter NULL = new NullSheetWriter();

        private class NullSheetWriter : SheetWriter
        {

            public override void WriteFields()
            {
                
            }

            public override void WriteTable()
            {
                
            }
        }
    }
}
