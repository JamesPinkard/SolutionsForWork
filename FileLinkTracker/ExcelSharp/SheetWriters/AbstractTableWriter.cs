using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    abstract public class AbstractTableWriter : IOfficeWriter
    {

        // From Null object example in Agile Principles
        public static readonly AbstractTableWriter NULL = new NullSheetWriter();

        private class NullSheetWriter : AbstractTableWriter
        {

            public override void WriteFields()
            {
                
            }

            public override void WriteBody()
            {
                
            }
        }

        virtual public void WriteFields()
        {
            // TODO Implement Template Method
        }

        virtual public void WriteBody()
        {
            // TODO Implement Template Method
        }
    }
}
