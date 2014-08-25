using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public interface ISheetWriter
    {
        void WriteFields();
        void WriteBody();
    }
}
