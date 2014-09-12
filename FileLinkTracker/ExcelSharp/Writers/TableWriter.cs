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
    }
}
