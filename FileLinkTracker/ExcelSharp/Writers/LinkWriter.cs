using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public abstract class LinkWriter : IOfficeWriter
    {
        public DateTime Date { get; protected set; }

        public void WriteFields()
        {
            throw new NotImplementedException();
        }

        public void WriteBody()
        {
            throw new NotImplementedException();
        }
    }
}
