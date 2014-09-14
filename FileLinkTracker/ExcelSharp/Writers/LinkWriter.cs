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

        public void Write()
        {
            subWrite();
        }

        protected abstract void subWrite();

        public Microsoft.Office.Interop.Excel._Worksheet worksheet
        {
            get { throw new NotImplementedException(); }
        }
    }
}
