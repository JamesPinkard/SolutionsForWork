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
        public string Source { get; set; }
        public Microsoft.Office.Interop.Excel._Worksheet worksheet
        {
            get;
            set;
        }

        public void Write()
        {
            subWrite();
        }

        protected abstract void subWrite();


    }
}
