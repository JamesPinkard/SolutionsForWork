using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public abstract class LinkWriter : IOfficeWriter
    {
        private DateTime linkDate;
        private bool dateSet = false;

        public bool IsDateSet
        {
            get { return dateSet; }            
        }
        
        public DateTime LinkDate 
        {
            get { return linkDate; }
            set { linkDate = value;
                dateSet = true; }
        }

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
