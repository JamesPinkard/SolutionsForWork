using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class WorkbookHyperlinksCommand : WorkbookCommand
    {
        private DateTime linkDate;

        public DateTime MyProperty
        {
            get { return linkDate; }            
        }
        
        public WorkbookHyperlinksCommand(Workbook workbook, DateTime linkDate) : base(workbook)
        {
            this.linkDate = linkDate;
        }
        public override void Execute()
        {
            throw new NotImplementedException();
        }
    }
}
