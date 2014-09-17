using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class HyperlinksByDateCommand : WorkbookCommand
    {
        private DateTime linkDate;

        public DateTime LinkDate
        {
            get { return linkDate; }            
        }
        
        public HyperlinksByDateCommand(Workbook workbook, DateTime linkDate) : base(workbook)
        {
            this.linkDate = linkDate;
        }
        public override void Execute()
        {
            throw new NotImplementedException();
        }
    }
}
