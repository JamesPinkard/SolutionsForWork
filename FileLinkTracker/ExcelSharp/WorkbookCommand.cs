using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public abstract class WorkbookCommand
    {
        private Workbook commandReceiver;
        protected Workbook Receiver
        {
            get { return commandReceiver; }
            set { commandReceiver = value; }
        }

        // Constructor
        public WorkbookCommand(Workbook receiver)
        {
            commandReceiver = receiver;
        }

        public abstract void Execute();
    }
}
