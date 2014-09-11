using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class OfficeCommand : ICommand
    {
        private IOfficeCommandable commandReceiver;
        protected IOfficeCommandable Receiver
        {
            get { return commandReceiver; }
            set { commandReceiver = value; }
        }

        // Constructor
        public OfficeCommand(IOfficeCommandable receiver)
        {
            commandReceiver = receiver;
        }

        public abstract void Execute();
    }
}
