using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class OfficeCommand : ICommand
    {
        protected IOfficeCommandable Receiver
        {
            get { return commandReceiver; }
            set { commandReceiver = value; }
        }
        private IOfficeCommandable commandReceiver;

        // Constructor
        public OfficeCommand(IOfficeCommandable receiver)
        {
            commandReceiver = receiver;
        }

        public abstract void Execute();
    }
}
