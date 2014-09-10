using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class ChangeRemoverCommand : OfficeCommand
    {
        protected abstract IRemover sourceRemover { get; }

        public ChangeRemoverCommand(IOfficeCommandable receiver)
            : base(receiver) { }

        public override void Execute()
        {
            Receiver.RemoveTool = sourceRemover;
        }
    }
}
