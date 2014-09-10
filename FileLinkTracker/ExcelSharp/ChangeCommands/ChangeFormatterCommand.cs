using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class ChangeFormatterCommand : OfficeCommand
    {
        protected abstract IFormatter sourceFormatter { get; }

        public ChangeFormatterCommand(IOfficeCommandable receiver) : base(receiver) { }

        public override void Execute()
        {
            Receiver.Formatter = sourceFormatter;
        }
    }
}
