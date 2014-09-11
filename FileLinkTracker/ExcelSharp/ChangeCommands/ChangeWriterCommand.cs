using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class ChangeWriterCommand : OfficeCommand
    {
        protected abstract IOfficeWriter sourceWriter { get; }

        public ChangeWriterCommand(IOfficeCommandable receiver)
            : base(receiver) { }
        
        public override void Execute()
        {
            Receiver.Writer = sourceWriter;
        }
    }
}
