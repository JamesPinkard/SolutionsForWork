using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class ChangeWriterCommand : OfficeCommand
    {
        private IOfficeWriter sourceWriter;

        public ChangeWriterCommand(IOfficeCommandable receiver, IOfficeWriter writer)
            : base(receiver)
        {
            this.sourceWriter = writer;
        }
        
        public override void Execute()
        {
            throw new NotImplementedException();
        }
    }
}
