using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class ChangeFormatterCommand : OfficeCommand
    {
        private IFormatter sourceFormatter;

        public ChangeFormatterCommand(IOfficeCommandable receiver, IFormatter formatter)
            : base(receiver)
        {
            this.sourceFormatter = formatter;
        }

        public override void Execute()
        {
            throw new NotImplementedException();
        }
    }
}
