using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class ChangeRemoverCommand : OfficeCommand
    {
        private IRemover sourceRemover;

        public ChangeRemoverCommand(IOfficeCommandable receiver, IRemover remover)
            :base(receiver)
        {
            this.sourceRemover = remover;
        }

        public override void Execute()
        {
            throw new NotImplementedException();
        }
    }
}
