using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class ChangeToLinkRemover : ChangeRemoverCommand
    {
        protected override IRemover sourceRemover
        {
            get { return new LinkRemover(); }
        }

        public ChangeToLinkRemover(IOfficeCommandable receiver) : base(receiver) { }
    }
}
