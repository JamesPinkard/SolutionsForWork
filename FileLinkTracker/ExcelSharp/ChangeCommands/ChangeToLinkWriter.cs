using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public class ChangeToLinkWriter : ChangeWriterCommand
    {
        protected override IOfficeWriter sourceWriter
        {
            get { return new LinkWriter(); }
        }

        public ChangeToLinkWriter(IOfficeCommandable receiver) : base(receiver) { }
    }
}
