using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public class ChangeToLinkEmbedder : ChangeEmbedderCommand
    {
        protected override IEmbedder sourceTool
        {
            get { return new LinkEmbedder(); }
        }

        public ChangeToLinkEmbedder(IOfficeCommandable receiver) : base(receiver) { }
    }
}
