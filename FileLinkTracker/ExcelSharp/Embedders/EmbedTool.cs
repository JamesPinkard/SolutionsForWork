using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class EmbedTool : IEmbedder
    {
        public void Embed()
        {
            throw new NotImplementedException();
        }

        public Microsoft.Office.Interop.Excel._Worksheet worksheet
        {
            get { throw new NotImplementedException(); }
        }
    }
}
