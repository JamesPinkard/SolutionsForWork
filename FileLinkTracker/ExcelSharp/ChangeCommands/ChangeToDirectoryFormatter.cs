using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public class ChangeToDirectoryFormatter : ChangeFormatterCommand 
    {

        public ChangeToDirectoryFormatter(IOfficeCommandable receiver)
            : base(receiver) { }

        protected override IFormatter sourceFormatter
        {
            get { return new DirectoryFormatter(); }
        }
    }
}
