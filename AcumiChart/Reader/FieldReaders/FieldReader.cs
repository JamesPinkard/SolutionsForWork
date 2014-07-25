using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reader.ProCode
{
    abstract public class FieldReader: FieldCommand
    {
        public string Source { get; set; }

        public FieldReader()
        {

        }
        public void Execute()
        {
            throw new NotImplementedException();
        }
    }
}
