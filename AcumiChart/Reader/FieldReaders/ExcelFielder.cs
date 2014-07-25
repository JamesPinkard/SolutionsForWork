using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reader.ProCode
{
    public class ExcelFielder
    {
        public IFieldManager FieldManager { get; private set; }
        public string FileSource { get; set; }
        public string WorksheetSource { get; set; }

        
        public ExcelFielder(IFieldManager fieldManager)
        {
            this.FieldManager = fieldManager;
        }

        public void Read()
        {

        }
    }
}
