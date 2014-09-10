using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public interface IOfficeCommandable
    {
        IEmbedder EmbedTool { get; set; }
        IRemover RemoveTool { get; set; }
        IOfficeWriter Writer { get; set; }
        IFormatter Formatter { get; set; }

        bool isValid();
    }
}
