using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public abstract class ChangeEmbedderCommand : OfficeCommand
    {
        protected abstract IEmbedder sourceTool { get; }

        public ChangeEmbedderCommand(IOfficeCommandable receiver)
            : base(receiver) { }      
        
        public override void Execute()
        {
            Receiver.EmbedTool = sourceTool;
        }
    }
}
