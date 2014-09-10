using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class ChangeToolCommand : OfficeCommand
    {
        private IEmbedder sourceTool;
        
        public ChangeToolCommand(IOfficeCommandable receiver, IEmbedder embedTool)
            : base(receiver)
        {
            this.sourceTool = embedTool;
        }
        
        public override void Execute()
        {
            Receiver.EmbedTool = sourceTool;
        }
    }
}
