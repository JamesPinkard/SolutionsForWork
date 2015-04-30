using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace GIS_AddIn
{
    public class AddGraphicsTool : ESRI.ArcGIS.Desktop.AddIns.Tool
    {
        public AddGraphicsTool()
        {
        }

        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
