using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace MapExporter
{
    public class MapExporter : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public MapExporter()
        {
        }

        protected override void OnClick()
        {
            //
            //  TODO: Sample code showing how to access button host
            //
            ArcMap.Application.CurrentTool = null;
            FormExport pFormExport = new FormExport();
            pFormExport.ShowDialog();
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
