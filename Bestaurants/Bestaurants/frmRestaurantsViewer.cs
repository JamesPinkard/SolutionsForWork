using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Catalog;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Output;

namespace ArcMapClassLibrary1
{
    public partial class frmRestaurantsViewer : Form
    {
        private IApplication m_application;

        private Dictionary<string, int> layerDictionary = new Dictionary<string, int>
            {
                {"Sparta Aquifer", 5},
                {"Weches Formation", 6},
                {"Queen City Aquifer", 7},
                {"Reklaw Formation", 8},
                {"Carrizo Aquifer", 9},
                {"Upper Wilcox Aquifer", 10},
                {"Middle Wilcox Aquifer", 11},
                {"Lower Wilcox Aquifer", 12},
                {"Woodbine Aquifer", 15},
                {"Washita/Fredericksburg Groups", 16},
                {"Paluxy Aquifer", 17},
                {"Glen Rose Formation", 18},
                {"Hensell Aquifer", 19},
                {"Pearsall Formation", 20},
                {"Hosston Aquifer", 21},
                {"Wichita Aquifer", 23},
                {"Cisco Bowie Aquifer", 24},
                {"Canyon Aquifer", 25},
                {"Strawn Aquifer", 26},
                {"Seymour Aquifer", 31},
                {"Blaine Aquifer", 32},
                {"Permian Aquifer", 33},
                {"Nacatoch Aquifer", 36}
            };

        private Dictionary<string, double> dictMCL = new Dictionary<string, double>
        {
            {"SO4", 250},
            {"Cl", 250},
            {"Fl", 4},
            {"NO3", 10},
            {"pH", 7.0},
        };

        IRgbColor White = MakeRGBColor(255, 252, 255);
        IRgbColor Purple = MakeRGBColor(174, 0, 255);
        IRgbColor Cyan = MakeRGBColor(48, 203, 255);
        IRgbColor Green = MakeRGBColor(85, 255, 0);
        IRgbColor Gold = MakeRGBColor(255, 179, 0);
        IRgbColor Poinsetta = MakeRGBColor(230, 0, 0);

        public IApplication ArcMapApplication
        {
            get { return m_application; }
            set { m_application = value; }
        }
        public frmRestaurantsViewer()
        {
            InitializeComponent();
        }

        private void frmRestaurantsViewer_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void cmbLayer_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnPrintTDS_Click(object sender, EventArgs e)
        {
            var classBreaks = new List<int> { 500, 1000, 2000, 3000, 5000, 10000 };
            string templatePath = @"S:\Projects\TARRANT-Witchita\GIS\Layers\major_layers_6Break\Layer_{0}.lyr";
            SetClassBreaks(classBreaks, templatePath);
            SetGroupDefinitionQuery();            
            RefreshMap();
            PrintMap();
        }

        private void btnPrintWQ_Click(object sender, EventArgs e)
        {
            var classBreaks = new List<int> { 500, 1000, 2000, 3000, 5000, 10000 };
            string templatePath = @"S:\Projects\TARRANT-Witchita\GIS\Layers\major_layers_6Break\Layer_{0}.lyr";
            foreach (KeyValuePair<string,double> entry in dictMCL)
            {
                //Todo Implement wq run
            }
            SetClassBreaks(classBreaks, templatePath);
            SetGroupDefinitionQuery();
            RefreshMap();
            PrintMap();
        }

        private void RefreshMap()
        {
            IMxDocument mxd = (IMxDocument)m_application.Document;
            mxd.UpdateContents();
            mxd.ActivatedView.Refresh();
        }

        private void SetColors() // Does not work
        {
            ILayer selectedLayer = GetLayerByName(cmbLayers.SelectedItem.ToString());
            ICompositeLayer groupLayer = (ICompositeLayer)selectedLayer;
            int groupCount = groupLayer.Count;

            for (int i = 0; i < groupCount; i++)
            {
                IGeoFeatureLayer geoFeatureLayer = (IGeoFeatureLayer)groupLayer.Layer[i];
                IClassBreaksRenderer geoRenderer = (IClassBreaksRenderer)geoFeatureLayer.Renderer;
                int breakCount = geoRenderer.BreakCount;

                for (int j = 0; j < breakCount; j++)
                {
                    IMultiLayerMarkerSymbol symbol = (IMultiLayerMarkerSymbol)geoRenderer.Symbol[j];
                    symbol.Layer[1].Color = White;                    
                }
            }
        }
        private void SetClassBreaks(List<int> classBreaks,string templatePath)
        {
            ILayer selectedLayer = GetLayerByName(cmbLayers.SelectedItem.ToString());
            ICompositeLayer groupLayer = (ICompositeLayer)selectedLayer;
            int groupCount = groupLayer.Count;

            for (int i = 0; i < groupCount; i++)
            {
                IGeoFeatureLayer geoFeatureLayer = (IGeoFeatureLayer)groupLayer.Layer[i];
                IGxLayer gxLayer = new GxLayer();
                IGxFile gxFile =(IGxFile)gxLayer;

                gxFile.Path = string.Format(templatePath, i + 1);

                IGeoFeatureLayer gxFeatureLayer = (IGeoFeatureLayer)gxLayer.Layer;
                IClassBreaksRenderer gxRenderer = (IClassBreaksRenderer)gxFeatureLayer.Renderer;

                geoFeatureLayer.Renderer = (IFeatureRenderer)gxRenderer;
                IClassBreaksRenderer geoRenderer = (IClassBreaksRenderer)geoFeatureLayer.Renderer;
                geoRenderer = gxRenderer;                

                int breakCount = classBreaks.Count;

                geoRenderer.BreakCount = breakCount;

                for (int j = 0; j < breakCount; j++)
                {
                    geoRenderer.Break[j] = classBreaks[j];

                    if (j == 0)
                    {
                        int maxRange = Convert.ToInt32(geoRenderer.Break[j]);
                        geoRenderer.Label[j] = string.Format("<={0}", maxRange);
                    }
                    else
                    {
                        int minRange = Convert.ToInt32(geoRenderer.Break[j-1]);
                        int maxRange = Convert.ToInt32(geoRenderer.Break[j]);
                        geoRenderer.Label[j] = string.Format("{0} to {1}",minRange, maxRange);  
                    }
                }
            }
        }

        private void PrintMap() 
        {
            IMxDocument mxd = (IMxDocument)m_application.Document;
            IActiveView docActiveView = mxd.ActivatedView;
            IExport docExport = new ExportPNGClass();
            IPrintAndExport docPrintExport = new PrintAndExport();
            docExport.ExportFileName = string.Format(@"S:\Projects\TARRANT-Witchita\GIS\Figures\WaterQuality\{0}_{1}",cmbLayers.SelectedItem.ToString(),"TDS");
            docPrintExport.Export(docActiveView, docExport, 300, false, null);
        }

        public static ESRI.ArcGIS.Display.IRgbColor MakeRGBColor(byte R, byte G, byte B)
        {
            ESRI.ArcGIS.Display.RgbColor RgbClr = new ESRI.ArcGIS.Display.RgbColorClass();
            RgbClr.Red = R;
            RgbClr.Green = G;
            RgbClr.Blue = B;
            return RgbClr;
        }

        public static ESRI.ArcGIS.Display.IHsvColor MakeHSVColor(byte H, byte S, byte V)
        {
            ESRI.ArcGIS.Display.HsvColor HsvClr = new ESRI.ArcGIS.Display.HsvColorClass();
            HsvClr.Hue = H;
            HsvClr.Saturation = S;
            HsvClr.Value = V;
            return HsvClr;
        }

        private void SetGroupDefinitionQuery() 
        {
            ILayer selectedLayer = GetLayerByName(cmbLayers.SelectedItem.ToString());
            ICompositeLayer groupLayer = (ICompositeLayer)selectedLayer;
            int groupCount = groupLayer.Count;
            
            string templateQuery = "AQ_LAY =  {0} AND PARM = '{1}' AND VALUE <>  -1";
            string chemString = "TDS";            

            for (int i = 0; i < groupCount; i++)
            {
                IFeatureLayerDefinition featureDefinition = (IFeatureLayerDefinition)groupLayer.Layer[i];
                string featureName = groupLayer.Layer[i].Name;
                featureDefinition.DefinitionExpression = string.Format(templateQuery, layerDictionary[featureName], chemString);
            }
                        
            string legendName = "Chem Legend";
            ILegend chemLegend = GetLegendByName(legendName);

            string templateLegendTitle = "{0}\nmg/l or ug/l";
            chemLegend.Title = string.Format(templateLegendTitle, chemString);
        }

        private ILayer GetLayerByName(string layerName)
        {
            IMxDocument mxd = (IMxDocument)m_application.Document;
            IMap map = (IMap)mxd.FocusMap;

            int layerCount = map.LayerCount;
            for (int i = layerCount - 1; i >= 0; i--)
            {
                ILayer currentLayer = map.Layer[i];
                if (currentLayer.Name == layerName)
                {
                    return currentLayer;
                }
            }

            return null;
        }

        private ILegend GetLegendByName(string legendName)
        {
            IMxDocument mxd = (IMxDocument)m_application.Document;
            IMap map = (IMap)mxd.FocusMap;
            int layerCount = map.MapSurroundCount;

            for (int i = layerCount - 1; i >= 0; i--)
            {
                if (map.MapSurround[i].Name == legendName)
                {
                    return (ILegend)map.MapSurround[i];
                }
            }

            return null;
        }

        
    }
}
