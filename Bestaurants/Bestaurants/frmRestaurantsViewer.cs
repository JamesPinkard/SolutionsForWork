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

        private Dictionary<string, double> infreqMCL = new Dictionary<string, double>
        {
            {"ALUMINUM", 200},
            {"ANTIMONY", 6},
            {"ARSENIC", 10},
            {"BARIUM", 200},
            {"CADMIUM", 5},
            {"CHROMIUM", 100},
            {"COPPER", 1000},
            {"IRON_DISS", 300},
            {"LEAD", 15},
            {"MANGNESE", 50},
            {"MERCURY", 2},
            {"NO2-N_DISS", 1},
            {"SELENIUM", 50},
            {"SILVER", 100},
            {"ZINC", 5000}
        };

        private Dictionary<string, string> units = new Dictionary<string, string>
        {
            {"ALUMINUM", "ug/l"},
            {"ANTIMONY", "ug/l"},
            {"ARSENIC", "ug/l"},
            {"BARIUM", "ug/l"},
            {"CADMIUM", "ug/l"},
            {"CHROMIUM", "ug/l"},
            {"COPPER", "ug/l"},
            {"IRON_DISS", "ug/l"},
            {"LEAD", "ug/l"},
            {"MANGNESE", "ug/l"},
            {"MERCURY", "ug/l"},
            {"NO2-N_DISS", "mg/l"},
            {"SELENIUM", "ug/l"},
            {"SILVER", "ug/l"},
            {"ZINC", "ug/l"},
            {"SO4", "mg/l"},
            {"Cl", "mg/l"},
            {"Fl", "mg/l"},
            {"NO3", "mg/l"},
            {"pH", "pH units"},
            {"TDS", "mg/l"}
        };

        private Dictionary<string, string> titleLine = new Dictionary<string, string>
        {
            {"ALUMINUM", "ALUMINUM"},
            {"ANTIMONY", "ANTIMONY"},
            {"ARSENIC", "ARSENIC"},
            {"BARIUM", "BARIUM"},
            {"CADMIUM", "CADMIUM"},
            {"CHROMIUM", "CHROMIUM"},
            {"COPPER", "COPPER"},
            {"IRON_DISS", "IRON_DISS"},
            {"LEAD", "LEAD"},
            {"MANGNESE", "MANGNESE"},
            {"MERCURY", "MERCURY"},
            {"NO2-N_DISS", "NO<sub>2</sub>-N_DISS"},
            {"SELENIUM", "SELENIUM"},
            {"SILVER", "SILVER"},
            {"ZINC", "ZINC"},
            {"SO4", "SO<sub>4</sub>"},
            {"Cl", "Cl"},
            {"Fl", "Fl"},
            {"NO3", "NO<sub>3</sub>"},
            {"pH", "pH"},
            {"TDS", "TDS"}
        };

        private Dictionary<string, string> mclSource = new Dictionary<string, string>
        {
            {"ALUMINUM", "Secondary Standard"},
            {"ANTIMONY", "Primary Standard"},
            {"ARSENIC", "Primary Standard"},
            {"BARIUM", "Primary Standard"},
            {"CADMIUM", "Primary Standard"},
            {"CHROMIUM", "Primary Standard"},
            {"COPPER", "Secondary Standard"},
            {"IRON_DISS", "Secondary Standard"},
            {"LEAD", "Primary Standard"},
            {"MANGNESE", "Secondary Standard"},
            {"MERCURY", "Primary Standard"},
            {"NO2-N_DISS", "Primary Standard"},
            {"SELENIUM", "Primary Standard"},
            {"SILVER", "Primary Standard"},
            {"ZINC", "Primary Standard"},
            {"SO4", "Secondary Standard"},
            {"Cl", "Secondary Standard"},
            {"Fl", "Secondary Standard"},
            {"NO3", "Primary Standard"},
            {"pH", "Secondary Standard"},
            {"TDS", "Secondary Standard"}
        };

        private Dictionary<string, double> majorMCL = new Dictionary<string, double>
        {
            {"SO4", 300},
            {"Cl", 300},
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
            var classBreaks = new List<double> { 250, 500, 1000, 3000, 5000, 10000 };
            string templatePath = @"S:\Projects\TARRANT-Witchita\GIS\Layers\major_number_layers\Layer_{0}.lyr";
            SetClassBreaks(classBreaks, templatePath);

            string templateQuery = "AQ_LAY =  {0} AND PARM = '{1}' AND VALUE <>  -1";
            SetGroupDefinitionQuery(templateQuery, "TDS");

            double tdsMcl = 1000;
            SetLegendTitle("TDS", tdsMcl);
            RefreshMap();
            PrintMap("TDS");
            SaveChemMap("TDS");
            MessageBox.Show("Done");
        }

        private void btnPrintWQ_Click(object sender, EventArgs e)
        {

            string templatePath = @"S:\Projects\TARRANT-Witchita\GIS\Layers\major_number_layers\Layer_{0}.lyr";
            foreach (KeyValuePair<string, double> entry in majorMCL)
            {
                List<double> multipliers = new List<double> { 0.25, 0.5, 1, 2, 4, 10 };
                List<double> classBreaks = new List<double>();
                double maxContaminant = entry.Value;

                if (entry.Key == "pH")
                {
                    classBreaks = new List<double> { 5, 6, 7, 8, 9, 14 };
                }
                else
                {
                    foreach (var m in multipliers)
                    {
                        classBreaks.Add(maxContaminant * m);
                    }
                }

                SetClassBreaks(classBreaks, templatePath);

                string templateQuery = "AQ_LAY =  {0} AND PARM = '{1}' AND VALUE <>  -1";
                SetGroupDefinitionQuery(templateQuery, entry.Key);
                SetLegendTitle(entry.Key, entry.Value);
                RefreshMap();
                PrintMap(entry.Key);
                SaveChemMap(entry.Key);                
            }
            MessageBox.Show("Done");
        }

        private void btnPrintInfreq_Click(object sender, EventArgs e)
        {
            string templatePath = @"S:\Projects\TARRANT-Witchita\GIS\Layers\minor_number_layers\Layer_{0}.lyr";
            foreach (KeyValuePair<string, double> entry in infreqMCL)
            {
                List<double> multipliers = new List<double> { 0.25, 0.5, 1, 2, 4, 10 };
                List<double> classBreaks = new List<double>();
                double maxContaminant = entry.Value;

                foreach (var m in multipliers)
                {
                    classBreaks.Add(maxContaminant * m);
                }

                SetClassBreaks(classBreaks, templatePath);
                string templateQuery = "Layer  =  {0} AND wq_nam  = '{1}'";
                SetGroupDefinitionQuery(templateQuery, entry.Key);
                SetLegendTitle(entry.Key, entry.Value);
                RefreshMap();
                PrintMap(entry.Key);
                SaveChemMap(entry.Key);
            }
            MessageBox.Show("Done");
        }




        private void SetClassBreaks(List<double> classBreaks, string templatePath)
        {
            ILayer selectedLayer = GetLayerByName(cmbLayers.SelectedItem.ToString());
            ICompositeLayer groupLayer = (ICompositeLayer)selectedLayer;
            int groupCount = groupLayer.Count;

            for (int i = 0; i < groupCount; i++)
            {
                IGeoFeatureLayer geoFeatureLayer = (IGeoFeatureLayer)groupLayer.Layer[i];
                IGxLayer gxLayer = new GxLayer();
                IGxFile gxFile = (IGxFile)gxLayer;

                gxFile.Path = string.Format(templatePath, i + 1);

                IGeoFeatureLayer gxFeatureLayer = (IGeoFeatureLayer)gxLayer.Layer;
                IClassBreaksRenderer gxRenderer = (IClassBreaksRenderer)gxFeatureLayer.Renderer;

                geoFeatureLayer.Renderer = (IFeatureRenderer)gxRenderer;
                IClassBreaksRenderer geoRenderer = (IClassBreaksRenderer)geoFeatureLayer.Renderer;
                geoRenderer.BreakCount = classBreaks.Count;
                // geoRenderer = gxRenderer;                

                int breakCount = classBreaks.Count;

                geoRenderer.BreakCount = breakCount;

                for (int j = 0; j < breakCount; j++)
                {
                    geoRenderer.Break[j] = classBreaks[j];
                    geoRenderer.Symbol[j] = gxRenderer.Symbol[j];

                    if (j == 0)
                    {
                        int maxRange = Convert.ToInt32(geoRenderer.Break[j]);
                        geoRenderer.Label[j] = string.Format("<={0}", maxRange);
                    }
                    else
                    {
                        int minRange = Convert.ToInt32(geoRenderer.Break[j - 1]);
                        int maxRange = Convert.ToInt32(geoRenderer.Break[j]);
                        geoRenderer.Label[j] = string.Format("{0} to {1}", minRange, maxRange);
                    }
                }
            }
        }

        private void SetGroupDefinitionQuery(string templateQuery, string chem)
        {
            ILayer selectedLayer = GetLayerByName(cmbLayers.SelectedItem.ToString());
            ICompositeLayer groupLayer = (ICompositeLayer)selectedLayer;
            int groupCount = groupLayer.Count;


            for (int i = 0; i < groupCount; i++)
            {
                IFeatureLayerDefinition featureDefinition = (IFeatureLayerDefinition)groupLayer.Layer[i];
                string featureName = groupLayer.Layer[i].Name;
                featureDefinition.DefinitionExpression = string.Format(templateQuery, layerDictionary[featureName], chem);
            }
        }

        private void SetLegendTitle(string chem, double mcl)
        {
            string legendName = "Chem Legend";
            ILegend chemLegend = GetLegendByName(legendName);

            string templateLegendTitle = "{0}\r\n{1} ({2})\r\nMCL = {3} ({2})";
            chemLegend.Title = string.Format(templateLegendTitle, mclSource[chem], titleLine[chem], units[chem], mcl);
        }

        private void RefreshMap()
        {
            IMxDocument mxd = (IMxDocument)m_application.Document;
            mxd.UpdateContents();
            mxd.ActivatedView.Refresh();
        }

        private void PrintMap(string chem)
        {
            IMxDocument mxd = (IMxDocument)m_application.Document;
            IActiveView docActiveView = mxd.ActivatedView;
            IExport docExport = new ExportPNGClass();
            IPrintAndExport docPrintExport = new PrintAndExport();
            docExport.ExportFileName = string.Format(@"S:\Projects\TARRANT-Witchita\GIS\Figures\WaterQuality\{0}_{1}.png", cmbLayers.SelectedItem.ToString(), chem);
            docPrintExport.Export(docActiveView, docExport, 300, false, null);
        }

        private void SaveChemMap(string chem)
        {
            IMxDocument mxd = (IMxDocument)m_application.Document;
            string mxdPath = string.Format(@"S:\Projects\TARRANT-Witchita\GIS\mxds\WaterQuality_Maps\Individual_Maps\{0}_{1}.mxd", cmbLayers.SelectedItem.ToString(), chem);
            m_application.SaveAsDocument(mxdPath, true);
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


    }
}
