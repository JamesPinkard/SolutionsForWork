using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ESRI.ArcGIS.ADF.BaseClasses;
using ESRI.ArcGIS.ADF.CATIDs;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Catalog;

namespace ArcMapClassLibrary1
{
    /// <summary>
    /// Summary description for cmdLoadData.
    /// </summary>
    [Guid("1eb3a98e-ee76-4f50-9dbb-e1ad324638a1")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ArcMapClassLibrary1.cmdLoadData")]
    public sealed class cmdLoadData : BaseCommand
    {
        #region COM Registration Function(s)
        [ComRegisterFunction()]
        [ComVisible(false)]
        static void RegisterFunction(Type registerType)
        {
            // Required for ArcGIS Component Category Registrar support
            ArcGISCategoryRegistration(registerType);

            //
            // TODO: Add any COM registration code here
            //
        }

        [ComUnregisterFunction()]
        [ComVisible(false)]
        static void UnregisterFunction(Type registerType)
        {
            // Required for ArcGIS Component Category Registrar support
            ArcGISCategoryUnregistration(registerType);

            //
            // TODO: Add any COM unregistration code here
            //
        }

        #region ArcGIS Component Category Registrar generated code
        /// <summary>
        /// Required method for ArcGIS Component Category registration -
        /// Do not modify the contents of this method with the code editor.
        /// </summary>
        private static void ArcGISCategoryRegistration(Type registerType)
        {
            string regKey = string.Format("HKEY_CLASSES_ROOT\\CLSID\\{{{0}}}", registerType.GUID);
            MxCommands.Register(regKey);

        }
        /// <summary>
        /// Required method for ArcGIS Component Category unregistration -
        /// Do not modify the contents of this method with the code editor.
        /// </summary>
        private static void ArcGISCategoryUnregistration(Type registerType)
        {
            string regKey = string.Format("HKEY_CLASSES_ROOT\\CLSID\\{{{0}}}", registerType.GUID);
            MxCommands.Unregister(regKey);

        }

        #endregion
        #endregion

        private IApplication m_application;
        public cmdLoadData()
        {
            //
            // TODO: Define values for the public properties
            //
            base.m_category = "Bestaurants_Tools"; //localizable text
            base.m_caption = "Load Data";  //localizable text
            base.m_message = "Load Data into ArcMap from external files";  //localizable text 
            base.m_toolTip = "Load Data";  //localizable text 
            base.m_name = "Bestaurants_Tools_LoadData";   //unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

            try
            {
                //
                // TODO: change bitmap name if necessary
                //
                string bitmapResourceName = GetType().Name + ".bmp";
                base.m_bitmap = new Bitmap(GetType(), bitmapResourceName);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap");
            }
        }

        #region Overridden Class Methods

        /// <summary>
        /// Occurs when this command is created
        /// </summary>
        /// <param name="hook">Instance of the application</param>
        public override void OnCreate(object hook)
        {
            if (hook == null)
                return;

            m_application = hook as IApplication;

            //Disable if it is not ArcMap
            if (hook is IMxApplication)
                base.m_enabled = true;
            else
                base.m_enabled = false;

            // TODO:  Add other initialization code
        }

        /// <summary>
        /// Occurs when this command is clicked
        /// </summary>
        public override void OnClick()
        {
            // TODO: Add cmdLoadData.OnClick implementation
            MessageBox.Show("Hello World!");
            m_application.Caption = "Bestaurants Application";

            IMxDocument pMxDoc = m_application.Document as IMxDocument;
            IMap pMap = pMxDoc.FocusMap;

            int layer_count = pMap.LayerCount;
        }

        private void ChangeBreaks()
        {
            IMxDocument pMxDoc = m_application.Document as IMxDocument;
            IMap pMap = pMxDoc.FocusMap;
            IGeoFeatureLayer pFeatureLayer = pMap.Layer[0] as IGeoFeatureLayer;

            IGxLayer pGXLayer = new GxLayer();
            IGxFile pGXFile = (IGxFile)pGXLayer;

            pGXFile.Path = @"S:\Projects\TARRANT-Witchita\GIS\Layers\infreq_layers\test_Layer_2.lyr";

            IGeoFeatureLayer pLayer = pGXLayer.Layer as IGeoFeatureLayer;
            IClassBreaksRenderer pClassBreaksRenderer = (IClassBreaksRenderer)pLayer.Renderer;
            int pNumberOfClassBreaks = pClassBreaksRenderer.BreakCount;

            for (int i = 0; i < pNumberOfClassBreaks; i++)
            {
                pClassBreaksRenderer.Break[i] = i + 1;
            }

            IClassBreaksRenderer qClassBreaksRenderer = (IClassBreaksRenderer)pFeatureLayer.Renderer;
            qClassBreaksRenderer = pClassBreaksRenderer;
            int qNumberOfClassBreaks = qClassBreaksRenderer.BreakCount;

            for (int i = 0; i < qNumberOfClassBreaks; i++)
            {
                qClassBreaksRenderer.Break[i] = i + 1;
            }



            //Create a new LayerFile instance.
            ILayerFile layerFile = new LayerFileClass();
            //Create a new layer file.
            layerFile.New(@"S:\Projects\TARRANT-Witchita\GIS\Layers\test_Layer_3.lyr");

            layerFile.ReplaceContents(pLayer);
            layerFile.Save();  
        }
        #endregion
    }
}
