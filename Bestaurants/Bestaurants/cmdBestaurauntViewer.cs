using System;
using System.Drawing;
using System.Runtime.InteropServices;
using ESRI.ArcGIS.ADF.BaseClasses;
using ESRI.ArcGIS.ADF.CATIDs;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;

namespace ArcMapClassLibrary1
{
    /// <summary>
    /// Summary description for cmdBestaurauntViewer.
    /// </summary>
    [Guid("3215f900-98f6-464a-8660-fb78f48d5b30")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ArcMapClassLibrary1.cmdBestaurauntViewer")]
    public sealed class cmdBestaurauntViewer : BaseCommand
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
        public cmdBestaurauntViewer()
        {
            //
            // TODO: Define values for the public properties
            //
            base.m_category = "Bestaurants_Tools"; //localizable text
            base.m_caption = "Bestaurants Viewer";  //localizable text
            base.m_message = "Viewer for Bestaurants";  //localizable text 
            base.m_toolTip = "Bestaurant Viewer";  //localizable text 
            base.m_name = "Bestaurants_Tools_Viewer";   //unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
            // TODO: Add cmdBestaurauntViewer.OnClick implementation
            frmRestaurantsViewer pRestaurantsViewer = new frmRestaurantsViewer();
            
            // show the form
            pRestaurantsViewer.ArcMapApplication = m_application;
            
            pRestaurantsViewer.Show();

            

            IMxDocument pMxDoc = m_application.Document as IMxDocument;
            IMap pMap = pMxDoc.FocusMap;
            
            int layerCount = pMap.LayerCount;
            for (int i = 0; i < layerCount; i++)
            {
                ILayer pLayer = pMap.Layer[i] as ILayer;
                pRestaurantsViewer.cmbLayers.Items.Add(pLayer.Name);
            }
            
            
            
            
        }

        #endregion
    }
}
