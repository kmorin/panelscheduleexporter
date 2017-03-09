using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace PanelScheduleExporter2015
{
    internal class App : IExternalApplication
    {
        private readonly string _path = Path.GetDirectoryName(
          Assembly.GetExecutingAssembly().Location);
        public static Assembly _assy = Assembly.GetExecutingAssembly
        ();
        public static string _assemblyName = _assy.GetName().Name;
        public static string _assyVersion = _assy.GetName().Version.ToString();

        public Result OnStartup(UIControlledApplication a)
        {
            try
            {
                AddRibbonPanel(a);
            }
            catch (Exception ex)
            {
                //TaskDialog.Show("Ribbon", ex.ToString());                
            }
            return Result.Succeeded;
        }

        private void AddRibbonPanel(UIControlledApplication app)
        {
            RibbonPanel panel = app.CreateRibbonPanel("Panel Schedule Exporter");

            string m_iconPath = string.Join(".",
                _assy.GetTypes().First().Namespace,
                "icons")+".";

            AddButton(panel,
                "exportpanelschedules",
                "Export Panel Schedules",
                string.Concat(m_iconPath, "panelScheduleExport_16.png"),
                string.Concat(m_iconPath, "panelScheduleExport.png"),
                Path.Combine(_path,_assemblyName+".dll"),
                "PanelScheduleExporter2015.PanelScheduleExport",
                "Export project Panel Schedules to Exel (XLSX) files",
                "",
                false
                );
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }

        #region private members

        /// <summary>
        ///   Add a pushbutton to a panel
        /// </summary>
        /// <param name="rPanel"></param>
        /// <param name="buttonName"></param>
        /// <param name="buttonText"></param>
        /// <param name="imagePath16"></param>
        /// <param name="imagePath32"></param>
        /// <param name="dllPath"></param>
        /// <param name="dllClass"></param>
        /// <param name="toolTip"></param>
        /// <param name="pbAvail"></param>
        /// <param name="separatorBeforeButton"></param>
        private void AddButton(RibbonPanel rPanel,
          string buttonName,
          string buttonText,
          string imagePath16,
          string imagePath32,
          string dllPath,
          string dllClass,
          string toolTip,
          string pbAvail,
          bool separatorBeforeButton)
        {
            //File path must exist
            if (!File.Exists(dllPath)) return;

            //Separator??
            if (separatorBeforeButton) rPanel.AddSeparator();

            try
            {
                //Create the pbData
                PushButtonData m_mPushButtonData = new PushButtonData(
                  buttonName,
                  buttonText,
                  dllPath,
                  dllClass);

                if (!string.IsNullOrEmpty(imagePath16))
                {
                    try
                    {
                        m_mPushButtonData.Image = LoadPngImageSource(imagePath16);
                    }
                    catch (Exception m_e)
                    {
                        throw new Exception(m_e.Message);
                    }
                }

                if (!string.IsNullOrEmpty(imagePath32))
                {
                    try
                    {
                        m_mPushButtonData.LargeImage = LoadPngImageSource(imagePath32);
                    }
                    catch (Exception m_e)
                    {
                        throw new Exception(m_e.Message);
                    }
                }

                m_mPushButtonData.ToolTip = toolTip;

                //Availability?
                if (!string.IsNullOrEmpty(pbAvail))
                {
                    m_mPushButtonData.AvailabilityClassName = pbAvail;
                }

                //Add button to the ribbon
                rPanel.AddItem(m_mPushButtonData);
                ContextualHelp help = new ContextualHelp(ContextualHelpType.ChmFile, _path + "/help.htm");
                m_mPushButtonData.SetContextualHelp(help);
            }
            catch (Exception m_e)
            {
                throw new Exception(m_e.Message);
            }
        }

        /// <summary>
        ///   Load the PNG image from file
        /// </summary>
        /// <param name="sourceName"></param>
        /// <returns></returns>
        private ImageSource LoadPngImageSource(string sourceName)
        {
            try
            {
                //Assembly and stream
                Assembly m_assembly = Assembly.GetExecutingAssembly();
                Stream m_icon = m_assembly.GetManifestResourceStream(sourceName);

                //Decode
                if (m_icon != null)
                {
                    PngBitmapDecoder m_decoder = new PngBitmapDecoder(
                      m_icon,
                      BitmapCreateOptions.PreservePixelFormat,
                      BitmapCacheOption.Default);

                    //Source
                    ImageSource m_source = m_decoder.Frames[0];
                    return (m_source);
                }
            }
            catch (Exception m_e)
            {
                throw new Exception(m_e.Message);
            }
            return null;
        }

        #endregion
    }
}
