using System;
using System.IO;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PanelScheduleExporter
{
    class App : IExternalApplication
    {
        static string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        static string assyPath = Path.Combine(dir, "PanelScheduleExporter.dll");
        static string _imgFolder = Path.Combine(dir, "Images");

        public Result OnStartup(UIControlledApplication a)
        {
            try
            {
                AddRibbonPanel(a);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ribbon", ex.ToString());                
            }
            return Result.Succeeded;
        }

        private void AddRibbonPanel(UIControlledApplication app)
        {
            RibbonPanel panel = app.CreateRibbonPanel("Panel Schedule Exporter");

            PushButtonData pbd_Export = new PushButtonData("Export Panel Schedules", "Export Panel Schedules", assyPath, "PanelScheduleExporter.PanelScheduleExport");
            PushButton pb_Ex = panel.AddItem(pbd_Export) as PushButton;
            pb_Ex.LargeImage = NewBitmapImage("panelScheduleExporter.png");
            pb_Ex.ToolTip = "Export project Panel Schedules to Exel (XLSX) files";
            pb_Ex.LongDescription = "Opens dialog box to select project panel schedules for exporting out to Excel documents in a folder specified by the user.";

            ContextualHelp help = new ContextualHelp(ContextualHelpType.ChmFile, dir + "/help.htm");
            pb_Ex.SetContextualHelp(help);
        }

        BitmapImage NewBitmapImage(string imgName)
        {
            return new BitmapImage(new Uri(Path.Combine(_imgFolder,imgName)));
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
