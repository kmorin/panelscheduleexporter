using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;

namespace PanelScheduleExporter
{
    /// <summary>
    /// Import the panel scheduel view to place on a sheet view.
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    class SheetImport : IExternalCommand
    {
        /// <summary>
        /// Implement this method as an external command for Revit.
        /// </summary>
        /// <param name="commandData">An object that is passed to the external application 
        /// which contains data related to the command, 
        /// such as the application object and active view.</param>
        /// <param name="message">A message that can be set by the external application 
        /// which will be displayed if a failure or cancellation is returned by 
        /// the external command.</param>
        /// <param name="elements">A set of elements to which the external application 
        /// can add elements that are to be highlighted in case of failure or cancellation.</param>
        /// <returns>Return the status of the external command. 
        /// A result of Succeeded means that the API external method functioned as expected. 
        /// Cancelled can be used to signify that the user cancelled the external operation 
        /// at some point. Failure should be returned if the application is unable to proceed with 
        /// the operation.</returns>
        public virtual Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData
            , ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;

            // get one sheet view to place panel schedule.
            ViewSheet sheet = doc.ActiveView as ViewSheet;
            if (null == sheet)
            {
                message = "please go to a sheet view.";
                return Result.Failed;
            }

            // get all PanelScheduleView instances in the Revit document.
            FilteredElementCollector fec = new FilteredElementCollector(doc);
            ElementClassFilter PanelScheduleViewsAreWanted = new ElementClassFilter(typeof(PanelScheduleView));
            fec.WherePasses(PanelScheduleViewsAreWanted);
            List<Element> psViews = fec.ToElements() as List<Element>;

            Transaction placePanelScheduleOnSheet = new Transaction(doc, "placePanelScheduleOnSheet");
            placePanelScheduleOnSheet.Start();

            XYZ nextOrigin = new XYZ(0.0, 0.0, 0.0);
            foreach (Element element in psViews)
            {
                PanelScheduleView psView = element as PanelScheduleView;
                if (psView.IsPanelScheduleTemplate())
                {
                    // ignore the PanelScheduleView instance which is a template.
                    continue;
                }

                PanelScheduleSheetInstance onSheet = PanelScheduleSheetInstance.Create(doc, psView.Id, sheet);
                onSheet.Origin = nextOrigin;
                BoundingBoxXYZ bbox = onSheet.get_BoundingBox(doc.ActiveView);
                double width = bbox.Max.X - bbox.Min.X;
                nextOrigin = new XYZ(onSheet.Origin.X + width, onSheet.Origin.Y, onSheet.Origin.Z);
            }

            placePanelScheduleOnSheet.Commit();

            return Result.Succeeded;

        }
    }
}
