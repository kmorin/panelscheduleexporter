using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace PanelScheduleExporter2018
{
    /// <summary>
    /// Create view instance for an electrical panel.
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class InstanceViewCreation : IExternalCommand
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

            Reference selected = commandData.Application.ActiveUIDocument.Selection.PickObject(ObjectType.Element);

            Transaction newInstanceView = new Transaction(doc, "Create instance view for an electrical panel.");
            newInstanceView.Start();
            PanelScheduleView instanceView = PanelScheduleView.CreateInstanceView(doc, doc.GetElement(selected).Id);
            if (null == instanceView)
            {
                newInstanceView.RollBack();
                message = "Please select one electrical panel.";
                return Result.Failed;
            }
            else
            {
                newInstanceView.Commit();
                return Result.Succeeded;
            }

        }
    }
}
