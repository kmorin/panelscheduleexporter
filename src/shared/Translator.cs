using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;

namespace PanelScheduleExporter
{
    /// <summary>
    /// Translate the panel schedule view data from Revit to some formats, HTML, CSV etc.
    /// </summary>
    abstract class Translator
    {
        /// <summary>
        /// the panel schedule view instance to be exported.
        /// </summary>
        public PanelScheduleView m_psView;

        public abstract string Export();        

        /// <summary>
        /// An utility method to replace illegal characters of the Panel Schedule view name.
        /// </summary>
        /// <param name="stringWithIllegalChar">the Panel Schedule view name.</param>
        /// <returns>the updated string without illegal characters.</returns>
        protected string ReplaceIllegalCharacters(string stringWithIllegalChar)
        {
            char[] illegalChars = System.IO.Path.GetInvalidFileNameChars();

            string updated = stringWithIllegalChar;
            foreach (char ch in illegalChars)
            {
                updated = updated.Replace(ch, '_');
            }

            return updated;
        }

        /// <summary>
        /// An utility method to get the number of rows and columns of the section which is exporting.
        /// </summary>
        /// <param name="doc">Revit document.</param>
        /// <param name="psView">the exporting panel schedule view</param>
        /// <param name="sectionType">the exporting section of the panel schedule.</param>
        /// <returns><!--<rows,cols>--></returns>
        protected Tuple<int,int> getNumberOfRowsAndColumns(Autodesk.Revit.DB.Document doc, PanelScheduleView psView, SectionType sectionType)
        {
            var openSectionData = new Transaction(doc, "openSectionData");
            openSectionData.Start();

            TableSectionData sectionData = psView.GetSectionData(sectionType);
            var nRows = sectionData.NumberOfRows;
            var nCols = sectionData.NumberOfColumns;

            openSectionData.RollBack();
            return new Tuple<int, int>(nRows, nCols);
        }
    }
}
