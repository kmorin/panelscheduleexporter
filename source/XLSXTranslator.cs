using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PanelScheduleExporter
{
    class XLSXTranslator : Translator
    {
      public string _s = "";
        //private static Microsoft.Office.Interop.Excel.Workbook MyBook = null;
        //private static Microsoft.Office.Interop.Excel.Application MyApp = null;
        //private static Microsoft.Office.Interop.Excel.Worksheet MySheet = null;
        private static ExcelWorksheet _ws = null;
        private int _nRows_Section;
        private int _nCols_Section;
        private int rr = 0;
        private int cc = 0;
        //private FamilyInstance _panel = null;    

        public XLSXTranslator(PanelScheduleView psView, Document _doc)
        {
            //m_psView = psView;
          ElementId psId = psView.Id;
          m_psView = _doc.GetElement(psId) as PanelScheduleView;
        }

        public override string Export()
        {
      //_panel = fi;
      //string excelTemplate = assemblyName.Replace("PanelSchedule.dll", "panelSchedTemplate.xlsx");
#if REVIT2019 || REVIT2020
      string viewName = m_psView.Name;
#else
      string viewName = m_psView.ViewName;
#endif
      string newFileName = Path.Combine(PanelScheduleExport._exportDirectory, viewName + ".xlsx");

            var overwrite = false;

            // Attempt to delete the file if already exists.
            if (File.Exists(newFileName)) {
                try {
                    File.Delete(newFileName);
                }
                catch {
                    try {
                    File.Delete(newFileName);
                    }
                    catch {
                        LogFailedFile(newFileName);
                        //Failed to delete the file, so just append and integer
                        //SetNewFileName(newFileName, out newFileName);
                    }
                }
            }

            try
            {
                //Initialize excel
                using (var wb = new ExcelPackage(new FileInfo(newFileName)))
                {
                    var ws = wb.Workbook.Worksheets.Add(viewName);
                    _ws = ws;
                    DumpPanelScheduleData();

                    _ws.Cells.AutoFitColumns();
                    wb.Save();
                }


                //Implement dump to excel for testing
                //DumpPanelScheduleData();

                //Resize cells
                //MySheet.Columns.AutoFit();
                //Save out excel to new file            
                //MyBook.SaveAs(Path.Combine(PanelScheduleExport._exportDirectory, newFileName));
                //MyBook.Close();

                //Working On portion (does nothing)??
                var psData = m_psView.GetTableData();
                int numSlots = psData.NumberOfSlots;              //Number of circuits
                int numCktRows = psData.GetNumberOfCircuitRows(); //example: 42ckt = 22 numCktRows            

                return newFileName;
            }
            catch (Exception ex)
            {
                System.Exception exc = ex as System.Exception;
                if (exc.Message.Contains("0x800A03EC"))//exc.HResult == -2146827284
                {                    
                    throw new IOException();
                }
                else
                    throw new SystemException();
            }
        }

    private void SetNewFileName(string inputFilename, out string newFileName)
    {
      //check if file already appends an iterator.
      // Remove extension
      var s = inputFilename.Remove(inputFilename.Length - 5);

      var i = 1;
      var isIterator = false;
      var list = s.Split('_').ToList();
      if (list.Count > 1)
      {
        isIterator = int.TryParse(list.Last(), out i);
      }

      if (isIterator)
      {
        //bump number
        i++;
      }
      list.Add(i.ToString());

      //Set the new fileName
      newFileName = string.Join("_", list) + ".xlsx";
    }

    private void LogFailedFile(string newFileName)
    {
      TaskDialog.Show("Failed", $"Failed to overwrite existing file: {newFileName} \nPlease make to close the worksheet.", TaskDialogCommonButtons.Ok);
    }

    private void DumpPanelScheduleData()
        {

            DumpSectionData(m_psView, SectionType.Header);
            DumpSectionData(m_psView, SectionType.Body);
            DumpSectionData(m_psView, SectionType.Summary);
            DumpSectionData(m_psView, SectionType.Footer);

            //ElectricalEquipment eePanel = _panel.MEPModel as ElectricalEquipment;
            //FilteredElementCollector fec = new FilteredElementCollector(_panel.Document)
            //.OfCategory(BuiltInCategory.OST_ElectricalCircuit);
            
            //ElectricalSystemSet eset = new ElectricalSystemSet();

            //foreach (ElectricalSystem es in fec)
            //{
            //    if (es.PanelName == _panel.get_Parameter("Panel Name").ToString())
            //    {
            //        eset.Insert(es);
            //    }                
            //}            

        }

        private void DumpSectionData(PanelScheduleView psView, SectionType sectionType)
        {
            _nRows_Section = 0;
            _nRows_Section = 0;            

            getNumberOfRowsAndColumns(m_psView.Document, m_psView, sectionType, ref _nRows_Section, ref _nCols_Section); //get rows/cols for schedule section.
            
            //Header Section

            //Body Section

            //Summary

            //Footer
           
            //Existing functionality
            for (int ii = 0; ii < _nRows_Section ; ++ii)
            {
                for (int jj = 0; jj < _nCols_Section; ++jj)
                {
                    try
                    {
                        cc = jj; //set excel column equal to schedule column
                        _s = m_psView.GetCellText(sectionType, ii, jj);  
                        int value = 0;
                        //Microsoft.Office.Interop.Excel.Range range = MySheet.Cells[rr, cc] as Microsoft.Office.Interop.Excel.Range;
                        var range = _ws.Cells[rr, cc];
                        //range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft; //Align cells left.
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;                        
                        // Format cells for VA
                        if (Regex.Match(_s,@"[0-9] VA").Success) //Regex to match "## VA" cells
                        {
                            _s = _s.Remove(_s.Length - 3);                            
                            int.TryParse(_s, out value);
                            
                            range.Style.Numberformat.Format = "0 VA";                            
                            //MySheet.Cells[rr, cc] = value;                            
                            range.Value = value;
                        }
                        else if (Regex.Match(_s, @"[0-9] A").Success) //Regex to match "## A" cells
                        {
                            _s = _s.Remove(_s.Length - 2);
                            int.TryParse(_s, out value);

                            range.Style.Numberformat.Format = "0 A";
                            //MySheet.Cells[rr, cc] = value;                            
                            range.Value = value;
                        }
                        else
                        {
                            //MySheet.Cells[rr, cc] = _s;       
                            range.Value = _s;
                        }                        
                    }
                    catch (Exception)
                    {
                        // do nothing.
                    }                    
                }
                rr++; //increment excel row   
            }
        }
    }
}
