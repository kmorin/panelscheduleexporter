﻿using System;
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
    private int _rowPointer = 1;
    //private FamilyInstance _panel = null;    

    public XLSXTranslator(PanelScheduleView psView, Document _doc) {
      //m_psView = psView;
      ElementId psId = psView.Id;
      m_psView = _doc.GetElement(psId) as PanelScheduleView;
    }

    public override string Export() {
      //_panel = fi;
      //string excelTemplate = assemblyName.Replace("PanelSchedule.dll", "panelSchedTemplate.xlsx");
#if REVIT2019 || REVIT2020 || REVIT2021
      string viewName = m_psView.Name;
#else
      string viewName = m_psView.ViewName;
#endif
      string newFileName = Path.Combine(PanelScheduleExport._exportDirectory, viewName + ".xlsx");

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

      try {
        //Initialize excel
        using (var wb = new ExcelPackage(new FileInfo(newFileName))) {
          var ws = wb.Workbook.Worksheets.Add(viewName);
          _ws = ws;
          DumpPanelScheduleData();

          _ws.Cells.AutoFitColumns();
          wb.Save();
        }

        //Working On portion (does nothing)??
        var psData = m_psView.GetTableData();
        int numSlots = psData.NumberOfSlots;              //Number of circuits
        int numCktRows = psData.GetNumberOfCircuitRows(); //example: 42ckt = 22 numCktRows            

        return newFileName;
      }
      catch (Exception ex) {
        System.Exception exc = ex as System.Exception;
        if (exc.Message.Contains("0x800A03EC"))//exc.HResult == -2146827284
        {
          throw new IOException();
        }
        else
          throw new SystemException();
      }
    }

    private void SetNewFileName(string inputFilename, out string newFileName) {
      //check if file already appends an iterator.
      // Remove extension
      var s = inputFilename.Remove(inputFilename.Length - 5);

      var i = 1;
      var isIterator = false;
      var list = s.Split('_').ToList();
      if (list.Count > 1) {
        isIterator = int.TryParse(list.Last(), out i);
      }

      if (isIterator) {
        //bump number
        i++;
      }
      list.Add(i.ToString());

      //Set the new fileName
      newFileName = string.Join("_", list) + ".xlsx";
    }

    private void LogFailedFile(string newFileName) {
      TaskDialog.Show("Failed", $"Failed to overwrite existing file: {newFileName} \nPlease make to close the worksheet.", TaskDialogCommonButtons.Ok);
    }

    private void DumpPanelScheduleData() {

      DumpSectionData(m_psView, SectionType.Header);
      DumpSectionData(m_psView, SectionType.Body);
      DumpSectionData(m_psView, SectionType.Summary);
      DumpSectionData(m_psView, SectionType.Footer);      

    }

    private void DumpSectionData(PanelScheduleView psView, SectionType sectionType) {

      var data = psView.GetSectionData(sectionType);
      _nRows_Section = data.NumberOfRows;
      _nCols_Section = data.NumberOfColumns;

      //Existing functionality
      for (int i = data.FirstRowNumber; i < _nRows_Section; i++) {
        for (int j = data.FirstColumnNumber; j < _nCols_Section; j++) {
          try {
            _s = m_psView.GetCellText(sectionType, i, j);
            //var cellTypeIsInPhaseLoads = m_psView.IsCellInPhaseLoads(i, j);
            //var cellIsInLoadSummary = m_psView.IsColumnInLoadSummary(j);
            var range = _ws.Cells[_rowPointer, j];
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            if (sectionType == SectionType.Body || sectionType == SectionType.Summary) {
              // Format cells for VA
              if (Regex.Match(_s, @"^[0-9]+\s(VA){1}[^a-z|A-Z]").Success) //Regex to match "## VA" cells
              {
                _s = _s.Remove(_s.Length - 3);
                int.TryParse(_s, out int value);

                range.Style.Numberformat.Format = "0 VA";
                range.Value = value;
              }
              //BUGFIX: 1 - make sure no trailing chars before or after.
              else if (Regex.Match(_s, @"^[0-9]+\s(A|kA|mA){1}[^a-z|A-Z]").Success) //Regex to match "## A | # kA | # mA" cells
              {
                _s = _s.Remove(_s.Length - 2);
                int.TryParse(_s, out int value);
                range.Value = _s;
              }
              else {
                range.Value = _s;
              }
            }
            else {
              range.Value = _s;
            }
          }
          catch (Exception) {
            // do nothing.
          }
        }
        _rowPointer++;
      }
    }
  }
}
