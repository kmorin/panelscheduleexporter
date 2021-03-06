﻿using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PanelScheduleExporter2016
{
    class XLSXTranslator : Translator
    {
      public string _s = "";
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
            string newFileName = Path.Combine(PanelScheduleExport._exportDirectory,m_psView.ViewName + ".xlsx");

            try
            {
                //Initialize excel
                using (var wb = new ExcelPackage(new FileInfo(newFileName)))
                {
                    var ws = wb.Workbook.Worksheets.Add(m_psView.ViewName);
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
                PanelScheduleData psData = m_psView.GetTableData();
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
