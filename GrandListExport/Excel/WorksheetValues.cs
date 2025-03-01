using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace GrandListExport
{
    public class WorksheetValues
    {
        public ExcelPackage excelPackageInput;
        public OfficeOpenXml.ExcelWorksheet inputWorksheet;
        public InputFileType inputFileType;
        public int numRows;
        public int numCols;
        private object misValue = System.Reflection.Missing.Value;
        //****************************************************************************************************************************
        public WorksheetValues(string filenameWithoutExtension, string extension, ColumnName sortColumn)
        {
            inputFileType = SetInputFileType(filenameWithoutExtension);
            string workbookPathname = GetSortedFilename(filenameWithoutExtension, extension, sortColumn);
            OpenWithEPPlus(workbookPathname);
        }
        //****************************************************************************************************************************
        public WorksheetValues(string filenameWithoutExtension, string extension)
        {
            inputFileType = SetInputFileType(filenameWithoutExtension);
            string filename = filenameWithoutExtension + "." + extension;
            if (!File.Exists(filename))
            {
                throw new Exception("Input file does not exist: " + filename);
            }
            inputFileType = InputFileType.Other;
            OpenWithEPPlus(filename);
        }
        //****************************************************************************************************************************
        public WorksheetValues()
        {
        }
        //****************************************************************************************************************************
        public void SaveAsXlsFile(string workbookPathname)
        {
            if (!File.Exists(workbookPathname))
            {
                throw new Exception("Output File does not exist: " + workbookPathname);
            }
            Workbook workbook = OpenExcelWorkbook(workbookPathname, false);
            workbook.SaveAs(workbookPathname, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                            XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        }
        //****************************************************************************************************************************
        private void OpenWithEPPlus(string workbookPathname)
        {
            if (!File.Exists(workbookPathname))
            {
                throw new Exception("Input file does not exist: " + workbookPathname);
            }
            excelPackageInput = OpenExcelWorkbook(workbookPathname);
            if (excelPackageInput.Workbook.Worksheets.Count == 0)
            {
                throw new Exception("Input file does not exist: " + workbookPathname);
            }
            this.inputWorksheet = excelPackageInput.Workbook.Worksheets[1];
            this.numRows = inputWorksheet.Dimension.End.Row;
            this.numCols = inputWorksheet.Dimension.End.Column;
        }
        //****************************************************************************************************************************
        public void CloseWorkbook()
        {
            if (excelPackageInput != null)
            {
                excelPackageInput.Dispose();
                excelPackageInput = null;
            }
        }
        //****************************************************************************************************************************
        private InputFileType SetInputFileType(string path)
        {
            string filename = Path.GetFileNameWithoutExtension(path);
            if (filename.ToLower().Contains("nemrc"))
            {
                return InputFileType.Nemrc;
            }
            if (filename.ToLower().Contains("patriot"))
            {
                return InputFileType.Patriot;
            }
            if (filename.ToLower().Contains("taxmap"))
            {
                return InputFileType.TaxMap;
            }
            return InputFileType.Other;
        }
        //****************************************************************************************************************************
        protected ExcelPackage OpenExcelWorkbook(string workbookPathname)
        {
            var fi = new FileInfo(workbookPathname);
            ExcelPackage excelPackage = new ExcelPackage(fi);
            return excelPackage;
        }
        //****************************************************************************************************************************
        private string GetSortedFilename(string filenameWithoutExtension, string extension, ColumnName sortColumn)
        {
            string filename = filenameWithoutExtension + "." + extension;
            if (!File.Exists(filename))
            {
                throw new Exception("Input file does not exist: " + filename);
            }
            DateTime modifiedDate1 = File.GetLastWriteTime(filename);
            string workbookPathname = filenameWithoutExtension + ".xlsx";
            if (File.Exists(workbookPathname))
            {
                DateTime modifiedDate2 = File.GetLastWriteTime(workbookPathname);
                if (modifiedDate2 > modifiedDate1)
                {
                    return workbookPathname;
                }
            }
            Workbook workbook = OpenExcelWorkbook(filename, false);
            int sortIndex = GetSortIndex(workbook, workbookPathname, sortColumn);

            Worksheet excepWorksheet = workbook.Worksheets[1];
            Range sortRange = SetSortRange(excepWorksheet);
            numRows = sortRange.Rows.Count;
            numCols = sortRange.Columns.Count;
            sortRange.Sort(sortRange.Columns[sortIndex], XlSortOrder.xlAscending);
            workbook.SaveAs(workbookPathname, XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue,
                            XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            return workbookPathname;
        }
        //****************************************************************************************************************************
        private int GetSortIndex(Workbook workbook, string workbookPathname, ColumnName sortColumn)
        {
            int sortIndex;
            if (inputFileType == InputFileType.Nemrc)
            {
                sortIndex = GetColumnNum(sortColumn);
            }
            else
            {
                workbook.SaveAs(workbookPathname, XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue,
                                XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                OpenWithEPPlus(workbookPathname);
                sortIndex = GetColumnNum(sortColumn);
                CloseWorkbook();
            }
            return sortIndex;
        }
        //****************************************************************************************************************************
        private Range SetSortRange(Worksheet excepWorksheet)
        {
            Range usedRange = (Range)excepWorksheet.UsedRange;
            numRows = usedRange.Rows.Count;
            numCols = usedRange.Columns.Count;
            int startRow;
            if (inputFileType == InputFileType.Nemrc)
            {
                startRow = 1;
            }
            else
            {
                startRow = 2;
            }
            return excepWorksheet.Range[excepWorksheet.Cells[startRow, 1], excepWorksheet.Cells[numRows, numCols]];
        }
        //****************************************************************************************************************************
        protected Workbook OpenExcelWorkbook(string workbookPathname, bool visible)
        {
            try
            {
                ExcelClass.xlApp.DisplayAlerts = false;
                Workbook workbook = ExcelClass.xlApp.Workbooks.Open(workbookPathname, ReadOnly: false);
                if (visible)
                {
                    ExcelClass.xlApp.DisplayAlerts = true;
                    ExcelClass.xlApp.Visible = true;
                    ExcelClass.xlApp.UserControl = true;
                    ExcelClass.xlApp.WindowState = XlWindowState.xlMaximized;
                }
                else
                {
                    ExcelClass.xlApp.DisplayAlerts = false;
                    ExcelClass.xlApp.Visible = false;
                }
                return workbook;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        //****************************************************************************************************************************
        public int GetColumnNum(ColumnName columnName)
        {
            switch (inputFileType)
            {
                case InputFileType.Nemrc: return GetNemrcColumnNum(columnName);
                case InputFileType.Patriot: return GetPatriotColumnNum(columnName);
                case InputFileType.TaxMap: return GetTaxMapColumnNum(columnName);
                case InputFileType.Other: return GetOtherColumnNum(columnName);
                default: throw new Exception("GetColumnNum: " + inputFileType.ToString());
            }
        }
        //****************************************************************************************************************************
        public int GetNemrcColumnNum(ColumnName columnName)
        {
            switch (columnName)
            {
                case ColumnName.ParcelId: return 1;
                case ColumnName.ParcelSubId: return 2;
                case ColumnName.Name1: return 3;
                case ColumnName.Name2: return 4;
                case ColumnName.Addr1: return 5;
                case ColumnName.Addr2: return 6;
                case ColumnName.City: return 7;
                case ColumnName.State: return 8;
                case ColumnName.Zip: return 9;
                case ColumnName.LocationA: return 10;
                case ColumnName.LocationB: return 11;
                case ColumnName.LocationC: return 12;
                case ColumnName.StreetNum: return 13;
                case ColumnName.StreetName: return 15;
                case ColumnName.TaxMapId: return 16;
                case ColumnName.Description: return 17;
                case ColumnName.Span: return 31;
                case ColumnName.Owner: return 34;
                case ColumnName.Acres: return 40;
                case ColumnName.Value: return 42;
                case ColumnName.buildingValue: return 44;
                case ColumnName.DateHomestead: return 63;
                case ColumnName.taxStatus: return 131;
                case ColumnName.NotesContiguous: return 152;
                case ColumnName.ContiguousID: return 170;
                case ColumnName.ContiguousSubID: return 171;
                default: throw new Exception("GetNemrcColumnNum: " + columnName.ToString());
            }
        }
        //****************************************************************************************************************************
        private int GetPatriotColumnNum(ColumnName columnName)
        {
            switch (columnName)
            {
                case ColumnName.Span: return GetColumnNumber("UserAccount");
                case ColumnName.TaxMapId: return GetColumnNumber("ParcelID");
                case ColumnName.Name1: return GetColumnNumber("Owner1LastName");
                case ColumnName.Addr1: return GetColumnNumber("CuOStreet1");
                case ColumnName.Addr2: return GetColumnNumber("CuOStreet2");
                case ColumnName.City: return GetColumnNumber("CuOCity");
                case ColumnName.State: return GetColumnNumber("CuOState");
                case ColumnName.Zip: return GetColumnNumber("CuOPostal");
                case ColumnName.Acres: return GetColumnNumber("TotalLand");
                case ColumnName.Value: return GetColumnNumber("MarketAdjCost");
                case ColumnName.FirstName1: return GetColumnNumber("FirstName");
                default: throw new Exception("GetPatriotColumnNum: " + columnName.ToString());
            }
        }
        //****************************************************************************************************************************
        private int GetTaxMapColumnNum(ColumnName columnName)
        {
            switch (columnName)
            {
                case ColumnName.polygon: return GetColumnNumber("OBJECTID");
                case ColumnName.Span: return GetColumnNumber("parcels.SPAN");
                case ColumnName.InactiveSpan: return GetColumnNumber("inactive.SPAN");
                case ColumnName.InactiveParentSpan: return GetColumnNumber("inactive.PARENTSPAN");
                case ColumnName.TaxMapId: return GetColumnNumber("parcels.MAPID");
                case ColumnName.InactiveTaxMapId: return GetColumnNumber("inactive.MAPID");
                case ColumnName.Name1: return GetColumnNumber("Name1");
                case ColumnName.Name2: return GetColumnNumber("Name2");
                case ColumnName.NemrcSpan: return GetColumnNumber("GrandList$.Span");
                case ColumnName.ActiveInactive: return GetColumnNumber("GrandList$.Active");
                case ColumnName.contiguous: return GetColumnNumber("GrandList$.Contiguous");
                case ColumnName.EditNote: return GetColumnNumber("EDITNOTE");
                default: throw new Exception("Cannot Get TaxMap ColumnNum: " + columnName.ToString());
            }
        }
        //****************************************************************************************************************************
        private int GetOtherColumnNum(ColumnName columnName)
        {
            switch (columnName)
            {
                case ColumnName.FloodZoneSpan: return GetColumnNumber("FloodZoneSpan");
                default: throw new Exception("Cannot Get Other ColumnNum: " + columnName.ToString());
            }
        }
        //****************************************************************************************************************************
        private int GetColumnNumber(string columnName)
        {
            for (int i = 1; i <= numCols; i++)
            {
                string headingValue = GetCellValue(1, i);
                if (headingValue.ToLower().Contains(columnName.ToLower()))
                {
                    return i;
                }
            }
            throw new Exception("GetColumnNumber: " + columnName.ToString());
        }
        //****************************************************************************************************************************
        protected string GetCellValue(int rowIndex, int colIndex)
        {
            if (inputWorksheet.Cells[rowIndex, colIndex] == null)
            {
                return "";
            }
            if (inputWorksheet.Cells[rowIndex, colIndex].Value == null)
            {
                return "";
            }
            return inputWorksheet.Cells[rowIndex, colIndex].Value.ToString().Trim();
        }
        //****************************************************************************************************************************
    }
}
