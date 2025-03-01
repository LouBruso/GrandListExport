using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace GrandListExport
{
    class Excel : IDisposable
    {
        private class NemrcColumns
        {
            public int columnNum;
            public string column;
            public string columnName;
            public string nemrcName;
            public NemrcColumns(int columnNum, string column, string columnName, string nemrcName)
            {
                this.columnNum = columnNum;
                this.column = column;
                this.columnName = columnName;
                this.nemrcName = nemrcName;
            } 
        }

        private void CreateColumnList(ArrayList columnList)
        {
            columnList.Add(new NemrcColumns(1, "A", "PropertyID", "1st Half Parcel Id"));
            columnList.Add(new NemrcColumns(2, "B", "PropertySubID", "2nd Half Parcel Id"));
            columnList.Add(new NemrcColumns(16, "P", "TaxMapID", "Tax Map"));
            columnList.Add(new NemrcColumns(13, "M", "StreetNum, ", "911 Number"));
            columnList.Add(new NemrcColumns(15, "O", "StreetName", "911 Street"));
            columnList.Add(new NemrcColumns(3, "C", "Name1", "Owner 1"));
            columnList.Add(new NemrcColumns(4, "D", "Name2", "Owner 2"));
            columnList.Add(new NemrcColumns(5, "E", "AddressA", "Address 1"));
            columnList.Add(new NemrcColumns(6, "F", "AddressB", "Address 2"));
            columnList.Add(new NemrcColumns(7, "G", "City", "City"));
            columnList.Add(new NemrcColumns(8, "H", "State", "State"));
            columnList.Add(new NemrcColumns(9, "I", "Zip", "Zip"));
            columnList.Add(new NemrcColumns(10, "J", "LocationA", "Location A"));
            columnList.Add(new NemrcColumns(11, "K", "LocationB", "Location B"));
            columnList.Add(new NemrcColumns(12, "L", "LocationC", "Location C"));
            columnList.Add(new NemrcColumns(17, "Q", "Description", "Property Desc"));
        }

        private bool _notYetDisposed = true;
        private Workbook inputWorkbook = null;
        private Workbook outputWorkbook = null;
        private string workbookPath;
        private Worksheet inputWorksheet;
        private Worksheet outputWorksheet;
        private int rowNumber;
        private System.Windows.Forms.ProgressBar progressBar;

        public Excel(System.Windows.Forms.ProgressBar progressBar)
        {
            this.progressBar = progressBar;
        }
        public void CompareGrandListWorksheets()
        {
            string filename = SelectInputFile();
            if (!string.IsNullOrEmpty(filename))
            {
                ArrayList TaxMapIds = CreateListOfTaxMapIds(filename);
                TaxMapIds.Sort();
                int index = TaxMapIds.BinarySearch("S-65");
                filename = SelectInputFile();
                if (!string.IsNullOrEmpty(filename))
                {
                    CheckToSeeIfTaxMapIdInList(filename);
                }
            }
        }
        public void CheckToSeeIfTaxMapIdInList(string filename)
        {
        }
        public ArrayList CreateListOfTaxMapIds(string filename)
        {
            try
            {
                inputWorkbook = OpenExcelWorkbook(filename, false);
                inputWorksheet = inputWorkbook.Worksheets[1];
                progressBar.Visible = true;
                progressBar.Value = 0;
                progressBar.Step = 1;
                int numCells = GetNumCells();
                progressBar.Maximum = numCells;
                ArrayList TaxMapIds = new ArrayList();
                for (int rowIndex = 2; rowIndex <= numCells; rowIndex++)
                {
                    TaxMapIds.Add(inputWorksheet.Cells[rowIndex, 3].Value);
                    progressBar.PerformStep();
                }
                progressBar.Visible = false;
                inputWorkbook.Close();
                inputWorkbook = null;
                return TaxMapIds;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw new Exception(ex.Message);
            }
        }
        public void SelectiveColumns()
        {
            string filename = SelectInputFile();
            if (!string.IsNullOrEmpty(filename))
            {
                inputWorkbook = OpenExcelWorkbook(filename, false);
                workbookPath = filename.Replace("Active.xlsx", "GrandList.xlsx"); 
                CreateNewExcelWorkbook();
                ArrayList columnList = new ArrayList();
                CreateColumnList(columnList);
                WriteHeadings(columnList);
                var numRows = 0;
                numRows = CopySelectedColumns(columnList, numRows);
                inputWorkbook.Close();
                filename = filename.Replace("Active", "Inactive");
                inputWorkbook = OpenExcelWorkbook(filename, false);
                CopySelectedColumns(columnList, numRows);
            }
        }
        public string SelectInputFile()
        {
            try
            {
                string sFilter = "Excel Files (xlsx)|*.xlsx";
                string sFolder = "c:\\Music\\MusicExcel";
                OpenFileDialog myDialog = new OpenFileDialog();
                myDialog.Title = "Grand List";
                myDialog.Filter = sFilter;
                myDialog.FilterIndex = 1;
                myDialog.RestoreDirectory = true;
                myDialog.InitialDirectory = sFolder;
                if (myDialog.ShowDialog() == DialogResult.OK)
                {
                    return myDialog.FileName;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw new Exception(ex.Message);
            }
        }
        private int CopySelectedColumns(ArrayList columnList, int cumulativeRowIndex)
        {
            inputWorksheet = inputWorkbook.Worksheets[1];
            progressBar.Visible = true;
            progressBar.Value = 0;
            progressBar.Step = 1;
            int numCells = GetNumCells();
            progressBar.Maximum = numCells;
            var numRows = 0;
            for (int rowIndex = 1; rowIndex <= numCells; rowIndex++)
            {
                numRows++;
                int columnNum = 0;
                foreach (NemrcColumns nemrcColumn in columnList)
                {
                    columnNum++;
                    outputWorksheet.Cells[cumulativeRowIndex + rowIndex + 1, columnNum] = inputWorksheet.Cells[rowIndex, nemrcColumn.columnNum].Value;
                }
                progressBar.PerformStep();
            }
            progressBar.Visible = false;
            return numRows;
        }
        private int GetNumCells()
        {
            var last = inputWorksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            var usedRange = inputWorksheet.UsedRange;
            return last.Row;
        }
        //****************************************************************************************************************************
        private Workbook OpenExcelWorkbook(string workbookPathname, bool visible)
        {
            try
            {
                ExcelApplication.xlApp.DisplayAlerts = false;
                Workbook workbook = ExcelApplication.xlApp.Workbooks.Open(workbookPathname, ReadOnly: false);
                if (visible)
                {
                    ExcelApplication.xlApp.DisplayAlerts = true;
                    ExcelApplication.xlApp.Visible = true;
                    ExcelApplication.xlApp.UserControl = true; 
                    ExcelApplication.xlApp.WindowState = XlWindowState.xlMaximized;
/*                  Microsoft.Office.Interop.Excel.Range firstRow = (Microsoft.Office.Interop.Excel.Range)inputWorksheet.Rows[1];
                    firstRow.Activate();
                    firstRow.Select();
                    firstRow.Application.ActiveWindow.FreezePanes = true;
                    firstRow.Activate();
                    firstRow.Select();
                    firstRow.AutoFilter(1, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);*/
                }
                else
                {
                    ExcelApplication.xlApp.DisplayAlerts = false;
                    ExcelApplication.xlApp.Visible = false;
                }
                return workbook;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void CreateNewExcelWorkbook()
        {
            try
            {
                ExcelApplication.xlApp.DisplayAlerts = false;
                outputWorkbook = ExcelApplication.xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                outputWorksheet = (Worksheet)outputWorkbook.Worksheets[1];
                if (outputWorksheet == null)
                {
                    Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
                }
                outputWorkbook.Worksheets.Add();
                outputWorksheet = (Worksheet)outputWorkbook.Worksheets[1];  // The added worksheet becomde worksheet 1
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void WriteHeadings(ArrayList columnList)
        {
            int columnNum = 0; 
            foreach (NemrcColumns nemrcColumn in columnList)
            {
                columnNum++;
                outputWorksheet.Cells[1, columnNum] = nemrcColumn.columnName;
            }
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                Dispose();
                // dispose managed resources
            }
            // free native resources
        }

        public void Dispose()
        {
            if (_notYetDisposed)
            {
                if (inputWorkbook != null)
                {
                    inputWorkbook.Close();
                    inputWorkbook = null;
                }
                if (outputWorkbook != null)
                {
                    var usedRange = (Microsoft.Office.Interop.Excel.Range)outputWorksheet.UsedRange;
                    usedRange.Columns.AutoFit(); // Autofit before hidding
                    Microsoft.Office.Interop.Excel.Range firstRow = (Microsoft.Office.Interop.Excel.Range)outputWorksheet.Rows[1];
                    outputWorksheet.Application.ActiveWindow.SplitRow = 1;
                    firstRow.Application.ActiveWindow.FreezePanes = true;
                    outputWorkbook.SaveAs(workbookPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    outputWorkbook.Close();
                    outputWorkbook = null;
                }
                _notYetDisposed = false;
            }
        }

    }
}
