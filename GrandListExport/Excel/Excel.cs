using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;
using OfficeOpenXml;

namespace GrandListExport
{
    public enum InputFileType
    {
        None = 0,
        Nemrc = 1,
        Patriot = 2,
        TaxMap = 3,
        Other = 4
    }
    public enum ColumnName
    {
        Span = 0,
        TaxMapId = 1,
        Name1 = 2,
        Name2 = 3,
        Addr1 = 4,
        Addr2 = 5,
        Acres = 6,
        Value = 7,
        StreetNum = 8,
        StreetName = 9,
        City = 10,
        State = 11,
        Zip = 12,
        FirstName1 = 13,
        taxStatus = 14,
        InactiveSpan = 15,
        InactiveParentSpan = 16,
        InactiveTaxMapId = 17,
        NemrcSpan = 18,
        ActiveInactive = 19,
        polygon = 20,
        contiguous = 21,
        ParcelId = 22,
        ParcelSubId = 23,
        NotesContiguous = 24,
        ContiguousID = 25,
        ContiguousSubID = 26,
        Description = 27,
        EditNote = 28,
        LocationA = 29,
        LocationB = 30,
        LocationC = 31,
        FloodZoneSpan = 32,
        Owner = 33,
        DateHomestead = 34,
        none = 35,
        buildingValue = 36
    }
    //****************************************************************************************************************************
    public static class Utilities
    {
        public static double ToDouble(this string dbl)
        {
            if (String.IsNullOrEmpty(dbl))
            {
                return 0.0;
            }
            try
            {
                return Convert.ToDouble(dbl);
            }
            catch (Exception)
            {
                MessageBox.Show("Invalid double: " + dbl);
                return 0.0;
            }
        }
    }
    //****************************************************************************************************************************
    public class ExcelClass : IDisposable
    {
        public static Microsoft.Office.Interop.Excel.Application xlApp;
        ExcelPackage ExcelPackageOutput;
        ExcelPackage ExcelPackageOutput2;
        //public static string NemrcExportsFolder = @"D:\Exports";
        //public static string PatriotExportsFolder = @"D:\PatriotExports";
        //public static string ReportsFolder = @"D:\Reports";
        public static string NemrcExportsFolder = @"Z:\NEMRC\Exports";
        public static string PatriotExportsFolder = @"X:\ListersInformation\PatriotExports";
        public static string ReportsFolder = @"X:\ListersInformation\Reports";
        protected static InputFileType fileType1;
        protected static InputFileType fileType2;
        protected struct ContiguousParcelId
        {
            public double acres;
            public string taxMapId;
            public string contigTaxMapId;
            public ContiguousParcelId(string taxMapId, string contigTaxMapId, double acres)
            {
                this.taxMapId = taxMapId;
                this.contigTaxMapId = contigTaxMapId;
                this.acres = acres;
            }
        }

        protected System.Windows.Forms.ProgressBar progressBar;
        protected System.Windows.Forms.Label ProgressBarLabel;
        protected string progressBarPrefix = "";
        protected int numProgress = 0;
        protected OfficeOpenXml.ExcelWorksheet outputWorksheet;
        protected OfficeOpenXml.ExcelWorksheet outputWorksheet2;
        protected string workbookPath;
        protected string workbook2Path;
        private bool _notYetDisposed = true;

        //****************************************************************************************************************************
        public ExcelClass(ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                throw new Exception("EXCEL could not be started");
            }
            this.progressBar = progressBar;
            this.ProgressBarLabel = ProgressBarLabel;
        }
        //****************************************************************************************************************************
        protected bool KnownInactiveInActiveFile(string TaxMapSpanActives)
        {
            if (TaxMapSpanActives == "32410110083")
            {
                return true;
            }
            if (TaxMapSpanActives == "32410110281")
            {
                return true;
            }
            if (TaxMapSpanActives == "32410110588")
            {
                return true;
            }
            if (TaxMapSpanActives == "32410111431")
            {
                return true;
            }
            if (TaxMapSpanActives == "32410111572")
            {
                return true;
            }
            return false;
        }
        //****************************************************************************************************************************
        protected bool KnownNemrcActiveNotInTaxMap(string TaxMapSpanActives)
        {
            if (TaxMapSpanActives == "32410111641")
            {
                return true;
            }
            return false;
        }
        //****************************************************************************************************************************
        protected bool KnownInactiveWithNoParent(string parentTaxMapId)
        {
            if (parentTaxMapId == "F-36.1")
            {
                return true;
            }
            return false;
        }
        //****************************************************************************************************************************
        protected WorksheetValues SelectInputFile(string sFolder,
                                         string titleString,
                                         string extension,
                                         ColumnName sortColumn,
                                         bool workingGrandList,
                                         int GrandListYear=0)
        {
            try
            {
                string filenameWithoutSuffix = sFolder + "\\" + titleString;
                filenameWithoutSuffix = (workingGrandList) ? filenameWithoutSuffix + "_Working" : filenameWithoutSuffix + "_AsBilled_" + GrandListYear;

                //string filenameWithoutExtension = (workingGrandList) ? filenameWithoutSuffix + "_Working" : GetMostRecentAsBilledYear(sFolder, filenameWithoutSuffix);
                //string sFilter = "Excel Files (xxx)|*.xlsx";
                //OpenFileDialog myDialog = new OpenFileDialog();
                //myDialog.Title = "Grand List " + titleString;
                //myDialog.Filter = sFilter;
                //myDialog.FilterIndex = 1;
                //myDialog.RestoreDirectory = true;
                //myDialog.InitialDirectory = sFolder;
                //if (myDialog.ShowDialog() == DialogResult.OK)
                //{
                //    return new WorksheetValues(myDialog.FileName);
                //}
                //else
                //{
                //    return null;
                //}
                if (sortColumn == ColumnName.none)
                {
                    return new WorksheetValues(filenameWithoutSuffix, extension);
                }
                else
                {
                    if (!File.Exists(filenameWithoutSuffix + "." + extension))
                    {
                        throw new Exception("File does not exist: " + filenameWithoutSuffix + "." + extension);
                    }
                    return new WorksheetValues(filenameWithoutSuffix, extension, sortColumn);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        //****************************************************************************************************************************
        protected string GetMostRecentAsBilledYear(string sFolder, string filenameWithoutSuffix)
        {
            string filenameWithoutDate = filenameWithoutSuffix + "_AsBilled_";
            string[] files = Directory.GetFiles(sFolder);
            int highestDate = 0;
            foreach (string file in files)
            {
                if (file.Contains(filenameWithoutDate))
                {
                    string extension = Path.GetExtension(file);
                    if (extension.ToLower() == ".csv" || extension.ToLower() == ".xls")
                    {
                        string dateStr = file.Substring(filenameWithoutDate.Length);
                        dateStr = Path.GetFileNameWithoutExtension(dateStr);
                        try
                        {
                            int dateInt = Convert.ToInt32(dateStr);
                            if (dateInt > highestDate)
                            {
                                highestDate = dateInt;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            if (highestDate < 2018 || highestDate > 2027)
            {
                throw new Exception("Cannot Find As Billed Date");
            }
            return filenameWithoutDate + highestDate.ToString();
        }
        //****************************************************************************************************************************
        protected int GetColumnNum(WorksheetValues worksheetValues, ColumnName columnName)
        {
            return worksheetValues.GetColumnNum(columnName);
        }
        //****************************************************************************************************************************
        protected string GetCellValue(OfficeOpenXml.ExcelWorksheet worksheet, int rowIndex, int colIndex)
        {
            if (worksheet.Cells[rowIndex, colIndex] == null)
            {
                return "";
            }
            if (worksheet.Cells[rowIndex, colIndex].Value == null)
            {
                return "";
            }
            return worksheet.Cells[rowIndex, colIndex].Value.ToString().Trim();
        }
        //****************************************************************************************************************************
        protected void AddContiguousParcelToList(ArrayList contiguousParcelList, string taxMapId, string contiguousParcelId, string acres)
        {
            if (!String.IsNullOrEmpty(contiguousParcelId))
            {
                contiguousParcelList.Add(new ContiguousParcelId(taxMapId, contiguousParcelId, acres.ToDouble()));
            }
        }
        //****************************************************************************************************************************
        protected string GetAcres(ArrayList contiguousParcelList, string taxMapId, string totalAcres)
        {
            return totalAcres;
        }
        //****************************************************************************************************************************
        protected void CreateNewExcelWorkbook(string name)
        {
            CreateExcelWorkbook(out ExcelPackageOutput, out outputWorksheet, name);
        }
        //****************************************************************************************************************************
        protected void Create2ndExcelWorkbook(string name)
        {
            CreateExcelWorkbook(out ExcelPackageOutput2, out outputWorksheet2, name);
        }
        //****************************************************************************************************************************
        protected void CreateExcelWorkbook(out ExcelPackage ExcelPackage, out ExcelWorksheet worksheet, string name)
        {
            try
            {
                ExcelPackage = new ExcelPackage();
                worksheet = ExcelPackage.Workbook.Worksheets.Add(name);
                //xlApp.DisplayAlerts = false;
                //outputWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                //worksheet = (Worksheet)outputWorkbook.Worksheets[1];
                //if (worksheet == null)
                //{
                //    Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
                //}
                //outputWorkbook.Worksheets.Add();
                //worksheet = (Worksheet)outputWorkbook.Worksheets[1];  // The added worksheet becomde worksheet 1
                //worksheet.Name = name;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        //****************************************************************************************************************************
        protected string ExtractContiguousId(string value)
        {
            int indexOf = value.IndexOf('@');
            if (indexOf > 0)
            {
                value = value.Remove(indexOf).Trim();
            }
            indexOf = value.IndexOf(' ');
            if (indexOf > 0)
            {
                value = value.Remove(indexOf).Trim();
            }
            indexOf = value.IndexOf('=');
            if (indexOf > 0)
            {
                value = value.Remove(indexOf).Trim();
            }
            return TrimLeadingZeroes(value);
        }
        //****************************************************************************************************************************
        protected string TrimLeadingZeroes(string value)
        {
            int indexOf = 0;
            while (indexOf < value.Length && value[indexOf] == '0')
            {
                indexOf++;
            }
            if (indexOf > 0)
            {
                return value.Substring(indexOf);
            }
            return value;
        }
        //****************************************************************************************************************************
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                Dispose();
                // dispose managed resources
            }
            // free native resources
        }
        //****************************************************************************************************************************
        public void Dispose()
        {
            if (_notYetDisposed)
            {
                if (ExcelPackageOutput != null)
                {
                    try
                    {
                        CloseOutput();
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Output File Must Be Open. Close and Try Again" );
                        CloseOutput();
                    }
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }
                this.ProgressBarLabel.Text = progressBarPrefix + progressBar.Maximum + " of " + this.progressBar.Maximum;
                progressBar.Value = progressBar.Maximum;
                _notYetDisposed = false;
                MessageBox.Show("Report Complete");
            }
            progressBar.Visible = false;
            ProgressBarLabel.Visible = false;
        }
        //****************************************************************************************************************************
        protected void CloseOutput()
        {
            if (outputWorksheet != null)
            {
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                outputWorksheet.View.FreezePanes(2, 1);
                ExcelPackageOutput.SaveAs(new FileInfo(workbookPath));
                outputWorksheet = null;
            }
            if (outputWorksheet2 != null)
            {
                outputWorksheet2.Cells[outputWorksheet2.Dimension.Address].AutoFitColumns();
                outputWorksheet2.View.FreezePanes(2, 1);
                ExcelPackageOutput2.SaveAs(new FileInfo(workbook2Path));
                outputWorksheet2 = null;
            }
        }
        //****************************************************************************************************************************
        protected void CloseOutputInXlsFormat()
        {
            CloseOutput();
            WorksheetValues worksheetValues = new WorksheetValues();
            worksheetValues.SaveAsXlsFile(workbookPath);
            if (!String.IsNullOrEmpty(workbook2Path))
            {
                WorksheetValues worksheetValues2 = new WorksheetValues();
                worksheetValues2.SaveAsXlsFile(workbook2Path);
            }
        }
        //****************************************************************************************************************************
        public void FormatStringCells(string cellRange)
        {
            outputWorksheet.Cells[outputWorksheet.Dimension.Address].Style.Numberformat.Format = "@";
            //Range range = outputWorksheet.Range[cellRange];
            //range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            //range.NumberFormat = "@";
        }
        //****************************************************************************************************************************
        protected void IncrementProgressBar()
        {
            int num = numProgress % 10;
            numProgress++;
            if (num == 0)
            {
                this.progressBar.PerformStep();
                this.ProgressBarLabel.Text = progressBarPrefix + numProgress + " of " + this.progressBar.Maximum;
            }
        }
        //****************************************************************************************************************************
    }
}
