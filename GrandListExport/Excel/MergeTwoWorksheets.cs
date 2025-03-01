using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;

namespace GrandListExport
{
    public class MergeTwoWorksheets : ExcelClass 
    {
        private WorksheetValues worksheetValues1;
        private WorksheetValues worksheetValues2;
        private bool valuesOnly = false;

        private class SalesInfo
        {
            public string CertNumber;
            public string SpanNumber;
            public SalesInfo(string CertNumber, string SpanNumber)
            {
                this.CertNumber = CertNumber;
                this.SpanNumber = SpanNumber;
            }
        }
        private ArrayList SalesData = new ArrayList();
        private int outputIndex;
        //****************************************************************************************************************************
        public MergeTwoWorksheets(System.Windows.Forms.ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
            : base(progressBar, ProgressBarLabel)
        {
        }
        //****************************************************************************************************************************
        public void MergeWorksheets(bool valuesOnly, bool workingGrandList)
        {
            try
            {
                this.valuesOnly = valuesOnly;
                //GetSalesData();
                worksheetValues1 = SelectInputFile(ExcelClass.PatriotExportsFolder, "MultipleCards", "xls", ColumnName.Span, workingGrandList);
                if (worksheetValues1 == null)
                {
                    return;
                }
                worksheetValues2 = SelectInputFile(ExcelClass.PatriotExportsFolder, "TaxMapContiguousParcels", "xls", ColumnName.Span, workingGrandList);
                if (worksheetValues2 == null)
                {
                    return;
                }
                workbookPath = ExcelClass.ReportsFolder + "\\MergedRows.xlsx";
                CreateNewExcelWorkbook("MergedRows");
                WriteHeadings();
                FindDifferencesPatriotNemrcColumn();
                worksheetValues1.CloseWorkbook();
                worksheetValues2.CloseWorkbook();
                CloseOutput();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private void WriteHeadings()
        {
            outputWorksheet.Cells[1, 1].Value = "TaxMapId";
        }
        //****************************************************************************************************************************
        private void FindDifferencesPatriotNemrcColumn()
        {
            this.ProgressBarLabel.Visible = true;
            this.progressBar.Visible = true;
            this.progressBar.Value = 0;
            this.progressBar.Step = 10;
            this.progressBar.Maximum = worksheetValues1.numRows;
            int rowIndex1 = 2;
            int rowIndex2 = 2;
            outputIndex = 2;
            while (rowIndex1 <= worksheetValues1.numRows && rowIndex2 <= worksheetValues2.numRows)
            {
                try
                {
                    string worksheet1Id = GetCellValue(worksheetValues1.inputWorksheet, rowIndex1, 8);
                    string worksheet2Id = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, 2);
                    char greaterLessthanEqual = GreaterLessthanEqual(worksheet1Id, worksheet2Id);
                    if (greaterLessthanEqual == '=')
                    {
                        ParcelIdEqual(rowIndex1, rowIndex2, worksheet1Id);
                        IncrementProgressBar();
                        rowIndex1++;
                        rowIndex2++;
                    }
                    else
                    {
                        if (greaterLessthanEqual == '<')
                        {
                            IncrementProgressBar();
                            rowIndex1++;
                        }
                        else
                        {
                            rowIndex2 = SkipDuplicateIds(rowIndex2);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        //****************************************************************************************************************************
        private void ParcelIdEqual(int rowIndex,
                                     int rowIndex2,
                                     string worksheetId)
        {
            outputIndex++;
            outputWorksheet.Cells[outputIndex, 1].Value = worksheetId;
        }
        //****************************************************************************************************************************
        private int SkipDuplicateIds(int rowIndex2)
        {
            rowIndex2++;
            return rowIndex2;
        }
        //****************************************************************************************************************************
        private char GreaterLessthanEqual(string worksheet1Id, string worksheet2Id)
        {
            worksheet1Id = ChangeDashToDotForSubParcel(worksheet1Id);
            int worksheet1Length = worksheet1Id.Length;
            int worksheet2Length = worksheet2Id.Length;
            int index = 0;
            while (index < worksheet1Length && index < worksheet2Length)
            {
                if (worksheet1Id[index] != worksheet2Id[index])
                {
                    if (worksheet1Id[index] < worksheet2Id[index])
                    {
                        return '<';
                    }
                    else
                    {
                        return '>';
                    }
                }
                index++;
            }
            if (worksheet1Length == worksheet2Length)
            {
                return '=';
            }
            if (worksheet1Length < worksheet2Length)
            {
                return '<';
            }
            return '>';
        }
        private string ChangeDashToDotForSubParcel(string worksheetId)
        {
            int indexOfDash1 = worksheetId.IndexOf('-');
            if (indexOfDash1 < 0)
            {
                return worksheetId;
            }
            int indexOfDash2 = worksheetId.IndexOf('-', indexOfDash1 + 1);
            if (indexOfDash2 < 0)
            {
                return worksheetId;
            }
            worksheetId = worksheetId.Remove(indexOfDash2, 1);
            worksheetId = worksheetId.Insert(indexOfDash2, ".");
            return worksheetId;
        }
    }
}
