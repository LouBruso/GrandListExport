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
    public class CombineActivesInactives : ExcelClass 
    {
        class TaxMapIdWithSpan
        {
            public string taxMapId;
            public string name1;
            public string name2;
            public string address1;
            public string address2;
            public string zip;
            public TaxMapIdWithSpan(string taxMapId, string name1, string name2, string address1, string address2, string zip)
            {
                this.name1 = name1;
                this.name2 = name2;
                this.taxMapId = taxMapId;
                this.address1 = address1;
                this.address2 = address2;
                this.zip = zip;
            }
        }
        private WorksheetValues worksheetValues1;
        private WorksheetValues worksheetValues2;

        int nemrcSpanIndex;
        int nemrcTaxMapIdIndex;
        int nemrcName1Index;
        int nemrcName2Index;
        int nemrcAddress1Index;
        int nemrcAddress2Index;
        int nemrcZipIndex;
        int nemrcTaxStatus;
        int nemrcStreetNum;
        int nemrcStreetName;
        int nemrcDescription;
        int nemrcAcres;
        int nemrcSurvey;
        int nemrcContiguous;
        int nemrcNotesContiguous;

        private ArrayList ContiguousActiveParcels = new ArrayList();
        private class NemrcColumns
        {
            public int columnNum;
            public string columnName;

            public NemrcColumns(int columnNum, string columnName)
            {
                this.columnNum = columnNum;
                this.columnName = columnName;
            }
        }
        ArrayList contiguousTaxMapIds = new ArrayList();
        ArrayList columnList;
        private int cumulativeRowIndex;
        private int cumulativeRowIndex2;
        //****************************************************************************************************************************
        public CombineActivesInactives(ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
            : base(progressBar, ProgressBarLabel)
        {
        }
        //****************************************************************************************************************************
        public void Combine(bool workingGrandList, int grandListYear)
        {
            try
            {
                ExcelClass.NemrcExportsFolder = ExcelClass.NemrcExportsFolder.Replace("Z:", "C:");
                string sGrandListYear = (grandListYear == 0) ? "" : grandListYear.ToString();
                worksheetValues1 = SelectInputFile(ExcelClass.NemrcExportsFolder, "NemrcActives", "xls", ColumnName.Span, workingGrandList, grandListYear);
                if (worksheetValues1 == null)
                {
                    return;
                }
                worksheetValues2 = SelectInputFile(ExcelClass.NemrcExportsFolder, "NemrcInactives", "xls", ColumnName.Span, workingGrandList, grandListYear);
                if (worksheetValues2 == null)
                {
                    return;
                }
                SetColumnIndexes();
                CreateColumnList(worksheetValues1);

                CreateNewExcelWorkbook("NEMRC_GrandList");
                Create2ndExcelWorkbook("ContiguousDifferences");
                //GetContiguousParcels(worksheetValues2);
                WriteHeadings();
                workbookPath = ReportsFolder + "\\NEMRC_GrandList.xls";
                workbook2Path = ReportsFolder + "\\ContiguousDifferences.xls";
                cumulativeRowIndex = 0;
                cumulativeRowIndex2 = 1;
                progressBarPrefix = "Inactives-";
                cumulativeRowIndex = CopySelectedColumns(worksheetValues2, "Inactive");
                progressBarPrefix = "Actives-";
                CopySelectedColumns(worksheetValues1, "Active");
                ProgressBarLabel.Visible = false;
                worksheetValues1.CloseWorkbook();
                worksheetValues2.CloseWorkbook();
                CloseOutputInXlsFormat();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private TaxMapIdWithSpan AddToTaxMapIdWithSpan(WorksheetValues worksheetValues, int rowIndex, string taxMapId)
        {
            if (!String.IsNullOrEmpty(taxMapId))
            {
                string name1 = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcName1Index);
                string name2 = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcName2Index);
                string address1 = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcAddress1Index);
                string address2 = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcAddress2Index);
                string zip = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcZipIndex);
                return new TaxMapIdWithSpan(taxMapId, name1, name2, address1, address2, zip);
            }
            return null;
        }
        //****************************************************************************************************************************
        private int CopySelectedColumns(WorksheetValues worksheetValues, string activeInactive)
        {
            var numRows = 0;
            try
            {
                OfficeOpenXml.ExcelWorksheet inputWorksheet = worksheetValues.inputWorksheet;
                ProgressBarLabel.Visible = true;
                progressBar.Visible = true;
                progressBar.Value = 0;
                progressBar.Step = 10;
                progressBar.Maximum = worksheetValues.numRows;
                for (int rowIndex = 1; rowIndex <= worksheetValues.numRows; rowIndex++)
                {
                    try
                    {
                        numRows++;
                        if (numRows == 139)
                        {

                        }
                        int columnNum = 0;
                        string contiguousTaxMapId = "";
                        if (activeInactive == "Active")
                        {
                            string taxMapId = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcTaxMapIdIndex);
                            if (!String.IsNullOrEmpty(taxMapId))
                            {
                                TaxMapIdWithSpan activeTaxMapId = AddToTaxMapIdWithSpan(worksheetValues, rowIndex, taxMapId);
                                CheckContiguousInactives(contiguousTaxMapIds, activeTaxMapId);
                            }
                        }
                        else
                        {
                            string parcelId = TrimLeadingZeroes(GetCellValue(inputWorksheet, rowIndex, nemrcContiguous).Trim());
                            string subParcelId = GetCellValue(inputWorksheet, rowIndex, nemrcContiguous + 1);
                            contiguousTaxMapId = parcelId.Trim();
                            if (!String.IsNullOrEmpty(subParcelId))
                            {
                                contiguousTaxMapId += "." + subParcelId.Trim();
                            }
                            if (!String.IsNullOrEmpty(contiguousTaxMapId))
                            {
                                contiguousTaxMapIds.Add(AddToTaxMapIdWithSpan(worksheetValues, rowIndex, contiguousTaxMapId));
                            }
                        }
                        foreach (NemrcColumns nemrcColumn in columnList)
                        {
                            columnNum++;
                            if (nemrcColumn.columnName == "Contiguous")
                            {
                                outputWorksheet.Cells[cumulativeRowIndex + rowIndex + 1, columnNum].Value = contiguousTaxMapId;
                            }
                            else
                            {
                                string value = GetCellValue(inputWorksheet, rowIndex, nemrcColumn.columnNum);
                                outputWorksheet.Cells[cumulativeRowIndex + rowIndex + 1, columnNum].Value = value;
                                if (columnNum == 1)
                                {
                                    columnNum++;
                                    outputWorksheet.Cells[cumulativeRowIndex + rowIndex + 1, columnNum].Value = activeInactive;
                                }
                            }
                        }
                        IncrementProgressBar();
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return numRows;
        }
        //****************************************************************************************************************************
        private void CheckContiguousInactives(ArrayList taxMapIdsWithSpan, TaxMapIdWithSpan activeTaxMapId)
        {
            foreach (TaxMapIdWithSpan taxMapIdWithSpan in taxMapIdsWithSpan)
            {
                if (taxMapIdWithSpan.taxMapId == activeTaxMapId.taxMapId)
                {
                    if (PropertiesAreDifferent(taxMapIdWithSpan, activeTaxMapId))
                    {
                        cumulativeRowIndex2++;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 1].Value = taxMapIdWithSpan.taxMapId;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 2].Value = activeTaxMapId.name1;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 3].Value = taxMapIdWithSpan.name1;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 4].Value = activeTaxMapId.name2;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 5].Value = taxMapIdWithSpan.name2;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 6].Value = activeTaxMapId.address1;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 7].Value = taxMapIdWithSpan.address1;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 8].Value = activeTaxMapId.address2;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 9].Value = taxMapIdWithSpan.address2;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 10].Value = activeTaxMapId.zip;
                        outputWorksheet2.Cells[cumulativeRowIndex2, 11].Value = taxMapIdWithSpan.zip;
                    }
                }
            }
        }
        //****************************************************************************************************************************
        private bool PropertiesAreDifferent(TaxMapIdWithSpan inactiveTaxMapIdWithSpan, TaxMapIdWithSpan activeTaxMapIdWithSpan)
        {
            if (activeTaxMapIdWithSpan.name1 != inactiveTaxMapIdWithSpan.name1)
            {
                return true;
            }
            if (activeTaxMapIdWithSpan.name2 != inactiveTaxMapIdWithSpan.name2)
            {
                return true;
            }
            if (activeTaxMapIdWithSpan.address1 != inactiveTaxMapIdWithSpan.address1)
            {
                return true;
            }
            if (activeTaxMapIdWithSpan.address2 != inactiveTaxMapIdWithSpan.address2)
            {
                return true;
            }
            if (activeTaxMapIdWithSpan.zip != inactiveTaxMapIdWithSpan.zip)
            {
                return true;
            }
            return false;
        }
        //****************************************************************************************************************************
        private void SetColumnIndexes()
        {
            nemrcSpanIndex = GetColumnNum(worksheetValues1, ColumnName.Span);
            nemrcTaxMapIdIndex = GetColumnNum(worksheetValues1, ColumnName.TaxMapId);
            nemrcName1Index = GetColumnNum(worksheetValues1, ColumnName.Name1);
            nemrcName2Index = GetColumnNum(worksheetValues1, ColumnName.Name2);
            nemrcAddress1Index = GetColumnNum(worksheetValues1, ColumnName.Addr1);
            nemrcAddress2Index = GetColumnNum(worksheetValues1, ColumnName.Addr2);
            nemrcZipIndex = GetColumnNum(worksheetValues1, ColumnName.Zip);
            nemrcTaxStatus = GetColumnNum(worksheetValues1, ColumnName.taxStatus);
            nemrcStreetNum = GetColumnNum(worksheetValues1, ColumnName.StreetNum);
            nemrcStreetName = GetColumnNum(worksheetValues1, ColumnName.StreetName);
            nemrcDescription = GetColumnNum(worksheetValues1, ColumnName.Description);
            nemrcAcres = GetColumnNum(worksheetValues1, ColumnName.Acres);
            nemrcSurvey = GetColumnNum(worksheetValues1, ColumnName.LocationB);
            nemrcContiguous = GetColumnNum(worksheetValues1, ColumnName.ContiguousID);
            nemrcNotesContiguous = GetColumnNum(worksheetValues1, ColumnName.NotesContiguous);

        }
        //****************************************************************************************************************************
        private void CreateColumnList(WorksheetValues worksheetValues)
        {
            columnList = new ArrayList();
            columnList.Add(new NemrcColumns(nemrcTaxMapIdIndex, "TaxMapID"));
            columnList.Add(new NemrcColumns(nemrcTaxStatus, "TaxStatus"));
            columnList.Add(new NemrcColumns(nemrcSpanIndex, "Span"));
            columnList.Add(new NemrcColumns(nemrcStreetNum, "StreetNum"));
            columnList.Add(new NemrcColumns(nemrcStreetName, "StreetName"));
            columnList.Add(new NemrcColumns(nemrcName1Index, "Name1"));
            columnList.Add(new NemrcColumns(nemrcName2Index, "Name2"));
            columnList.Add(new NemrcColumns(nemrcDescription, "Description"));
            columnList.Add(new NemrcColumns(nemrcAcres, "Acres"));
            columnList.Add(new NemrcColumns(nemrcSurvey, "Survey"));
            columnList.Add(new NemrcColumns(nemrcContiguous, "Contiguous"));
        }
        //****************************************************************************************************************************
        public void WriteHeadings()
        {
            int columnNum = 0;
            foreach (NemrcColumns nemrcColumn in columnList)
            {
                columnNum++;
                outputWorksheet.Cells[1, columnNum].Value = nemrcColumn.columnName;
                if (columnNum == 1)
                {
                    columnNum++;
                    outputWorksheet.Cells[1, columnNum].Value = "Active";
                }
            }
            outputWorksheet2.Cells[1, 1].Value = "TaxMapId";
            outputWorksheet2.Cells[1, 2].Value = "Active Name1";
            outputWorksheet2.Cells[1, 3].Value = "Inactive Name1";
            outputWorksheet2.Cells[1, 4].Value = "Active Name2";
            outputWorksheet2.Cells[1, 5].Value = "Inactive Name2";
            outputWorksheet2.Cells[1, 6].Value = "Active Address1";
            outputWorksheet2.Cells[1, 7].Value = "Inactive Address1";
            outputWorksheet2.Cells[1, 8].Value = "Active Address2";
            outputWorksheet2.Cells[1, 9].Value = "Inactive Address2";
            outputWorksheet2.Cells[1, 10].Value = "Active Zip";
            outputWorksheet2.Cells[1, 11].Value = "Inactive Zip";
        }
        //****************************************************************************************************************************
    }
}
