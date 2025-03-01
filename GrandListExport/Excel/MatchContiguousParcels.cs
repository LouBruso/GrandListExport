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
    public class MatchContiguousParcels : ExcelClass 
    {
        private WorksheetValues worksheetValues1;
        private WorksheetValues worksheetValues2;

        private ArrayList columnList = new ArrayList();
        private ArrayList SkipedParcels = new ArrayList();

        class TaxMapIdWithSpan
        {
            public string span;
            public string parentTaxMapId;
            public string parentSpan;
            public string name1;
            public string name2;
            public TaxMapIdWithSpan(string parentTaxMapId, string span, string parentSpan, string name1, string name2)
            {
                this.name1 = name1;
                this.name2 = name2;
                this.parentTaxMapId = parentTaxMapId;
                this.span = span;
                this.parentSpan = parentSpan;
            }
            public TaxMapIdWithSpan(string span, string name1, string name2)
            {
                this.name1 = name1;
                this.name2 = name2;
                this.span = span;
                this.parentTaxMapId = "";
                this.parentSpan = "";
            }
        }

        ArrayList contiguousTaxMapIds;
        ArrayList activeTaxMapIdsWithSpan;

        private int taxmapActivesSpanIndex;
        private int taxmapActivesTaxMapIdIndex;
        private int taxmapActivesNemrcSpanIndex;

        private int taxmapInactivesSpanIndex;
        private int taxmapInactivesParentSpanIndex;
        private int taxmapInactivesTaxMapIdIndex;
        private int taxmapInactivesParentTaxMapIdIndex;
        private int taxmapInactivesNemrcSpanIndex;
        private int taxmapInactivesname1Index;
        private int taxmapInactivesname2Index;
        private int taxmapActivesname1Index;
        private int taxmapActivesname2Index;
        private int taxmapActivesActiveInactive;

        //****************************************************************************************************************************
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
        private int cumulativeRowIndex;
        //****************************************************************************************************************************
        public MatchContiguousParcels(ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
            : base(progressBar, ProgressBarLabel)
        {
        }
        //****************************************************************************************************************************
        public void MatchContiguous(bool workingGrandList, int grandListYear)
        {
            try
            {
                try
                {
                    worksheetValues1 = SelectInputFile(ExcelClass.PatriotExportsFolder, "TaxMapActives", "xls", ColumnName.Span, workingGrandList, grandListYear);
                    if (worksheetValues1 == null)
                    {
                        return;
                    }
                    worksheetValues2 = SelectInputFile(ExcelClass.PatriotExportsFolder, "TaxMapInactives", "xls", ColumnName.InactiveSpan, workingGrandList, grandListYear);
                    if (worksheetValues2 == null)
                    {
                        return;
                    }

                    SetColumnIndexes();
                    progressBar.Visible = true;
                    progressBar.Step = 10;
                    ProgressBarLabel.Visible = true;
                    GetContiguousParcels(worksheetValues2);
                    workbookPath = ExcelClass.ReportsFolder + "\\TaxMapMatchInactiveContiguousIds.xlsx";
                    CreateNewExcelWorkbook("TaxMapMatchInactiveContiguousIds");
                    cumulativeRowIndex = 0;
                    activeTaxMapIdsWithSpan = new ArrayList();

                    cumulativeRowIndex = 0;
                    CheckActives();
                    progressBarPrefix = "Inactives-";
                    CheckInactives();
                    worksheetValues1.CloseWorkbook();
                    worksheetValues2.CloseWorkbook();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private void SetColumnIndexes()
        {
            taxmapActivesSpanIndex = GetColumnNum(worksheetValues1, ColumnName.Span);
            taxmapActivesTaxMapIdIndex = GetColumnNum(worksheetValues1, ColumnName.TaxMapId);
            taxmapActivesNemrcSpanIndex = GetColumnNum(worksheetValues1, ColumnName.NemrcSpan);
            taxmapActivesname1Index = GetColumnNum(worksheetValues1, ColumnName.Name1);
            taxmapActivesname2Index = GetColumnNum(worksheetValues1, ColumnName.Name2);
            taxmapActivesActiveInactive = GetColumnNum(worksheetValues1, ColumnName.ActiveInactive);

            taxmapInactivesSpanIndex = GetColumnNum(worksheetValues2, ColumnName.InactiveSpan);
            taxmapInactivesParentSpanIndex = GetColumnNum(worksheetValues2, ColumnName.InactiveParentSpan);
            taxmapInactivesTaxMapIdIndex = GetColumnNum(worksheetValues2, ColumnName.InactiveTaxMapId);
            taxmapInactivesParentTaxMapIdIndex = GetColumnNum(worksheetValues2, ColumnName.contiguous);
            taxmapInactivesNemrcSpanIndex = GetColumnNum(worksheetValues2, ColumnName.NemrcSpan);
            taxmapInactivesname1Index = GetColumnNum(worksheetValues2, ColumnName.Name1);
            taxmapInactivesname2Index = GetColumnNum(worksheetValues2, ColumnName.Name2);
        }
        //****************************************************************************************************************************
        private void CheckActives()
        {
            try
            {
                progressBarPrefix = "Actives-";
                progressBar.Value = 0;
                progressBar.Maximum = worksheetValues1.numRows;
                cumulativeRowIndex++;
                outputWorksheet.Cells[cumulativeRowIndex, 1].Value = "Actives";
                outputWorksheet.Cells[cumulativeRowIndex, 2].Value = "TaxMapSpanActives";
                outputWorksheet.Cells[cumulativeRowIndex, 3].Value = "TaxMapSpanInActives";
                outputWorksheet.Cells[cumulativeRowIndex, 4].Value = "Name1";
                int rowIndex2 = 1;
                while (rowIndex2 <= worksheetValues1.numRows)
                {
                    rowIndex2++;
                    string TaxMapSpanActives = GetCellValue(worksheetValues1.inputWorksheet, rowIndex2, taxmapActivesSpanIndex);
                    string NemrcSpanActives = GetCellValue(worksheetValues1.inputWorksheet, rowIndex2, taxmapActivesNemrcSpanIndex);
                    NemrcSpanActives = NemrcSpanActives.Replace("-", "");
                    string ActivesActiveInactive = GetCellValue(worksheetValues1.inputWorksheet, rowIndex2, taxmapActivesActiveInactive);
                    if (TaxMapSpanActives.Contains("11242"))
                    {
                    }
                    string activeTaxMapName1 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex2, taxmapActivesname1Index);
                    string activeTaxMapName2 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex2, taxmapActivesname2Index);
                    TaxMapIdWithSpan activeTaxMapIdWithSpan = new TaxMapIdWithSpan(TaxMapSpanActives, activeTaxMapName1, activeTaxMapName2);
//                    xxxx
                    if (TaxMapSpanActives != NemrcSpanActives)
                    {
                        cumulativeRowIndex++;
                        outputWorksheet.Cells[cumulativeRowIndex, 1].Value = rowIndex2;
                        outputWorksheet.Cells[cumulativeRowIndex, 2].Value = TaxMapSpanActives;
                        outputWorksheet.Cells[cumulativeRowIndex, 3].Value = NemrcSpanActives;
                        outputWorksheet.Cells[cumulativeRowIndex, 4].Value = activeTaxMapName1;
                    }
                    else if (ActivesActiveInactive.ToLower() == "inactive" && !KnownInactiveInActiveFile(TaxMapSpanActives))
                    {
                        cumulativeRowIndex++;
                        outputWorksheet.Cells[cumulativeRowIndex, 1].Value = rowIndex2;
                        outputWorksheet.Cells[cumulativeRowIndex, 2].Value = TaxMapSpanActives;
                        outputWorksheet.Cells[cumulativeRowIndex, 3].Value = ActivesActiveInactive;
                        outputWorksheet.Cells[cumulativeRowIndex, 4].Value = activeTaxMapName1;
                    }
                    string taxMapId = GetCellValue(worksheetValues1.inputWorksheet, rowIndex2, taxmapActivesTaxMapIdIndex);
                    string contiguousTaxMapId = GetCellValue(worksheetValues1.inputWorksheet, rowIndex2, taxmapInactivesParentTaxMapIdIndex);
                    if (TaxMapIdFound(contiguousTaxMapIds, taxMapId, TaxMapSpanActives, true, activeTaxMapIdWithSpan))
                    {
                        activeTaxMapIdsWithSpan.Add(new TaxMapIdWithSpan(taxMapId, NemrcSpanActives, TaxMapSpanActives, activeTaxMapName1, activeTaxMapName2));
                    }
                    IncrementProgressBar();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private void CheckInactives()
        {
            try
            {
                progressBarPrefix = "Inactives-";
                progressBar.Value = 0;
                progressBar.Maximum = worksheetValues2.numRows;
                cumulativeRowIndex++;
                cumulativeRowIndex++;
                outputWorksheet.Cells[cumulativeRowIndex, 1].Value = "Inactives";

                int rowIndex = 1;
                while (rowIndex <= worksheetValues2.numRows)
                {
                    rowIndex++;
                    string parentTaxMapId = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesParentTaxMapIdIndex);
                    string TaxMapSpanParent = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesParentSpanIndex);
                    string TaxMapSpanInactives = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesSpanIndex);
                    string TaxMapSpanActives = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesNemrcSpanIndex);
                    if (TaxMapSpanInactives == "324-101-11557")
                    {
                    }
                    if (TaxMapSpanInactives != TaxMapSpanActives)
                    {
                        cumulativeRowIndex++;
                        string taxMapName1 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesname1Index);
                        string taxMapName2 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesname2Index);
                        outputWorksheet.Cells[cumulativeRowIndex, 1].Value = rowIndex;
                        outputWorksheet.Cells[cumulativeRowIndex, 2].Value = TaxMapSpanActives;
                        outputWorksheet.Cells[cumulativeRowIndex, 3].Value = TaxMapSpanInactives;
                        outputWorksheet.Cells[cumulativeRowIndex, 4].Value = taxMapName1;
                        outputWorksheet.Cells[cumulativeRowIndex, 5].Value = taxMapName2;
                    }
                    if (!String.IsNullOrEmpty(parentTaxMapId))
                    {
                        string inactiveTaxMapName1 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesname1Index);
                        string inactiveTaxMapName2 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesname2Index);
                        if (!TaxMapIdFound(activeTaxMapIdsWithSpan, parentTaxMapId, TaxMapSpanParent, false, null))
                        {
                            cumulativeRowIndex++;
                            string taxMapId = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesTaxMapIdIndex);
                            outputWorksheet.Cells[cumulativeRowIndex, 1].Value = rowIndex;
                            outputWorksheet.Cells[cumulativeRowIndex, 2].Value = taxMapId;
                            outputWorksheet.Cells[cumulativeRowIndex, 3].Value = parentTaxMapId;
                            outputWorksheet.Cells[cumulativeRowIndex, 4].Value = inactiveTaxMapName1;
                            outputWorksheet.Cells[cumulativeRowIndex, 5].Value = inactiveTaxMapName2;
                        }
                    }
                    IncrementProgressBar();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private bool TaxMapIdFound(ArrayList taxMapIdsWithSpan, string parentTaxMapId, string TaxMapSpanParent, bool checkInactives, TaxMapIdWithSpan activeTaxMapIdWithSpan)
        {
            bool found = false;
            foreach (TaxMapIdWithSpan taxMapIdWithSpan in taxMapIdsWithSpan)
            {
                string arrayParentTaxMapId = taxMapIdWithSpan.parentTaxMapId;
                string arrayParentSpan = taxMapIdWithSpan.parentSpan;
                if (arrayParentSpan == TaxMapSpanParent)
                {
                    found = true;
                }
                if (arrayParentTaxMapId == parentTaxMapId)
                {
                    found = true;
                    if (checkInactives && PropertiesAreDifferent(taxMapIdWithSpan, activeTaxMapIdWithSpan))
                    {
                        cumulativeRowIndex++;
                        outputWorksheet.Cells[cumulativeRowIndex, 1].Value = taxMapIdWithSpan.span;
                        outputWorksheet.Cells[cumulativeRowIndex, 2].Value = taxMapIdWithSpan.parentSpan;
                        outputWorksheet.Cells[cumulativeRowIndex, 3].Value = activeTaxMapIdWithSpan.name1;
                        outputWorksheet.Cells[cumulativeRowIndex, 4].Value = taxMapIdWithSpan.name1;
                        outputWorksheet.Cells[cumulativeRowIndex, 5].Value = activeTaxMapIdWithSpan.name2;
                        outputWorksheet.Cells[cumulativeRowIndex, 6].Value = taxMapIdWithSpan.name2;
                    }
                }
            }
            if (KnownInactiveWithNoParent(parentTaxMapId))
            {
                found = true;
            }
            return found;
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
            return false;
        }
        //****************************************************************************************************************************
        private void GetContiguousParcels(WorksheetValues worksheetValues)
        {
            progressBarPrefix = "Contiguous Table-";
            progressBar.Value = 0;
            progressBar.Maximum = worksheetValues.numRows;
            contiguousTaxMapIds = new ArrayList();
            int rowIndex = 1;
            while (rowIndex <= worksheetValues.numRows)
            {
                rowIndex++;
                string taxMapId = GetCellValue(worksheetValues.inputWorksheet, rowIndex, taxmapInactivesTaxMapIdIndex);
                string TaxMapSpanInactives = GetCellValue(worksheetValues.inputWorksheet, rowIndex, taxmapInactivesSpanIndex);
                string TaxMapSpanParent = GetCellValue(worksheetValues.inputWorksheet, rowIndex, taxmapInactivesParentSpanIndex);
                string parentTaxMapId = GetCellValue(worksheetValues.inputWorksheet, rowIndex, taxmapInactivesParentTaxMapIdIndex);
                if (!String.IsNullOrEmpty(taxMapId))
                {
                    string span = GetCellValue(worksheetValues.inputWorksheet, rowIndex, taxmapInactivesSpanIndex);
                    string inactiveTaxMapName1 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesname1Index);
                    string inactiveTaxMapName2 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex, taxmapInactivesname2Index);
                    TaxMapIdWithSpan taxMapIdWithSpan = new TaxMapIdWithSpan(parentTaxMapId, span, TaxMapSpanParent, inactiveTaxMapName1, inactiveTaxMapName2);
                    contiguousTaxMapIds.Add(taxMapIdWithSpan);
                }
                IncrementProgressBar();
            }
        }
        //****************************************************************************************************************************
    }
}
