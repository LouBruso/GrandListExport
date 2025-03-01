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
    public class NemrcOwnershipCheck : ExcelClass 
    {
        private WorksheetValues worksheetValues1;
        private WorksheetValues worksheetValues2;
        private ArrayList NemrcNotInPatriotRow = new ArrayList();
        private ArrayList PatriotNotInNemrcRow = new ArrayList();
        private ArrayList PatriotBlankSpanRow = new ArrayList();
        private ArrayList differentValues = new ArrayList();
        private class SpanId
        {
            public string Id;
            public string SpanNumber;
            public SpanId(string Id, string SpanNumber)
            {
                this.Id = Id;
                this.SpanNumber = SpanNumber;
            }
        }
        private ArrayList SalesData = new ArrayList();
        private int outputIndex;
        private bool activesDesired;

        private int nemrcAcresIndex;
        private int nemrcName1Index;
        private int nemrcSpanIndex;
        private int nemrcTaxMapIdIndex;
        private int nemrcTaxStatusIndex;

        protected int taxmapSpanIndex;
        protected int taxmapTaxMapIdIndex;
        protected int ActiveInactiveIndex;
        protected int PolygonIndex;
        protected int taxmapName1Index;
        protected int taxmapEditNoteIndex;
        //****************************************************************************************************************************
        public NemrcOwnershipCheck(ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
            : base(progressBar, ProgressBarLabel)
        {
        }
        //****************************************************************************************************************************
        public void NemrcTaxMapDifferences(bool activesDesired, bool workingGrandList, int grandListYear)
        {
            try
            {
                this.activesDesired = activesDesired;
                string fileLookingFor = (activesDesired) ? "NemrcActives" : "NemrcInactives";
                worksheetValues1 = SelectInputFile(ExcelClass.NemrcExportsFolder, fileLookingFor, "csv", ColumnName.Span, workingGrandList, grandListYear);
                if (worksheetValues1 == null)
                {
                    return;
                }
                fileLookingFor = (activesDesired) ? "TaxMapActives" : "TaxMapInactives";
                ColumnName columnName = (activesDesired) ? ColumnName.Span : ColumnName.InactiveSpan;
                worksheetValues2 = SelectInputFile(ExcelClass.PatriotExportsFolder, fileLookingFor, "xls", columnName, workingGrandList, grandListYear);
                if (worksheetValues2 == null)
                {
                    return;
                }
                string prefix = (activesDesired) ? "Active" : "Inactive";
                workbookPath = ReportsFolder + "\\" + prefix + "NemrcTaxMapDifferences.xlsx";
                CreateNewExcelWorkbook(prefix + " Nemrc TaxMap Differences");
                SetColumnIndexes();
                FindTaxMapDifferencs();
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
        private void SetColumnIndexes()
        {
            nemrcSpanIndex = GetColumnNum(worksheetValues1, ColumnName.Span);
            nemrcTaxMapIdIndex = GetColumnNum(worksheetValues1, ColumnName.TaxMapId);
            nemrcTaxStatusIndex = GetColumnNum(worksheetValues1, ColumnName.taxStatus);
            nemrcAcresIndex = GetColumnNum(worksheetValues1, ColumnName.Acres);
            nemrcName1Index = GetColumnNum(worksheetValues1, ColumnName.Name1);

            ColumnName spanColumnName = (activesDesired) ? ColumnName.Span : ColumnName.InactiveSpan;
            taxmapSpanIndex = GetColumnNum(worksheetValues2, spanColumnName);
            ColumnName taxMapIdColumnName = (activesDesired) ? ColumnName.TaxMapId : ColumnName.InactiveTaxMapId;
            taxmapTaxMapIdIndex = GetColumnNum(worksheetValues2, taxMapIdColumnName);
            ActiveInactiveIndex = GetColumnNum(worksheetValues2, ColumnName.ActiveInactive);
            PolygonIndex = GetColumnNum(worksheetValues2, ColumnName.polygon);
            taxmapName1Index = GetColumnNum(worksheetValues2, ColumnName.Name1);
            taxmapEditNoteIndex = GetColumnNum(worksheetValues2, ColumnName.EditNote);
        }
        //****************************************************************************************************************************
        private void WriteHeadings(bool  activesDesired)
        {
            string activeInactive = (activesDesired) ? "Actives" : "Inactives";
            outputWorksheet.Cells[1, 1].Value = "In NEMRC " + activeInactive;
            outputWorksheet.Cells[1, 2].Value = "In TaxMap " + activeInactive;
            outputWorksheet.Cells[1, 3].Value = "Name";
        }
        //****************************************************************************************************************************
        private void FindTaxMapDifferencs()
        {
            progressBar.Visible = true;
            ProgressBarLabel.Visible = true;
            progressBar.Value = 0;
            progressBar.Step = 10;
            progressBar.Maximum = (worksheetValues1.numRows > worksheetValues2.numRows) ? worksheetValues1.numRows : worksheetValues2.numRows;
            progressBarPrefix = (activesDesired) ? "Actives-" : "Inactives-";
            int rowIndex1 = 1;
            int rowIndex2 = 2;  // Skip Headings
            while (rowIndex1 <= worksheetValues1.numRows && rowIndex2 <= worksheetValues2.numRows)
            {
                try
                {
                    string NemrcSpan = GetCellValue(worksheetValues1.inputWorksheet, rowIndex1, nemrcSpanIndex);
                    string taxMapSpan = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, taxmapSpanIndex);
                    if (activesDesired)
                    {
                        NemrcSpan = NemrcSpan.Replace("-", "");
                    }
                    char greaterLessthanEqual = GreaterLessthanEqual(NemrcSpan, taxMapSpan);
                    if (greaterLessthanEqual == '=')
                    {
                        rowIndex1++;
                        rowIndex2++;
                    }
                    else if (greaterLessthanEqual == '<')
                    {
                        if (!KnownNemrcActiveNotInTaxMap(NemrcSpan))
                        {
                            AddNemrcToList(NemrcSpan, rowIndex1);
                        }
                        rowIndex1++;
                    }
                    else 
                    {
                        if (!KnownInactiveInActiveFile(taxMapSpan))
                        {
                            AddTaxMapToList(taxMapSpan, rowIndex2);
                        }
                        rowIndex2++;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                IncrementProgressBar();
            }
            while (rowIndex1 < worksheetValues1.numRows)
            {
                string NemrcSpan = GetCellValue(worksheetValues1.inputWorksheet, rowIndex1, nemrcSpanIndex);
                AddNemrcToList(NemrcSpan, rowIndex1);
                rowIndex1++;
            }
            while (rowIndex2 < worksheetValues2.numRows)
            {
                string taxMapSpan = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, taxmapSpanIndex);
                AddTaxMapToList(taxMapSpan, rowIndex2);
                rowIndex2++;
            }
            WriteHeadings(activesDesired);
            int outputIndex = 2;
            foreach (SpanId spanId in NemrcNotInPatriotRow)
            {
                outputIndex++;
                outputWorksheet.Cells[outputIndex, 1].Value = spanId.SpanNumber;
                outputWorksheet.Cells[outputIndex, 3].Value = spanId.Id;
            }
            foreach (SpanId spanId in PatriotNotInNemrcRow)
            {
                outputIndex++;
                outputWorksheet.Cells[outputIndex, 2].Value = spanId.SpanNumber;
                outputWorksheet.Cells[outputIndex, 3].Value = spanId.Id;
            }
        }
        //****************************************************************************************************************************
        private void AddNemrcToList(string NemrcSpan, int rowIndex)
        {
            double NemrcAcres = ToReal(GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcAcresIndex), NemrcSpan);
            string taxStatus = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcTaxStatusIndex);
            if (!Exclude(NemrcSpan, NemrcAcres, taxStatus))
            {
                string id = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcName1Index);
                NemrcNotInPatriotRow.Add(new SpanId(id, NemrcSpan));
            }
        }
        //****************************************************************************************************************************
        private void AddTaxMapToList(string taxMapSpan, int rowIndex2)
        {
            string taxMapActive = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, ActiveInactiveIndex);
            string taxMapId = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, taxmapTaxMapIdIndex);
            if (TownOrStateRoad(taxMapActive, taxMapId))
            {
                return;
            }
            //if (taxMapActive.ToLower() != "inactive")
            {
                if (string.IsNullOrEmpty(taxMapSpan))
                {
                    string polygonID = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, PolygonIndex);
                    PatriotNotInNemrcRow.Add(new SpanId("", "Polygon " + polygonID));
                }
                else
                {
                    string id = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, taxmapName1Index);
                    if (string.IsNullOrEmpty(id))
                    {
                        id = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, taxmapEditNoteIndex);
                    }
                    PatriotNotInNemrcRow.Add(new SpanId(id, taxMapSpan));
                }
            }
        }
        //****************************************************************************************************************************
        private bool TownOrStateRoad(string taxMapActive, string taxMapId)
        {
            if (!string.IsNullOrEmpty(taxMapActive))
            {
                return false;
            }
            if (taxMapId.ToLower().Contains("road") || taxMapId.ToLower().Contains("state"))
            {
                return true;
            }
            return false;
        }
        //****************************************************************************************************************************
        private double ToReal(string str, string span)
        {
            try
            {
                return Convert.ToDouble(str);
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to Convert Int for span " + span + ": " + str);
                return 99999;
            }
        }
        //****************************************************************************************************************************
        private bool Exclude(string span, double NemrcAcres, string taxStatus)
        {
            if (NemrcAcres == 0.0)
            {
                return true;
            }
            if (taxStatus == "S")
            {
                return true;
            }
            if (span == "324-101-10079") // Mobile Home
            {
                return true;
            }
            if (span == "324-101-10082") // Condo
            {
                return true;
            }
            if (span == "324-101-10225") // leased Lot
            {
                return true;
            }
            if (span == "324-101-10260") // Mobile Home
            {
                return true;
            }
            if (span == "324-101-10281") // Cole Pond
            {
                return true;
            }
            if (span == "324-101-10363") // leased Lot
            {
                return true;
            }
            if (span == "324-101-10556") // leased Lot
            {
                return true;
            }
            if (span == "324-101-10556") // leased Lot
            {
                return true;
            }
            if (span == "324-101-10559") // Mobile Home
            {
                return true;
            }
            if (span == "324-101-10588") // HIGHLAND FOREST LOT OWNERS 
            {
                return true;
            }
            if (span == "324-101-10653") // Town of Jamaica
            {
                return true;
            }
            if (span == "324-101-10659") // Town of Jamaica
            {
                return true;
            }
            if (span == "324-101-10737") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-10740") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-10752") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-10823") // Condo 
            {
                return true;
            }
            if (span == "324-101-10824") // Condo 
            {
                return true;
            }
            if (span == "324-101-10825") // Condo 
            {
                return true;
            }
            if (span == "324-101-10826") // Condo 
            {
                return true;
            }
            if (span == "324-101-10827") // Condo
            {
                return true;
            }
            if (span == "324-101-10828") // Condo
            {
                return true;
            }
            if (span == "324-101-10829") // Condo
            {
                return true;
            }
            if (span == "324-101-10830") // Condo
            {
                return true;
            }
            if (span == "324-101-10831") // Condo
            {
                return true;
            }
            if (span == "324-101-10832") // Condo
            {
                return true;
            }
            if (span == "324-101-10833") // Condo
            {
                return true;
            }
            if (span == "324-101-10834") // Condo
            {
                return true;
            }
            if (span == "324-101-10835") // Condo
            {
                return true;
            }
            if (span == "324-101-10867") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-10881") // Mobile Home
            {
                return true;
            }
            if (span == "324-101-10885") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-10982") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-11044") // Condo
            {
                return true;
            }
            if (span == "324-101-11045") // Condo
            {
                return true;
            }
            if (span == "324-101-11046") // Condo
            {
                return true;
            }
            if (span == "324-101-11047") // Condo
            {
                return true;
            }
            if (span == "324-101-11048") // Condo
            {
                return true;
            }
            if (span == "324-101-11134") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-11157") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-11184") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-11210") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-11262") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-11349") // VTrans
            {
                return true;
            }
            if (span == "324-101-11350") // VTrans
            {
                return true;
            }
            if (span == "324-101-11360") // VTrans
            {
                return true;
            }
            if (span == "324-101-11362") // Vtrans
            {
                return true;
            }
            if (span == "324-101-11412") // Condo
            {
                return true;
            }
            if (span == "324-101-11431") // Condo
            {
                return true;
            }
            if (span == "324-101-11556") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-11556") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-11558") // Leased Lot
            {
                return true;
            }
            if (span == "324-101-11572") // Condo
            {
                return true;
            }
            if (span == "324-101-11620") // VTrans
            {
                return true;
            }
            if (span == "324-101-11623") // Test
            {
                return true;
            }
            return false;
        }
        //****************************************************************************************************************************
        private void OutputValues(ArrayList list1, ArrayList list2, string parcelId, string span, string name1, string name2)
        {
            differentValues.Add(parcelId);
            //outputWorksheet.Cells[outputIndex, 1] = parcelId;
            outputWorksheet.Cells[outputIndex, 1].Value = name1;
            outputWorksheet.Cells[outputIndex, 2].Value = list1[0];
            outputWorksheet.Cells[outputIndex, 3].Value = list1[1];
            outputWorksheet.Cells[outputIndex, 4].Value = list1[2];
            outputWorksheet.Cells[outputIndex, 5].Value = list1[3];
            outputWorksheet.Cells[outputIndex, 6].Value = span;
            outputWorksheet.Cells[outputIndex, 7].Value = name2;
            outputWorksheet.Cells[outputIndex, 8].Value = list2[0];
            outputWorksheet.Cells[outputIndex, 9].Value = list2[1];
            outputWorksheet.Cells[outputIndex, 10].Value = list2[2];
            outputWorksheet.Cells[outputIndex, 11].Value = list2[3];
            //outputWorksheet.Cells[outputIndex, 8].Value = list[5];
            //outputWorksheet.Cells[outputIndex, 9].Value = list[6];
            //outputWorksheet.Cells[outputIndex, 10].Value = list[7];
            outputIndex++;
        }
        //****************************************************************************************************************************
        private string RemovePunctuation(string value)
        {
            value = value.Replace(";", "");
            value = value.Replace(",", "");
            value = value.Replace(".", "");
            value = value.Replace("-", "");
            return value;
        }
        //****************************************************************************************************************************
        private char GreaterLessthanEqual(string grandListValue, string ParcelsValue)
        {
            int grandListLength = grandListValue.Length;
            int parcelsLength = ParcelsValue.Length;
            int index = 0;
            while (index < grandListLength && index < parcelsLength)
            {
                if (grandListValue[index] != ParcelsValue[index])
                {
                    if (grandListValue[index] < ParcelsValue[index])
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
            if (grandListLength == parcelsLength)
            {
                return '=';
            }
            if (grandListLength < parcelsLength)
            {
                return '<';
            }
            return '>';
        }
        //****************************************************************************************************************************
    }
}
