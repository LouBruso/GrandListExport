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
    public class CreateListOfParcelIds : ExcelClass 
    {
        private WorksheetValues worksheetValues1;
        private WorksheetValues worksheetValues2;

        private ArrayList columnList = new ArrayList();
        private int numSkipped = 0;
        private ArrayList SkipedParcels = new ArrayList();
        private ArrayList contiguousParcelList = new ArrayList();

        struct TaxMapIdWithSpan
        {
            public string taxMapId;
            public string span;
            public TaxMapIdWithSpan(string taxMapId, string span)
            {
                this.taxMapId = taxMapId;
                this.span = span;
            }
        }

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

        private int nemrcParcelIdIndex;
        private int nemrcParcelSubIdIndex;
        private int nemrcDescriptionIndex;
        private int nemrcTaxMapIdIndex;
        private int nemrcNotesContiguousIndex;
        private int nemrcContiguousIDIndex;
        private int nemrcContiguousSubIDIndex;
        private int nemrcAcresIndex;
        private int nemrcName1Index;
        //****************************************************************************************************************************
        public CreateListOfParcelIds(ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
            : base(progressBar, ProgressBarLabel)
        {
        }
        //****************************************************************************************************************************
        public void CreateList(bool workingGrandList)
        {
            try
            {
                worksheetValues1 = SelectInputFile(ExcelClass.NemrcExportsFolder, "      TaxMapInactives", "xlsx", ColumnName.Span, workingGrandList);
                if (worksheetValues1 == null)
                {
                    return;
                }
                worksheetValues2 = SelectInputFile(ExcelClass.PatriotExportsFolder, "      TaxMapActives", "xlsx", ColumnName.InactiveSpan, workingGrandList);
                if (worksheetValues2 == null)
                {
                    return;
                }
                SetColumnIndexes();

                workbookPath = ReportsFolder + "\\NEMRC_GrandListWithNewParcelIds.xlsx";
                CreateNewExcelWorkbook("NEMRC_GrandListWithNewParcelIds");

                CreateColumnList();
                WriteHeadings();

                cumulativeRowIndex = 1;
                ProgressBarLabel.Text = "Inactive";
                CopySelectedColumns(worksheetValues1, "I");
                ProgressBarLabel.Text = "Active";
                CopySelectedColumns(worksheetValues2, "A");

                worksheetValues1.CloseWorkbook();
                worksheetValues2.CloseWorkbook();
                ProgressBarLabel.Visible = false;
                MessageBox.Show("Num Skipped: " + numSkipped);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private void SetColumnIndexes()
        {
            nemrcParcelIdIndex = GetColumnNum(worksheetValues1, ColumnName.ParcelId);
            nemrcParcelSubIdIndex = GetColumnNum(worksheetValues1, ColumnName.ParcelSubId);
            nemrcTaxMapIdIndex = GetColumnNum(worksheetValues1, ColumnName.TaxMapId);
            nemrcNotesContiguousIndex = GetColumnNum(worksheetValues1, ColumnName.NotesContiguous);
            nemrcContiguousIDIndex = GetColumnNum(worksheetValues1, ColumnName.ContiguousID);
            nemrcContiguousSubIDIndex = GetColumnNum(worksheetValues1, ColumnName.ContiguousSubID);
            nemrcDescriptionIndex = GetColumnNum(worksheetValues1, ColumnName.Description);
            nemrcAcresIndex = GetColumnNum(worksheetValues1, ColumnName.Acres);
            nemrcName1Index = GetColumnNum(worksheetValues1, ColumnName.Name1);
        }
        //****************************************************************************************************************************
        private void CopySelectedColumns(WorksheetValues worksheetValues, string activeInactive)
        {
            try
            {
                ProgressBarLabel.Visible = true;
                progressBar.Visible = true;
                progressBar.Value = 0;
                progressBar.Step = 1;
                progressBar.Maximum = worksheetValues.numRows;
                var rowIndex = 0;
                for (int i = 1; i <= worksheetValues.numRows; i++)
                {
                    rowIndex++;
                    string parcelId = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcParcelIdIndex).Trim();
                    string contiguousParcelId = ExtractContiguousId(GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcNotesContiguousIndex));
                    if (contiguousParcelId.Contains("(A-9)"))
                    {
                        contiguousParcelId = contiguousParcelId.Replace("(A-9)", "A-9");
                    }
                    string contigParcelId = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcContiguousIDIndex).Trim();
                    string contigParcelIdNoZeros = RemoveLeadingZeros(contigParcelId);
                    string contigSubParcelId = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcContiguousSubIDIndex);
                    if (contigSubParcelId == "(A-9)")
                    {
                        contigSubParcelId = "A-9";
                    }
                    string contiguousCombinedParcelId = CombineParcelWithSubparcel(contigParcelIdNoZeros.Trim(), contigSubParcelId.Trim());
                    contiguousCombinedParcelId = RemoveLeadingZeros(contiguousCombinedParcelId);
                    if (contiguousParcelId != contiguousCombinedParcelId)
                    {
                    }
                    string taxMapId = GetCellValue(worksheetValues.inputWorksheet, rowIndex, nemrcTaxMapIdIndex);
                    int indexOfDot = taxMapId.IndexOf('.');
                    string subParcelId = (indexOfDot > 0) ? taxMapId.Substring(indexOfDot + 1) : "";
                    //if (!String.IsNullOrEmpty(subParcelId))
                    //{
                    //    parcelId += ("-" + subParcelId);
                    //}
                    if (subParcelId == "(A-9)")
                    {
                        subParcelId = "A-9";
                    }
                    ConvertParcelId(worksheetValues.inputWorksheet, taxMapId, parcelId, subParcelId, activeInactive, rowIndex, contigParcelId, contigSubParcelId);
                    //outputWorksheet.Cells[cumulativeRowIndex + rowIndex + 1, 1].Value = parcelId;
                    //int columnNum = 1;
                    //foreach (NemrcColumns nemrcColumn in columnList)
                    // {
                    //     columnNum++;
                    //     outputWorksheet.Cells[cumulativeRowIndex + rowIndex + 1, columnNum].Value = GetCellValue(rowIndex, nemrcColumn.columnNum);
                    //     if (columnNum == 6)
                    //     {
                    //         columnNum++;
                    //         outputWorksheet.Cells[cumulativeRowIndex + rowIndex + 1, columnNum].Value = activeInactive;
                    //     }
                    // }
                    progressBar.PerformStep();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private string RemoveLeadingZeros(string parcelId)
        {
            if (!String.IsNullOrEmpty(parcelId))
            {
                int indexOfDash = parcelId.IndexOf('-');
                if (indexOfDash < 1)
                {
                }
                else
                {
                    parcelId = parcelId.Remove(0, indexOfDash - 1);
                }
            }
            return parcelId;
        }
        //****************************************************************************************************************************
        private void ConvertParcelId(OfficeOpenXml.ExcelWorksheet inputWorksheet, string taxMapId, string parcelId, string subParcelId, string activeInactive, int rowIndex, string contiguousParcelId, string contiguousSubParcelId)
        {
            string initialParcelId = parcelId;
            string initialSubParcelId = subParcelId;
            if (String.IsNullOrEmpty(parcelId))
            {
                numSkipped++;
                return;
            }
            string newParcelId = (parcelId.Contains("MH")) ? MobileHome(ref subParcelId) : RegularParcelId(parcelId, ref subParcelId);
            if (String.IsNullOrEmpty(newParcelId))
            {
                newParcelId = SpecialParcelId(parcelId, ref subParcelId);
                if (String.IsNullOrEmpty(newParcelId))
                {
                    SkipedParcels.Add(CombineParcelWithSubparcel(parcelId, initialSubParcelId));
                    numSkipped++;
                    return;
                }
            }
            cumulativeRowIndex++;
            outputWorksheet.Cells[cumulativeRowIndex, 1].Value = activeInactive;
            outputWorksheet.Cells[cumulativeRowIndex, 2].Value = initialParcelId;
            outputWorksheet.Cells[cumulativeRowIndex, 3].Value = initialSubParcelId;
            outputWorksheet.Cells[cumulativeRowIndex, 4].Value = CombineParcelWithSubparcel(parcelId, initialSubParcelId);
            //outputWorksheet.Cells[cumulativeRowIndex, 4] = newParcelId;
            string mapNoDash = newParcelId.Replace("-", "");
            newParcelId = CombineParcelWithSubparcel(newParcelId, subParcelId);
            //outputWorksheet.Cells[cumulativeRowIndex, 7] = newParcelId;
            outputWorksheet.Cells[cumulativeRowIndex, 6].Value = mapNoDash;
            outputWorksheet.Cells[cumulativeRowIndex, 7].Value = subParcelId;
            outputWorksheet.Cells[cumulativeRowIndex, 8].Value = CombineParcelWithSubparcel(mapNoDash, subParcelId);
            //if (!String.IsNullOrEmpty(subParcelId))
            //{
            //    mapNoDash += '.' + subParcelId;
            //}
            //outputWorksheet.Cells[cumulativeRowIndex, 6] = mapNoDash;
            string description = GetCellValue(inputWorksheet, rowIndex, nemrcDescriptionIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 10].Value = description;

            string totalAcres = GetCellValue(inputWorksheet, rowIndex, nemrcAcresIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 15].Value = totalAcres;

            if (activeInactive == "A")
            {
                string parcelAcres = ActiveAcres(contiguousParcelList, taxMapId, totalAcres);
                if (parcelAcres != totalAcres)
                {
                    outputWorksheet.Cells[cumulativeRowIndex, 16].Value = GetCellValue(inputWorksheet, rowIndex, nemrcName1Index).Trim();
                    outputWorksheet.Cells[cumulativeRowIndex, 17].Value = parcelAcres;
                }
                return;
            }

            string initialContigParcelId = contiguousParcelId;
            string initialContigSubParcelId = contiguousSubParcelId;
            string contigTaxMapId = CombineParcelWithSubparcel(contiguousParcelId, contiguousSubParcelId);
            contigTaxMapId = RemoveLeadingZeros(contigTaxMapId);

            string newContiguousParcelId = RegularParcelId(contiguousParcelId, ref contiguousSubParcelId);
            if (String.IsNullOrEmpty(newContiguousParcelId))
            {
                newContiguousParcelId = SpecialParcelId(newContiguousParcelId, ref contiguousSubParcelId);
            }
            newContiguousParcelId = newContiguousParcelId.Replace("-", "");
            string combinedParcelId = CombineParcelWithSubparcel(initialContigParcelId, initialContigSubParcelId);
            outputWorksheet.Cells[cumulativeRowIndex, 12].Value = combinedParcelId;
            combinedParcelId = CombineParcelWithSubparcel(newContiguousParcelId, contiguousSubParcelId);
            outputWorksheet.Cells[cumulativeRowIndex, 13].Value = combinedParcelId;
            AddContiguousParcelToList(contiguousParcelList, taxMapId, contigTaxMapId, totalAcres);
        }
        //****************************************************************************************************************************
        private string ActiveAcres(ArrayList contiguousParcelList, string taxMapId, string acres)
        {
            int colIndex = 18;
            double totalAcres = 0.0;
            foreach (ContiguousParcelId contiguousParcelId in contiguousParcelList)
            {
                if (contiguousParcelId.contigTaxMapId == taxMapId)
                {
                    totalAcres += contiguousParcelId.acres;
                    outputWorksheet.Cells[cumulativeRowIndex, colIndex++].Value = contiguousParcelId.taxMapId;
                    outputWorksheet.Cells[cumulativeRowIndex, colIndex++].Value = contiguousParcelId.acres;
                }
            }
            return (acres.ToDouble() - totalAcres).ToString();
        }
        //****************************************************************************************************************************
        private string CombineParcelWithSubparcel(string parcelId, string subParcelId)
        {
            if (!String.IsNullOrEmpty(subParcelId))
            {
                parcelId += "." + subParcelId;
            }
            return parcelId;
        }
        private string SpecialParcelId(string parcelId, ref string subParcelId)
        {
            if (parcelId == "Q-23-A")
            {
                subParcelId = "3";
                return "Q-023";
            }
            if (parcelId == "0S-66A")
            {
                subParcelId = "3.3";
                return "S-066";
            }
            if (parcelId == "00J-9A")
            {
                subParcelId = "1.1";
                return "J-009";
            }
            if (parcelId == "0O-44A")
            {
                subParcelId = "2";
                return "O-044";
            }
            if (parcelId == "00COMC")
            {
                subParcelId = "CMCST";
                return "Z-000";
            }
            if (parcelId == "CQQQQQ")
            {
                subParcelId = "GMP";
                return "Z-000";
            }
            if (parcelId == "00SVCC")
            {
                subParcelId = "SVCBL";
                return "Z-000";
            }
            if (parcelId == "VELCO_")
            {
                subParcelId = "VELCO";
                return "Z-000";
            }
            if (parcelId == "test00")
            {
                subParcelId = "TEST";
                return "Z-000";
            }
            return "";
        }
        //****************************************************************************************************************************
        private bool Condo(string parcelId, string subParcelId)
        {
            char firstCharOfSubId = (string.IsNullOrEmpty(subParcelId)) ? ' ' : subParcelId[0];
            if (parcelId == "F-031" && firstCharOfSubId == '2')
            {
                if (subParcelId.Length == 1)
                {
                }
                else
                {
                    return true;
                }
            }
            if (parcelId == "F-036" && firstCharOfSubId == '1' && subParcelId.Length > 1)
            {
                if (subParcelId.Length == 1)
                {
                }
                else
                {
                    return true;
                }
            }
            if (parcelId == "J-009" && subParcelId.Contains("18"))
            {
                if (subParcelId.Contains('A'))
                {
                }
                else
                {
                    return true;
                }
            }
            return false;
        }
        //****************************************************************************************************************************
        private string RegularParcelId(string parcelId, ref string subParcelId)
        {
            LeasedLot(ref parcelId, ref subParcelId);
            string newParcelId = RemoveLeadingZeros(parcelId);
            if (newParcelId.Length < 2)
            {
                return "";
            }
            string map = NewParcelString(newParcelId, out newParcelId);
            int intParcelId = 0;
            if (newParcelId.Contains("LL"))
            {
                newParcelId = newParcelId.Replace("LL", "");
                intParcelId = 190;
            }
            else if (newParcelId.Contains("L"))
            {
                newParcelId = newParcelId.Replace("L", "");
                intParcelId = 100;
            }
            string returnParcelId = ReturnNewMarcelId(map, newParcelId, intParcelId);
            if (Condo(returnParcelId, subParcelId))
            {
                subParcelId = "C" + subParcelId;
            }
            return returnParcelId;
        }
        //****************************************************************************************************************************
        private void LeasedLot(ref string parcelId, ref string subparcelId)
        {
            if (parcelId == "00J-L1" && subparcelId == "1")
            {
                parcelId = "000F-5";
                subparcelId = "1.3";
                return; 
            }
            if (parcelId == "0J-L21" && String.IsNullOrEmpty(subparcelId))
            {
                parcelId = "00J-14";
                subparcelId = "";
                return;
            }
            if (parcelId == "0J-L26" && subparcelId == "14")
            {
                parcelId = "000J-9";
                subparcelId = "20";
                return;
            }
            if (parcelId == "0O-LL3" && String.IsNullOrEmpty(subparcelId))
            {
                parcelId = "000O-3";
                subparcelId = "";
                return;
            }
        }
        //****************************************************************************************************************************
        private string MobileHome(ref string subParcelId)
        {
            string intParcelId;
            string map = NewParcelString(subParcelId, out intParcelId);
            int indexOfPoint = intParcelId.IndexOf('.');
            subParcelId = "MH";
            if (indexOfPoint > 0)
            {
                subParcelId += intParcelId.Substring(indexOfPoint + 1);
                intParcelId = intParcelId.Substring(0, indexOfPoint); 
            }
            return ReturnNewMarcelId(map, intParcelId);
        }
        //****************************************************************************************************************************
        private string ReturnNewMarcelId(string map, string newParcelId, int intParcelId = 0)
        {
            try
            {
                intParcelId += Convert.ToInt32(newParcelId);
            }
            catch (Exception)
            {
                return "";
                //MessageBox.Show(parcelId + " " + ex.Message);
            }
            return map + String.Format("{0:000}", intParcelId);
        }
        //****************************************************************************************************************************
        private string NewParcelString(string shortParcelID, out string intParcelId)
        {
            string map = shortParcelID.Substring(0, 2);
            if (map.Length != 2 || map[1] != '-')
            {
                if (shortParcelID == "P27.4.1")
                {
                    map = "P-";
                    intParcelId = "27.4.1";
                    return map; 
                }
                else
                {
                    intParcelId = "";
                    return "";
                }
            }
            intParcelId = shortParcelID.Remove(0, 2);
            return map;
        }
        //****************************************************************************************************************************
        private void CreateColumnList()
        {
            columnList.Add(new NemrcColumns(1, "Status"));
            columnList.Add(new NemrcColumns(2, "Old ParcelId"));
            columnList.Add(new NemrcColumns(3, "Old Sub ParcelId"));
            columnList.Add(new NemrcColumns(4, "Combined Old"));
            columnList.Add(new NemrcColumns(6, "New ParcelId"));
            columnList.Add(new NemrcColumns(7, "New Sub ParcelId"));
            columnList.Add(new NemrcColumns(8, "Combined New"));
            columnList.Add(new NemrcColumns(10, "Description"));
            columnList.Add(new NemrcColumns(12, "Old Contig Combined Parcel Id"));
            columnList.Add(new NemrcColumns(13, "New Contig Combined Parcel Id"));
            columnList.Add(new NemrcColumns(15, "Total Acres"));
            columnList.Add(new NemrcColumns(16, "Active Parcel"));
            columnList.Add(new NemrcColumns(17, "Active Acres"));
            FormatStringCells("C1:C1596");
            FormatStringCells("G1:G1596");
        }
        //****************************************************************************************************************************
        public void WriteHeadings()
        {
            int columnNum = 0;
            foreach (NemrcColumns nemrcColumn in columnList)
            {
                columnNum++;
                outputWorksheet.Cells[1, nemrcColumn.columnNum].Value = nemrcColumn.columnName;
            }
        }
        //****************************************************************************************************************************
    }
}
