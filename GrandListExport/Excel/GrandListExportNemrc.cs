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
    public class GrandListExportNemrc : ExcelClass
    {
        private WorksheetValues worksheetValues1;

        private ArrayList columnList = new ArrayList();

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
        private int nemrcName1Index;
        private int nemrcName2Index;
        private int nemrcAddress1Index;
        private int nemrcAddress2Index;
        private int nemrcCityIndex;
        private int nemrcStateIndex;
        private int nemrcZipIndex;
        private int nemrcLocationAIndex;
        private int nemrcLocationBIndex;
        private int nemrcLocationCIndex;
        private int nemrcStreetNumIndex;
        private int nemrcStreetNameIndex;
        private int nemrcTaxMapIdIndex;
        private int nemrcSpanIndex;
        private int nemrcDescriptionIndex;
        private int nemrcOwnerIndex;
        private int nemrcDateHomesteadIndex;
        private int nemrcBuildingValueIndex;

        //****************************************************************************************************************************
        public GrandListExportNemrc(ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
            : base(progressBar, ProgressBarLabel)
        {
        }
        //****************************************************************************************************************************
        public void CreateList(bool workingGrandList, int grandListYear = 0)
        {
            try
            {
                CreateActiveInactives("NemrcActives", workingGrandList, grandListYear);
                CreateActiveInactives("NemrcInactives", workingGrandList, grandListYear);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private void CreateActiveInactives(string whichFile, bool workingGrandList, int grandListYear)
        {
            worksheetValues1 = SelectInputFile(ExcelClass.NemrcExportsFolder, whichFile, "xlsx", ColumnName.Span, workingGrandList, grandListYear);
            if (worksheetValues1 == null)
            {
                return;
            }
            SetColumnIndexes();
            workbookPath = ReportsFolder + "\\GrandListExport" + whichFile + ".xlsx";
            CreateNewExcelWorkbook("GrandListExport" + whichFile);
            CreateColumnList();
            WriteHeadings();
            cumulativeRowIndex = 1;
            ProgressBarLabel.Text = whichFile;
            CopySelectedColumns(worksheetValues1);
            worksheetValues1.CloseWorkbook();
            CloseOutput();
            ProgressBarLabel.Visible = false;
        }
        //****************************************************************************************************************************
        private void SetColumnIndexes()
        {
            nemrcParcelIdIndex = GetColumnNum(worksheetValues1, ColumnName.ParcelId);
            nemrcParcelSubIdIndex = GetColumnNum(worksheetValues1, ColumnName.ParcelSubId);
            nemrcName1Index = GetColumnNum(worksheetValues1, ColumnName.Name1);
            nemrcName2Index = GetColumnNum(worksheetValues1, ColumnName.Name2);
            nemrcAddress1Index = GetColumnNum(worksheetValues1, ColumnName.Addr1);
            nemrcAddress2Index = GetColumnNum(worksheetValues1, ColumnName.Addr2);
            nemrcCityIndex = GetColumnNum(worksheetValues1, ColumnName.City);
            nemrcStateIndex = GetColumnNum(worksheetValues1, ColumnName.State);
            nemrcZipIndex = GetColumnNum(worksheetValues1, ColumnName.Zip);
            nemrcLocationAIndex = GetColumnNum(worksheetValues1, ColumnName.LocationA);
            nemrcLocationBIndex = GetColumnNum(worksheetValues1, ColumnName.LocationB);
            nemrcLocationCIndex = GetColumnNum(worksheetValues1, ColumnName.LocationC);
            nemrcStreetNumIndex = GetColumnNum(worksheetValues1, ColumnName.StreetNum);
            nemrcStreetNameIndex = GetColumnNum(worksheetValues1, ColumnName.StreetName);
            nemrcTaxMapIdIndex = GetColumnNum(worksheetValues1, ColumnName.TaxMapId);
            nemrcDescriptionIndex = GetColumnNum(worksheetValues1, ColumnName.Description);
            nemrcOwnerIndex = GetColumnNum(worksheetValues1, ColumnName.Owner);
            nemrcDateHomesteadIndex = GetColumnNum(worksheetValues1, ColumnName.DateHomestead);
            nemrcSpanIndex = GetColumnNum(worksheetValues1, ColumnName.Span);
            nemrcBuildingValueIndex = GetColumnNum(worksheetValues1, ColumnName.buildingValue);
        }
        //****************************************************************************************************************************
        private void CopySelectedColumns(WorksheetValues worksheetValues)
        {
            try
            {
                ProgressBarLabel.Visible = true;
                progressBar.Visible = true;
                progressBar.Value = 0;
                progressBar.Step = 1;
                progressBar.Maximum = worksheetValues.numRows;
                var rowIndex = 1;
                for (int i = 1; i < worksheetValues.numRows; i++)
                {
                    rowIndex++;
                    MoveNemrcFields(worksheetValues.inputWorksheet, rowIndex);
                    progressBar.PerformStep();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private void MoveNemrcFields(OfficeOpenXml.ExcelWorksheet inputWorksheet, int rowIndex)
        {
            cumulativeRowIndex++;
            string taxMapId = GetCellValue(inputWorksheet, rowIndex, nemrcTaxMapIdIndex);
            int indexOfDot = taxMapId.IndexOf('.');
            string subParcelId = (indexOfDot > 0) ? taxMapId.Substring(indexOfDot + 1) : "";
            outputWorksheet.Cells[cumulativeRowIndex, 1].Value = GetCellValue(inputWorksheet, rowIndex, nemrcParcelIdIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 2].Value = subParcelId;
            outputWorksheet.Cells[cumulativeRowIndex, 3].Value = GetCellValue(inputWorksheet, rowIndex, nemrcName1Index).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 4].Value = GetCellValue(inputWorksheet, rowIndex, nemrcName2Index).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 5].Value = GetCellValue(inputWorksheet, rowIndex, nemrcAddress1Index).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 6].Value = GetCellValue(inputWorksheet, rowIndex, nemrcAddress2Index).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 7].Value = GetCellValue(inputWorksheet, rowIndex, nemrcCityIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 8].Value = GetCellValue(inputWorksheet, rowIndex, nemrcStateIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 9].Value = GetCellValue(inputWorksheet, rowIndex, nemrcZipIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 10].Value = GetCellValue(inputWorksheet, rowIndex, nemrcLocationAIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 11].Value = GetCellValue(inputWorksheet, rowIndex, nemrcLocationBIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 12].Value = GetCellValue(inputWorksheet, rowIndex, nemrcLocationCIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 13].Value = GetCellValue(inputWorksheet, rowIndex, nemrcStreetNumIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 14].Value = GetCellValue(inputWorksheet, rowIndex, nemrcStreetNameIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 15].Value = taxMapId;
            string description = GetCellValue(inputWorksheet, rowIndex, nemrcDescriptionIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 16].Value = description;
            outputWorksheet.Cells[cumulativeRowIndex, 17].Value = GetCellValue(inputWorksheet, rowIndex, nemrcOwnerIndex).Trim();
            string dateHomestead = GetCellValue(inputWorksheet, rowIndex, nemrcDateHomesteadIndex).Trim();
            if (!String.IsNullOrEmpty(dateHomestead))
            {
                if (dateHomestead.IndexOf('/') == 0)
                {
                    dateHomestead = dateHomestead.Replace("/", "").Trim();
                }
            }
            outputWorksheet.Cells[cumulativeRowIndex, 18].Value = dateHomestead;
            outputWorksheet.Cells[cumulativeRowIndex, 19].Value = GetCellValue(inputWorksheet, rowIndex, nemrcSpanIndex).Trim();
            outputWorksheet.Cells[cumulativeRowIndex, 20].Value = VacantLand(GetCellValue(inputWorksheet, rowIndex, nemrcBuildingValueIndex).Trim(), description);
            outputWorksheet.Cells[cumulativeRowIndex, 21].Value = GetCellValue(inputWorksheet, rowIndex, nemrcBuildingValueIndex).Trim();
        }
        //****************************************************************************************************************************
        private string VacantLand(string buildingValue, string description)
        {
            int iBuildingValue;
            try
            {
                iBuildingValue = Convert.ToInt32(buildingValue);
                if (iBuildingValue == 0)
                {
                    return "1";
                }
                if (iBuildingValue > 16000)
                {
                    return "0";
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to Convert To Int: " + buildingValue);
                return "0";
            }
            if (String.IsNullOrEmpty(description))
            {
                return "1";
            }
            if (description == "LAND")
            {
                return "1";
            }
            if (description == "LAND ONLY")
            {
                return "1";
            }
            if (description == "LAND & SHED")
            {
                return "1";
            }
            if (description.Contains("LOT") && iBuildingValue < 4000)
            {
                return "1";
            }
            if (description == "LAND & SUBSTATION")
            {
                return "1";
            }
            if (description == "LAND & SHED")
            {
                return "1";
            }
            if (description == "INCLUDES L-3")
            {
                return "1";
            }
            if (description == "LAND 10 ACRES UNDER USFE; INCL Q-5")
            {
                return "1";
            }
            if (description == "LAND AND CAMPER")
            {
                return "1";
            }
            if (description == "LAND, SHED/BARN")
            {
                return "1";
            }
            return "0";
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
        //****************************************************************************************************************************
        private void CreateColumnList()
        {
            columnList.Add(new NemrcColumns(1, "ParcelId"));
            columnList.Add(new NemrcColumns(2, "SubParcelId"));
            columnList.Add(new NemrcColumns(3, "Name1"));
            columnList.Add(new NemrcColumns(4, "Name2"));
            columnList.Add(new NemrcColumns(5, "Address1"));
            columnList.Add(new NemrcColumns(6, "Address2"));
            columnList.Add(new NemrcColumns(7, "City"));
            columnList.Add(new NemrcColumns(8, "State"));
            columnList.Add(new NemrcColumns(9, "Zip"));
            columnList.Add(new NemrcColumns(10, "LocationA"));
            columnList.Add(new NemrcColumns(11, "LocationB"));
            columnList.Add(new NemrcColumns(12, "LocationC"));
            columnList.Add(new NemrcColumns(13, "StreetNum"));
            columnList.Add(new NemrcColumns(14, "StreetName"));
            columnList.Add(new NemrcColumns(15, "TaxMapID"));
            columnList.Add(new NemrcColumns(16, "Description"));
            columnList.Add(new NemrcColumns(17, "Owner"));
            columnList.Add(new NemrcColumns(18, "DateHomestead"));
            columnList.Add(new NemrcColumns(19, "Span"));
            columnList.Add(new NemrcColumns(20, "VacantLand"));
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
