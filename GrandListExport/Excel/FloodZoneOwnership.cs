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
    public class FloodZoneOwnership : ExcelClass 
    {
        private WorksheetValues worksheetValues1;
        private WorksheetValues worksheetValues2;
        private ArrayList NemrcNotInPatriotRow = new ArrayList();
        private ArrayList PatriotNotInNemrcRow = new ArrayList();
        private ArrayList PatriotBlankSpanRow = new ArrayList();
        private ArrayList differentValues = new ArrayList();
        private bool valuesOnly = false;
        private string nemrcTaxMapId = "";
        private string nemrcName1 = "";
        private string nemrcName2 = "";
        private string nemrcAddr1 = "";
        private string nemrcAddr2 = "";
        private string nemrcCity = "";
        private string nemrcState = "";
        private string nemrcZip = "";
        private string compareNemrcAddr = "";

        private int nemrcSpanIndex;
        private int nemrcName1Index;
        private int nemrcName2Index;
        private int nemrcAddr1Index;
        private int nemrcAddr2Index;
        private int nemrcCityIndex;
        private int nemrcStateIndex;
        private int nemrcZipIndex;
        private int nemrcTaxMapIdIndex;

        private int FloodZoneSpanIndex;
        ArrayList GreetingMapIds = new ArrayList();

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
        public FloodZoneOwnership(System.Windows.Forms.ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
            : base(progressBar, ProgressBarLabel)
        {
        }
        //****************************************************************************************************************************
        public void FloodZoneOwnershipSetup(bool workingGrandList, bool valuesOnly)
        {
            try
            {
                this.valuesOnly = valuesOnly;

                worksheetValues1 = SelectInputFile(ExcelClass.NemrcExportsFolder, "NemrcActives", "csv", ColumnName.Span, workingGrandList);
                if (worksheetValues1 == null)
                {
                    return;
                }
                worksheetValues2 = SelectInputFile(ExcelClass.PatriotExportsFolder, "FloodzoneProperties", "xlsx", ColumnName.none, workingGrandList);
                if (worksheetValues2 == null)
                {
                    return;
                }
                SetColumnIndexes();
                workbookPath = ExcelClass.ReportsFolder + "\\" + "FloodZoneAddresses" + ".xlsx";
                CreateNewExcelWorkbook("FloodZoneAddresses");
                CreateFloodZoneAddress();
                worksheetValues1.CloseWorkbook();
                worksheetValues2.CloseWorkbook();
                CloseOutput();
                workbookPath = ExcelClass.ReportsFolder + "\\" + "FloodZoneMapIds" + ".xlsx";
                CreateNewExcelWorkbook("FloodZoneMapIds");
                AddMapidsToWorkbook();
                CloseOutput();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void AddMapidsToWorkbook()
        {
            outputIndex = 1;
            foreach (string mapId in GreetingMapIds)
            {
                outputWorksheet.Cells[outputIndex, 1].Value = mapId;
                outputIndex++;
            }
        }
        //****************************************************************************************************************************
        private void WriteHeadings()
        {
            DateTime today = DateTime.Today;
            outputWorksheet.Cells[1, 1].Value = today.ToString("MM/dd/yyyy"); ;
            outputWorksheet.Cells[2, 1].Value = "SPAN";
            if (valuesOnly)
            {
                outputWorksheet.Cells[1, 2].Value = "Value Differences";
                outputWorksheet.Cells[2, 2].Value = "Nemrc Acres";
                outputWorksheet.Cells[2, 3].Value = "Patriot Acres";
                outputWorksheet.Cells[2, 4].Value = "Nemrc Value";
                outputWorksheet.Cells[2, 5].Value = "Patriot Value";
                outputWorksheet.Cells[2, 6].Value = "Nemrc Name";
            }
            else
            {
                outputWorksheet.Cells[1, 2].Value = "Name/Address Differences";
                outputWorksheet.Cells[2, 2].Value = "Nemrc Name";
                outputWorksheet.Cells[2, 3].Value = "Patriot Name";
                outputWorksheet.Cells[2, 4].Value = "Nemrc Address 1";
                outputWorksheet.Cells[2, 5].Value = "Patriot Address 1";
                outputWorksheet.Cells[2, 6].Value = "Nemrc Address 2";
                outputWorksheet.Cells[2, 7].Value = "Patriot Address 2";
                outputWorksheet.Cells[2, 8].Value = "Nemrc City";
                outputWorksheet.Cells[2, 9].Value = "Patriot City";
                outputWorksheet.Cells[2, 10].Value = "Nemrc State";
                outputWorksheet.Cells[2, 11].Value = "Patriot State";
                outputWorksheet.Cells[2, 12].Value = "Nemrc Zip";
                outputWorksheet.Cells[2, 13].Value = "Patriot Zip";
                outputWorksheet.Cells[2, 14].Value = "Nemrc Acres";
                outputWorksheet.Cells[2, 15].Value = "Patriot Acres";
            }
        }
        //****************************************************************************************************************************
        private void SetColumnIndexes()
        {
            nemrcSpanIndex = GetColumnNum(worksheetValues1, ColumnName.Span);
            nemrcName1Index = GetColumnNum(worksheetValues1, ColumnName.Name1);
            nemrcName2Index = GetColumnNum(worksheetValues1, ColumnName.Name2);
            nemrcAddr1Index = GetColumnNum(worksheetValues1, ColumnName.Addr1);
            nemrcAddr2Index = GetColumnNum(worksheetValues1, ColumnName.Addr2);
            nemrcCityIndex = GetColumnNum(worksheetValues1, ColumnName.City);
            nemrcStateIndex = GetColumnNum(worksheetValues1, ColumnName.State);
            nemrcZipIndex = GetColumnNum(worksheetValues1, ColumnName.Zip);
            nemrcTaxMapIdIndex = GetColumnNum(worksheetValues1, ColumnName.TaxMapId);

            FloodZoneSpanIndex = GetColumnNum(worksheetValues2, ColumnName.FloodZoneSpan);
        }
        //****************************************************************************************************************************
        private void CreateFloodZoneAddress()
        {
            this.ProgressBarLabel.Visible = true;
            this.progressBar.Visible = true;
            this.progressBar.Value = 0;
            this.progressBar.Step = 10;
            this.progressBar.Maximum = worksheetValues1.numRows;
            int rowIndex1 = 1;
            int rowIndex2 = 2;  // Skip Headings
            outputIndex = 1;
            while (rowIndex1 <= worksheetValues1.numRows && rowIndex2 <= worksheetValues2.numRows)
            {
                try
                {
                    string grandListSpan = GetCellValue(worksheetValues1.inputWorksheet, rowIndex1, nemrcSpanIndex);
                    string FloodZoneSpan = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, FloodZoneSpanIndex);
                    char greaterLessthanEqual = GreaterLessthanEqual(grandListSpan, FloodZoneSpan);
                    if (greaterLessthanEqual == '=')
                    {
                        SpanNumberEqual(rowIndex1, rowIndex2, grandListSpan);
                        IncrementProgressBar();
                        rowIndex1++;
                        rowIndex2++;
                    }
                    else
                    if (greaterLessthanEqual == '>')

                    {
                        rowIndex2++;
                    }
                    else
                    {
                        rowIndex1++;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        //****************************************************************************************************************************
        private void SpanNumberEqual(int rowIndex,
                                     int rowIndex2,
                                     string grandListSpan)
        {
            nemrcTaxMapId = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcTaxMapIdIndex);
            nemrcName1 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcName1Index);
            if (nemrcName1.ToUpper().Contains("VERMONT STATE OF"))
            {
                return;
            }
            if (nemrcName1.ToUpper().Contains("TOWN OF JAMAICA"))
            {
                return;
            }
            if (nemrcName1.ToUpper().Contains("JAMAICA TOWN OF"))
            {
                return;
            }
            if (nemrcName1.ToUpper().Contains("UNITED STATES OF AMERICA"))
            {
                return;
            }
            nemrcName2 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcName2Index);
            nemrcAddr1 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcAddr1Index);
            nemrcAddr2 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcAddr2Index);
            if (String.IsNullOrEmpty(nemrcAddr1))
            {
                nemrcAddr1 = nemrcAddr2;
            }
            nemrcCity = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcCityIndex);
            nemrcState = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcStateIndex);
            nemrcZip = GetZipCode(GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcZipIndex));

            AdjustNemrcAddress(ref nemrcAddr1, nemrcAddr2);
            compareNemrcAddr = GetAddress(nemrcAddr1);
            GreetingMapIds.Add(nemrcTaxMapId);
            outputWorksheet.Cells[outputIndex, 1].Value = nemrcName1;
            outputWorksheet.Cells[outputIndex, 2].Value = "Map ID " + nemrcTaxMapId;
            outputWorksheet.Cells[outputIndex, 3].Value = nemrcAddr1;
            outputWorksheet.Cells[outputIndex, 4].Value = nemrcCity.Trim() + ", " + nemrcState.Trim() + " " + nemrcZip.Trim();
            outputIndex++;
        }
        //****************************************************************************************************************************
        private void AdjustNemrcAddress(ref string nemrcAddr1, string nemrcAddr2)
        {
            if (nemrcAddr2.Length == 0)
            {
                return;
            }
            if (nemrcAddr1.Substring(0, 3).ToUpper() == "C/O")
            {
                nemrcAddr1 = nemrcAddr2.Trim() + " " + nemrcAddr1.Trim();
            }
            else
            {
                nemrcAddr1 = nemrcAddr1.Trim() + " " + nemrcAddr2.Trim();
            }
/*            if (nemrcAddr2[0] == '#')
            {
                nemrcAddr1 += nemrcAddr2;
            }
            if (nemrcAddr2.ToLower().Contains("apt"))
            {
                nemrcAddr1 += nemrcAddr2;
            }
            if (nemrcAddr2.ToLower().Contains("unit"))
            {
                nemrcAddr1 += nemrcAddr2;
            }
            */
        }
        //****************************************************************************************************************************
        private string GetAddress(string address)
        {
            address = address.ToLower().Replace(" ", "");
            address = address.ToLower().Replace(".", "");
            address = address.ToLower().Replace(",", "");
            address = address.ToLower().Replace("route", "");
            address = address.ToLower().Replace("lane", "ln");
            address = address.ToLower().Replace("road", "rd");
            address = address.ToLower().Replace("drive", "dr");
            address = address.ToLower().Replace("street", "st");
            address = address.ToLower().Replace("fifth", "5th");
            address = address.ToLower().Replace("south", "s");
            address = address.ToLower().Replace("boulevard", "blvd");
            address = address.ToLower().Replace("avenue", "ave");
            address = address.ToLower().Replace("vermont", "vt");
            address = address.ToLower().Replace("west", "w");
            address = address.ToLower().Replace("court", "ct");
            return address.ToLower();
        }
        //****************************************************************************************************************************
        private string GetZipCode(string zip)
        {
            if (string.IsNullOrEmpty(zip))
            {
                return "";
            }
            zip = zip.Replace('O', '0');
            int indexOfDash = zip.IndexOf('-');
            string zipCode = (indexOfDash < 0) ? zip : zip.Substring(0, indexOfDash);
            try
            {
                int intZip = Convert.ToInt32(zipCode);
                if (intZip == 0)
                {
                    MessageBox.Show("Invalid Zip Code: " + zip);
                }
                zip = String.Format("{0:00000}", intZip);
                return zip;
            }
            catch (Exception)
            {
                MessageBox.Show("Invalid Zip Code: " + zip);
                return "";
            }
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
    }
}
