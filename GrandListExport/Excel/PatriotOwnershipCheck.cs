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
    public class PatriotOwnershipCheck : ExcelClass
    {
        private WorksheetValues worksheetValues1;
        private WorksheetValues worksheetValues2;
        private ArrayList NemrcNotInPatriotRow = new ArrayList();
        private ArrayList PatriotNotInNemrcRow = new ArrayList();
        private ArrayList PatriotBlankSpanRow = new ArrayList();
        private ArrayList differentValues = new ArrayList();
        private bool valuesOnly = false;
        private string nemrcName1 = "";
        private string nemrcLastName = "";
        private string nemrcName2 = "";
        private string nemrcAddr1 = "";
        private string nemrcAddr2 = "";
        private string nemrcCity = "";
        private string nemrcState = "";
        private string nemrcZip = "";
        private string nemrcValue = "";
        private string patriotName1 = "";
        private string patriotLastName = "";
        private string patriotFirstName1 = "";
        private string patriotAddr1 = "";
        private string patriotAddr2 = "";
        private string patriotCity = "";
        private string patriotState = "";
        private string patriotZip = "";
        private string patriotValue = "";
        private string compareNemrcAddr = "";
        private string comparePatriotAddr = "";

        private int nemrcSpanIndex;
        private int nemrcName1Index;
        private int nemrcName2Index;
        private int nemrcAddr1Index;
        private int nemrcAddr2Index;
        private int nemrcCityIndex;
        private int nemrcStateIndex;
        private int nemrcZipIndex;
        private int nemrcTaxMapIdIndex;
        private int nemrcAcresIndex;
        private int nemrcValueIndex;

        private int patriotSpanIndex;
        private int patriotParcelIdIndex;
        private int patriotLastName1Index;
        private int patriotFirstNameIndex;
        private int patriotAddr1Index;
        private int patriotAddr2Index;
        private int patriotCityIndex;
        private int patriotStateIndex;
        private int patriotZipIndex;
        private int patriotAcresIndex;
        private int patriotValueIndex;

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
        public PatriotOwnershipCheck(System.Windows.Forms.ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
            : base(progressBar, ProgressBarLabel)
        {
        }
        //****************************************************************************************************************************
        public void PatriotOwnershipDifferences(bool workingGrandList, bool valuesOnly, int GrandListYear)
        {
            try
            {
                this.valuesOnly = valuesOnly;
                //GetSalesData();

                worksheetValues1 = SelectInputFile(ExcelClass.NemrcExportsFolder, "NemrcActives", "csv", ColumnName.Span, workingGrandList, GrandListYear);
                if (worksheetValues1 == null)
                {
                    return;
                }
                worksheetValues2 = SelectInputFile(ExcelClass.PatriotExportsFolder, "PatriotActives", "csv", ColumnName.Span, workingGrandList, GrandListYear);
                if (worksheetValues2 == null)
                {
                    return;
                }
                SetColumnIndexes();
                string prefix = (valuesOnly) ? "Values" : "NameAddress";
                string suffix = (workingGrandList) ? "Working" : "AsBilled";
                workbookPath = ExcelClass.ReportsFolder + "\\" + prefix + "Differences" + suffix + ".xlsx";
                CreateNewExcelWorkbook("Difference List");
                WriteHeadings();
                FindDifferencesPatriotNemrcColumn();
                worksheetValues1.CloseWorkbook();
                worksheetValues2.CloseWorkbook();
                //FormatStringCells("C1:C1596");
                CloseOutput();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        public void PatriotInactiveDifferences(bool workingGrandList, int GrandListYear)
        {
            try
            {
                this.valuesOnly = false;
                //GetSalesData();

                worksheetValues1 = SelectInputFile(ExcelClass.NemrcExportsFolder, "NemrcInactives", "csv", ColumnName.Span, workingGrandList, GrandListYear);
                if (worksheetValues1 == null)
                {
                    return;
                }
                worksheetValues2 = SelectInputFile(ExcelClass.PatriotExportsFolder, "PatriotInactives", "csv", ColumnName.Span, workingGrandList, GrandListYear);
                if (worksheetValues2 == null)
                {
                    return;
                }
                SetColumnIndexes();
                string prefix = (valuesOnly) ? "Values" : "NameAddress";
                string suffix = (workingGrandList) ? "Working" : "AsBilled";
                workbookPath = ExcelClass.ReportsFolder + "\\" + prefix + "InactiveDifferences" + suffix + ".xlsx";
                CreateNewExcelWorkbook("Inactive Difference List");
                WriteHeadings();
                FindDifferencesPatriotNemrcColumn(true);
                worksheetValues1.CloseWorkbook();
                worksheetValues2.CloseWorkbook();
                //FormatStringCells("C1:C1596");
                CloseOutput();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //****************************************************************************************************************************
        private void GetSalesData(bool workingGrandList)
        {
            worksheetValues1 = SelectInputFile(ExcelClass.NemrcExportsFolder, "      Patriot Sales Data", "xls", ColumnName.Span, workingGrandList);
            if (worksheetValues1 == null)
            {
                return;
            }
            int rowIndex = 1;
            while (rowIndex <= worksheetValues1.numRows)
            {
                rowIndex++;
                string CertNumber = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, 1);
                string PatriotSpan = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, 13);
                SalesInfo salesInfo = new SalesInfo(CertNumber, PatriotSpan);
                SalesData.Add(salesInfo);
            }
            worksheetValues1.CloseWorkbook();
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
            nemrcAcresIndex = GetColumnNum(worksheetValues1, ColumnName.Acres);
            nemrcValueIndex = GetColumnNum(worksheetValues1, ColumnName.Value);
            nemrcTaxMapIdIndex = GetColumnNum(worksheetValues1, ColumnName.TaxMapId);

            patriotSpanIndex = GetColumnNum(worksheetValues2, ColumnName.Span);
            patriotLastName1Index = GetColumnNum(worksheetValues2, ColumnName.Name1);
            patriotParcelIdIndex = GetColumnNum(worksheetValues2, ColumnName.TaxMapId);
            patriotFirstNameIndex = GetColumnNum(worksheetValues2, ColumnName.FirstName1);
            patriotAddr1Index = GetColumnNum(worksheetValues2, ColumnName.Addr1);
            patriotAddr2Index = GetColumnNum(worksheetValues2, ColumnName.Addr2);
            patriotCityIndex = GetColumnNum(worksheetValues2, ColumnName.City);
            patriotStateIndex = GetColumnNum(worksheetValues2, ColumnName.State);
            patriotZipIndex = GetColumnNum(worksheetValues2, ColumnName.Zip);
            patriotAcresIndex = GetColumnNum(worksheetValues2, ColumnName.Acres);
            patriotValueIndex = GetColumnNum(worksheetValues2, ColumnName.Value);
        }
        //****************************************************************************************************************************
        private void FindDifferencesPatriotNemrcColumn(bool inactives = false)
        {
            this.ProgressBarLabel.Visible = true;
            this.progressBar.Visible = true;
            this.progressBar.Value = 0;
            this.progressBar.Step = 10;
            this.progressBar.Maximum = worksheetValues1.numRows;
            int rowIndex1 = 1;
            int rowIndex2 = 2;  // Skip Headings
            outputIndex = 3;
            while (rowIndex1 <= worksheetValues1.numRows && rowIndex2 <= worksheetValues2.numRows)
            {
                try
                {
                    string grandListSpan = GetCellValue(worksheetValues1.inputWorksheet, rowIndex1, nemrcSpanIndex);
                    string PatriotSpan = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotSpanIndex);
                    char greaterLessthanEqual = GreaterLessthanEqual(grandListSpan, PatriotSpan);
                    if (greaterLessthanEqual == '=')
                    {
                        SpanNumberEqual(rowIndex1, rowIndex2, grandListSpan, inactives);
                        IncrementProgressBar();
                        rowIndex1++;
                        rowIndex2++;
                    }
                    else
                    {
                        if (!inactives && rowIndex2 == 2)
                        {
                            if (MessageBox.Show("First Span Numbers Do Not Equal.  Use Sorted Files", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                return;
                            }
                        }
                        if (String.IsNullOrEmpty(PatriotSpan))
                        {
                            PatriotBlankSpanRow.Add(rowIndex2);
                            rowIndex2++;
                        }
                        else
                        if (greaterLessthanEqual == '<')
                        {
                            NemrcNotInPatriotRow.Add(rowIndex1);
                            IncrementProgressBar();
                            rowIndex1++;
                        }
                        else
                        {
                            PatriotNotInNemrcRow.Add(rowIndex2);
                            rowIndex2++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            if (inactives)
            {
                AddMissingPatriotSpanNumber();
            }
            else
            {
                AddMissingSpanNumber(rowIndex1, worksheetValues1, rowIndex2, worksheetValues2);
            }
        }
        //****************************************************************************************************************************
        private void AddMissingPatriotSpanNumber()
        {
            outputIndex++;
            outputIndex++;
            outputWorksheet.Cells[outputIndex, 1].Value = "Patriot Spans not in NEMRC";
            foreach (int rowIndex2 in PatriotNotInNemrcRow)
            {
                string PatriotSpan = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotSpanIndex);
                string patriotParcelId = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotParcelIdIndex);
                outputIndex++;
                outputWorksheet.Cells[outputIndex, 1].Value = PatriotSpan; 
                outputWorksheet.Cells[outputIndex, 2].Value = patriotParcelId; 
            }
        }
        //****************************************************************************************************************************
        private void AddMissingSpanNumber(int rowIndex1, WorksheetValues worksheetValues1, int rowIndex2, WorksheetValues worksheetValues2)
        {
            outputIndex++;
            foreach (int spanRow in NemrcNotInPatriotRow)
            {
                string grandListSpan = GetCellValue(worksheetValues1.inputWorksheet, spanRow, nemrcSpanIndex);
                string nemrcName1 = GetCellValue(worksheetValues1.inputWorksheet, spanRow, nemrcName1Index);
                outputWorksheet.Cells[outputIndex, 1].Value = grandListSpan;
                outputWorksheet.Cells[outputIndex, 2].Value = nemrcName1;
                outputIndex++;
            }
            while (rowIndex1 <= worksheetValues1.numRows)
            {
                string grandListSpan = GetCellValue(worksheetValues1.inputWorksheet, rowIndex1, nemrcSpanIndex);
                string nemrcName1 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex1, nemrcName1Index);
                outputWorksheet.Cells[outputIndex, 1].Value = grandListSpan;
                outputWorksheet.Cells[outputIndex, 2].Value = nemrcName1;
                outputIndex++;
                rowIndex1++;
            }
            foreach (int spanRow2 in PatriotBlankSpanRow)
            {
                string grandListSpan = GetCellValue(worksheetValues2.inputWorksheet, spanRow2, patriotSpanIndex);
                string patriotName1 = GetCellValue(worksheetValues2.inputWorksheet, spanRow2, patriotLastName1Index);
                outputWorksheet.Cells[outputIndex, 3].Value = grandListSpan;
                outputWorksheet.Cells[outputIndex, 4].Value = patriotName1;
                outputIndex++;
            }
            while (rowIndex2 <= worksheetValues2.numRows)
            {
                string grandListSpan = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotSpanIndex);
                string patriotName1 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotLastName1Index);
                outputWorksheet.Cells[outputIndex, 3].Value = grandListSpan;
                outputWorksheet.Cells[outputIndex, 4].Value = patriotName1;
                outputIndex++;
                rowIndex2++;
            }
            foreach (int spanRow in PatriotBlankSpanRow)
            {
                string parcelId = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotParcelIdIndex);
                string patriotName1 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotLastName1Index);
                outputWorksheet.Cells[outputIndex, 3].Value = parcelId;
                outputWorksheet.Cells[outputIndex, 4].Value = patriotName1;
                outputIndex++;
            }
        }
        //****************************************************************************************************************************
        private void SpanNumberEqual(int rowIndex,
                                     int rowIndex2,
                                     string grandListSpan,
                                     bool inactives)
        {
            string patriotAcres = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotAcresIndex);
            string nemrcAcres = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcAcresIndex);

            if (valuesOnly)
            {
                nemrcValue = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcValueIndex);
                patriotValue = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotValueIndex);
            }
            else
            {
                nemrcName1 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcName1Index);
                nemrcLastName = GetLastName(nemrcName1);
                nemrcName2 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcName2Index);
                nemrcAddr1 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcAddr1Index);
                nemrcAddr2 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcAddr2Index);
                if (String.IsNullOrEmpty(nemrcAddr1))
                {
                    nemrcAddr1 = nemrcAddr2;
                }
                nemrcCity = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcCityIndex);
                nemrcState = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcStateIndex);
                nemrcZip = GetZipCode(GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcZipIndex), inactives);

                patriotName1 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotLastName1Index);
                patriotLastName = GetLastName(patriotName1);
                patriotFirstName1 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotFirstNameIndex);
                patriotAddr1 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotAddr1Index);
                patriotAddr2 = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotAddr2Index);

                patriotCity = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotCityIndex);
                patriotState = GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotStateIndex);
                patriotZip = GetZipCode(GetCellValue(worksheetValues2.inputWorksheet, rowIndex2, patriotZipIndex), inactives);
                AdjustNemrcAddress(ref nemrcAddr1, nemrcAddr2, patriotAddr2);
                compareNemrcAddr = GetAddress(nemrcAddr1);
                comparePatriotAddr = GetAddress(patriotAddr1);
            }
            //****************************************************************************************************************************
            if (PatriotDifferentThanNemrc(nemrcLastName, patriotLastName, nemrcZip, patriotZip, compareNemrcAddr, comparePatriotAddr, nemrcAcres, patriotAcres, nemrcValue, patriotValue))
            {
                string CertNumber = GetSalesInfo(grandListSpan);

                outputWorksheet.Cells[outputIndex, 1].Value = grandListSpan;
                if (valuesOnly)
                {
                    outputWorksheet.Cells[outputIndex, 2].Value = nemrcAcres;
                    outputWorksheet.Cells[outputIndex, 3].Value = patriotAcres;
                    outputWorksheet.Cells[outputIndex, 4].Value = nemrcValue;
                    outputWorksheet.Cells[outputIndex, 5].Value = patriotValue;
                    nemrcName1 = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, nemrcName1Index);
                    outputWorksheet.Cells[outputIndex, 6].Value = nemrcName1;
                }
                else
                {
                    outputWorksheet.Cells[outputIndex, 2].Value = nemrcName1;
                    outputWorksheet.Cells[outputIndex, 3].Value = patriotName1 + " " + patriotFirstName1;
                    outputWorksheet.Cells[outputIndex, 4].Value = nemrcAddr1;
                    outputWorksheet.Cells[outputIndex, 5].Value = patriotAddr1;
                    outputWorksheet.Cells[outputIndex, 6].Value = nemrcAddr2;
                    outputWorksheet.Cells[outputIndex, 7].Value = patriotAddr2;
                    outputWorksheet.Cells[outputIndex, 8].Value = nemrcCity;
                    outputWorksheet.Cells[outputIndex, 9].Value = patriotCity;
                    outputWorksheet.Cells[outputIndex, 10].Value = nemrcState;
                    outputWorksheet.Cells[outputIndex, 11].Value = patriotState;
                    outputWorksheet.Cells[outputIndex, 12].Value = nemrcZip;
                    outputWorksheet.Cells[outputIndex, 13].Value = patriotZip;
                    outputWorksheet.Cells[outputIndex, 14].Value = nemrcAcres;
                    outputWorksheet.Cells[outputIndex, 15].Value = patriotAcres;
                }
                outputIndex++;
            }
        }
        //****************************************************************************************************************************
        private string GetLastName(string nemrcName1)
        {
            int indexOf = nemrcName1.IndexOf(' ');
            if (indexOf > 0)
            {
                return nemrcName1.Remove(indexOf).Trim();
            }
            return nemrcName1.Trim();
        }
        //****************************************************************************************************************************
        private void AdjustNemrcAddress(ref string nemrcAddr1, string nemrcAddr2, string patriotAddr2)
        {
            if (nemrcAddr2.Length > 0 && nemrcAddr2[0] == '#' && (patriotAddr2.Length == 0 || patriotAddr2[0] != '#'))
            {
                nemrcAddr1 += nemrcAddr2;
            }
            if (nemrcAddr2.Length > 0 && nemrcAddr2.ToLower().Contains("apt") && !patriotAddr2.ToLower().Contains("apt"))
            {
                nemrcAddr1 += nemrcAddr2;
            }
            if (nemrcAddr2.Length > 0 && nemrcAddr2.ToLower().Contains("unit") && !patriotAddr2.ToLower().Contains("unit"))
            {
                nemrcAddr1 += nemrcAddr2;
            }

        }
        //****************************************************************************************************************************
        private string GetSalesInfo(string spanNumber)
        {
            foreach (SalesInfo salesInfo in SalesData)
            {
                if (salesInfo.SpanNumber == spanNumber)
                {
                    return salesInfo.CertNumber;
                }
            }
            return "";
        }
        //****************************************************************************************************************************
        private bool PatriotDifferentThanNemrc(string nemrcLastName,
                                               string patriotLastName,
                                               string nemrcZip,
                                               string patriotZip,
                                               string compareNemrcAddr,
                                               string comparePatriotAddr, 
                                               string nemrcAcres, 
                                               string patriotAcres, 
                                               string nemrcValue, 
                                               string patriotValue)
        {
            if (nemrcAcres != patriotAcres)
            {
                return true;
            }
            if (valuesOnly)
            {
                if (nemrcValue != patriotValue)
                {
                    return true;
                }
            }
            else
            {
                if (nemrcLastName.ToUpper() != patriotLastName.ToUpper())
                {
                    return true;
                }
                if (nemrcZip != patriotZip)
                {
                    return true;
                }
                if (compareNemrcAddr != comparePatriotAddr)
                {
                    return true;
                }
            }
            return false;
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
        private string GetZipCode(string zip, bool inactives)
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
                if (intZip == 0 && !inactives)
                {
                    MessageBox.Show("Invalid Zip Code: " + zip);
                }
                zip = String.Format("{0:00000}", intZip);
                return zip;
            }
            catch (Exception)
            {
                if (!inactives)
                {
                    MessageBox.Show("Invalid Zip Code: " + zip);
                }
                return "";
            }
        }
        /*
        //****************************************************************************************************************************
        private int SkipMobileHomes(Worksheet inputWorksheet, int rowIndex, int numCells)
        {
            rowIndex++;
            while (rowIndex <= numCells && inputWorksheet.Cells[rowIndex, 1].Value != null && inputWorksheet.Cells[rowIndex, 1].Value.Contains("MH."))
            {
                rowIndex++;
            }
            while (rowIndex <= numCells && inputWorksheet.Cells[rowIndex, 1].Value != null && inputWorksheet.Cells[rowIndex, 1].Value.Contains("J-9.18"))
            {
                rowIndex++;
            }
            while (rowIndex <= numCells && inputWorksheet.Cells[rowIndex, 1].Value != null && inputWorksheet.Cells[rowIndex, 1].Value.Contains("F-31."))
            {
                rowIndex++;
            }
            while (rowIndex <= numCells && inputWorksheet.Cells[rowIndex, 1].Value != null && inputWorksheet.Cells[rowIndex, 1].Value.Contains("F-36."))
            {
                rowIndex++;
            }
            while (rowIndex <= numCells && inputWorksheet.Cells[rowIndex, 1].Value != null && inputWorksheet.Cells[rowIndex, 1].Value.Contains("O-1"))
            {
                rowIndex++;
            }
            return rowIndex;
        }
        //****************************************************************************************************************************
        private int SkipDuplicateIds(int rowIndex2)
        {
            string previousValue = worksheetValues2.inputWorksheet.Cells[rowIndex2, 1].Value.ToString();
            rowIndex2++;
            while (rowIndex2 <= worksheetValues2.numRows && worksheetValues2.inputWorksheet.Cells[rowIndex2, 1].Value.ToString() == previousValue)
            {
                rowIndex2++;
            }
            return rowIndex2;
        }*/
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
