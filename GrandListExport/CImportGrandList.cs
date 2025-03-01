using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Data;
using System.Windows.Forms;
using SQL_Library;

namespace HistoricJamaica
{
    public class CImportGrandList : CImport
    {
        private enum NemrcColumns
        {
            ParcelId = 1,
            SubParcelId = 2,
            Name1 = 3,
            Name2 = 4,
            Address1 = 5,
            Address2 = 6,
            City = 7,
            State = 8,
            ZipCode = 9,
            LocationA = 10,
            LocationB = 11,
            LocationC = 12,
            StreetNum = 13,
            StreetName = 14,
            TaxMapID = 15,
            PropertyDescription = 16,
            Owner = 17,
            DateHomestead = 18
        }
        private EPPlus epPlus;
        private ArrayList taxMapList = new ArrayList();
        private ArrayList taxMapIdNoInNewFile = new ArrayList();
        private DataTable grandListTbl;
        private DataTable modernRoadValueTbl;
        //****************************************************************************************************************************
        public CImportGrandList(CSql Sql, string sDataDirectory)
            : base(Sql, sDataDirectory)
        {
            try
            {
                using (epPlus = new EPPlus())
                {
                    grandListTbl = SQL.GetAllGrandList();
                    modernRoadValueTbl = SQL.GetAllModernRoadValues();
                    GetActivesInactives();
                    GetActivesInactives();
                    CheckAllInactiveGrandlistRecords(grandListTbl);
                    SQL.UpdateInsertDeleteGrandList(grandListTbl);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //****************************************************************************************************************************
        private void GetActivesInactives()
        {
            string filename;
            GetExcelInputFile(@"c:\Reports\", out filename);
            if (!string.IsNullOrEmpty(filename))
            {
                GetGrandListRecords(filename);
            }
        }
        //****************************************************************************************************************************
        private void CheckAllInactiveGrandlistRecords(DataTable grandListTbl)
        {
            DataTable buildingTbl = SQL.GetAllBuildings();
            foreach (DataRow grandListRow in grandListTbl.Rows)
            {
                string grandListId = grandListRow[U.GrandListID_col].ToString();
                string taxMapId = grandListRow[U.TaxMapID_col].ToString();
                if (!String.IsNullOrEmpty(taxMapId) && !IsTaxMapIdInList(taxMapId))
                {
                    string selectStatement = U.BuildingGrandListID_col + "=" + grandListId;
                    DataRow[] foundRows = buildingTbl.Select(selectStatement);
                    if (foundRows.Length != 0)
                    {
                        taxMapIdNoInNewFile.Add(taxMapId);
                        MessageBox.Show("TaxMapID in Database not in new File: ", taxMapId);
                    }
                    else
                    {
                        grandListRow.Delete();
                    }
                }
            }
        }
        //****************************************************************************************************************************
        private bool IsTaxMapIdInList(string taxMapId)
        {
            foreach (string taxMapIdInList in taxMapList)
            {
                if (taxMapId == taxMapIdInList)
                {
                    return true;
                }
            }
            return false;
        }
        //****************************************************************************************************************************
        private void GetGrandListRecords(string filename)
        {
            epPlus.OpenWithEPPlus(filename);
            if (string.IsNullOrEmpty(filename))
            {
                return;
            }
            int rowIndex = 2;
            while (rowIndex <= epPlus.numRows)
            {
                try
                {
                    AddRecordToDatabase(rowIndex);
                    rowIndex++;
                }
                catch (Exception ex)
                {
                    string message = "Row: " + rowIndex + " - " + ex.Message;
                    throw new Exception(message);
                }
            }
        }
        //****************************************************************************************************************************
        private void AddRecordToDatabase(int rowIndex)
        {
            if (rowIndex == 1200)
            {
            }
            string parcelId = epPlus.GetCellValue(rowIndex, (int) NemrcColumns.ParcelId).ToString();
            string SubParcelId = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.SubParcelId).ToString();
            string Name1 = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.Name1).ToString();
            string Name2 = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.Name2).ToString();
            string Address1 = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.Address1).ToString();
            string Address2 = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.Address2).ToString();
            string City = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.City).ToString();
            string State = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.State).ToString();
            string ZipCode = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.ZipCode).ToString();
            string LocationA = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.LocationA).ToString();
            string LocationB = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.LocationB).ToString();
            string LocationC = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.LocationB).ToString();
            int StreetNum = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.StreetNum).ToInt();
            string StreetName = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.StreetName).ToString();
            string TaxMapID = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.TaxMapID).ToString();
            string PropertyDescription = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.PropertyDescription).ToString();
            string Owner = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.Owner).ToString();
            string DateHomestead = epPlus.GetCellValue(rowIndex, (int)NemrcColumns.DateHomestead).ToString();
            if (TaxMapID == "F-L6.1")
            {
            }
            if (ZipCode.Length == 4)
            {
                ZipCode = ZipCode.Insert(0, "0");
            }
            if (ExcludedProperty(Name1, TaxMapID))
            {
                return;
            }
            string selectStatement = U.TaxMapID_col + " = '" + TaxMapID + "'";
            DataRow[] foundRows = grandListTbl.Select(selectStatement);
            if (foundRows.Length == 1)
            {
                SetNewValueIfDifferent(foundRows[0], U.Name1_col, Name1);
                SetNewValueIfDifferent(foundRows[0], U.Name2_col, Name2);
                SetNewValueIfDifferent(foundRows[0], U.StreetName_col, StreetName);
                if (foundRows[0][U.WhereOwnerLiveID_col].ToChar() != Owner[0])
                {
                    foundRows[0][U.WhereOwnerLiveID_col] = Owner[0];
                }
                if (foundRows[0][U.StreetNum_col].ToInt() != StreetNum)
                {
                    foundRows[0][U.StreetNum_col] = StreetNum;
                }
                int oldId = foundRows[0][U.BuildingRoadValueID_col].ToInt();
                int newId = GetModernRoadValue(foundRows[0][U.StreetName_col].ToString(), StreetNum);
                if (newId != oldId)
                {
                    if (TaxMapID == "O-L35" || TaxMapID == "O-L36") // Williams Rd
                    {
                        newId = oldId;
                    }
                    else if (oldId == 189) // Hemlock City Lane
                    {
                        newId = oldId;
                    }
                    else
                    {
                        if (newId == 0)
                        {
                        }
                        else if (StreetNum == 0)
                        {
                        }
                        else if (oldId == 0)
                        {
                        }
                        else
                        {
                        }
                        foundRows[0][U.BuildingRoadValueID_col] = newId;
                    }
                }
            }
            else if (foundRows.Length > 1)
            {
                MessageBox.Show("Multiple Grand List Entries: " + TaxMapID);
            }
            else if (String.IsNullOrEmpty(StreetName))
            {
            }
            else
            {
                DataRow grandListNewRow = grandListTbl.NewRow();
                grandListNewRow[U.GrandListID_col] = 0;
                grandListNewRow[U.TaxMapID_col] = TaxMapID;
                grandListNewRow[U.Name1_col] = CapitalizeLowerCase(Name1.ToLower());
                grandListNewRow[U.Name2_col] = CapitalizeLowerCase(Name2.ToLower());
                grandListNewRow[U.StreetName_col] = CapitalizeLowerCase(StreetName.ToLower());
                grandListNewRow[U.StreetNum_col] = StreetNum;
                grandListNewRow[U.WhereOwnerLiveID_col] = Owner[0];
                grandListNewRow[U.BuildingRoadValueID_col] = GetModernRoadValue(grandListNewRow[U.StreetName_col].ToString(), StreetNum); 
                grandListTbl.Rows.Add(grandListNewRow);
            }
            taxMapList.Add(TaxMapID);
        }
        //****************************************************************************************************************************
        private int GetModernRoadValue(string streetName, int streetNum)
        {
            if (String.IsNullOrEmpty(streetName))
            {
                return 0;
            }
            streetName = streetName.Replace("Mtn", "Mountain");
            streetName = streetName.Replace("Old Rte 8", "Old Route 8");
            streetName = streetName.Replace("Olde", "Old");
            streetName = streetName.Replace("`", "");
            streetName = streetName.Replace("'", "");
            if (!SpecialRoadValue(ref streetName, streetNum))
            {
                streetName = ReplaceAbbrevation(streetName, " Rd ", " Road");
                streetName = ReplaceAbbrevation(streetName, " St ", " Street");
                streetName = ReplaceAbbrevation(streetName, " Ln ", " Lane");
                streetName = ReplaceAbbrevation(streetName, " Dr ", " Drive");
            }
            string selectStatement = U.ModernRoadValueValue_col + " = '" + streetName + "'";
            DataRow[] foundRows = modernRoadValueTbl.Select(selectStatement);
            if (foundRows.Length == 0)
            {
                return 0;
            }
            return foundRows[0][U.ModernRoadValueID_col].ToInt();
        }
        //****************************************************************************************************************************
        private bool SpecialRoadValue(ref string streetName, int streetNum)
        {
            int indexOf = streetName.IndexOf(" Ln A");
            if (indexOf > 0)
            {
                streetName = streetName.Substring(0, indexOf) + " Lane";
                return true;
            }
            indexOf = streetName.IndexOf(" Ln B");
            if (indexOf > 0)
            {
                streetName = streetName.Substring(0, indexOf) + " Lane";
                return true;
            }
            indexOf = streetName.IndexOf(" Ln C");
            if (indexOf > 0)
            {
                streetName = streetName.Substring(0, indexOf) + " Lane";
                return true;
            }
            if (streetName.ToUpper() == "PIKES FALLS RD" && streetNum < 300)
            {
                streetName = "Pikes Falls-Mechanic Street";
                return true;
            }
            if (streetName.ToUpper().Contains("WEST HILL"))
            {
                streetName = streetName.Replace(" Rd ", " Road ");
                return true;
            }
            if (streetName.Substring(0, 4).ToUpper() == "VT R")
            {
                streetName = SubstituteVtRoute(streetName, streetNum);
                return true;
            }
            return false;
        }
        //****************************************************************************************************************************
        private string SubstituteVtRoute(string streetName, int streetNum)
        {
            int indexOf = streetName.ToUpper().IndexOf("VT ROUTE ");
            if (indexOf >= 0)
            {
                streetName = "Route " + streetName.Remove(0, 9);
            }
            indexOf = streetName.ToUpper().IndexOf("VT RTE ");
            if (indexOf >= 0)
            {
                streetName = "Route " + streetName.Remove(0, 7);
            }
            indexOf = streetName.ToUpper().IndexOf("VT RT ");
            if (indexOf >= 0)
            {
                streetName = "Route " + streetName.Remove(0, 6);
            }
            if (streetName.Contains("30"))
            {
                indexOf = streetName.IndexOf("@");
                if (indexOf > 0)
                {
                    streetName = streetName.Replace("Rd", "Road");
                }
                else if (streetNum < 3458)
                {
                    streetName += " South";
                }
                else if (streetNum > 8550)
                {
                    char lastCharInString = streetName[streetName.Length - 1];
                    if (lastCharInString != '0')
                    {
                        streetName = streetName.Replace("30 A", "30");
                        streetName = streetName.Replace("30 B", "30");
                        streetName = streetName.Replace("30 C", "30");
                        streetName = streetName.Replace("30 D", "30");
                        streetName = streetName.Replace("30 E", "30");
                        streetName = streetName.Replace("30 F", "30");
                        streetName = streetName.Replace("30 G", "30");
                        streetName = streetName.Replace("30 H", "30");
                        streetName = streetName.Replace("30 I", "30");
                    }
                    streetName += "-Rawsonville";
                }
                else if (streetNum > 3924)
                {
                    streetName += " North";
                }
                else
                {
                    streetName = "Main Street";
                }
            }
            else
            {
                char lastCharInString = streetName[streetName.Length - 1];
                if (lastCharInString == 'S' || lastCharInString == 'N')
                {
                    streetName = streetName.Replace("100 S", "100 South");
                    streetName = streetName.Replace("100 N", "100 North");
                }

                else
                {
                    streetName = streetName.Replace("NORTH", "North");
                    streetName = streetName.Replace("SOUTH", "South");
                }
            }
            return streetName;
        }
        //****************************************************************************************************************************
        private string ReplaceAbbrevation(string streetName, string abbrevation, string fullWord)
        {
            streetName = streetName + " ";
            int indexOf = streetName.IndexOf(abbrevation);
            if (indexOf > 0)
            {
                return streetName.Replace(abbrevation, fullWord).Trim();
            }
            return streetName.Trim();
        }
        //****************************************************************************************************************************
        private void SetNewValueIfDifferent(DataRow grandListRow, string col, string newValue)
        {
            newValue = newValue.Replace("  ", " ");
            if (!String.IsNullOrEmpty(newValue))
            {
                newValue = CapitalizeLowerCase(newValue.ToLower());
            }
            if (grandListRow[col].ToString() != newValue)
            {
                grandListRow[col] = newValue;
            }
        }
        //****************************************************************************************************************************
        private string CapitalizeLowerCase(string str)
        {
            string returnString = "";
            string[] words = str.Split(' ');
            foreach (string word in words)
            {
                if (!String.IsNullOrEmpty(word))
                {
                    returnString += CapitalizeFirstChar(word) + " ";
                }
            }
            return returnString.Trim();
        }
        //****************************************************************************************************************************
        private string CapitalizeFirstChar(string str)
        {
            return char.ToUpper(str[0]) + str.Substring(1);
        }
        //****************************************************************************************************************************
        private bool ExcludedProperty(string name, string taxMapId)
        {
            if (name.ToLower().Contains("vermont state of"))
            {
                return true;
            }
            if (String.IsNullOrEmpty(taxMapId))
            {
                return true;
            }
            return false;
        }
        //****************************************************************************************************************************
    }
}
