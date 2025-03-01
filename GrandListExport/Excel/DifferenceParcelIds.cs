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
    public class DifferenceListOfParcelIds : ExcelClass 
    {
        private WorksheetValues worksheetValues1;

        public DifferenceListOfParcelIds(ProgressBar progressBar, System.Windows.Forms.Label ProgressBarLabel)
            : base(progressBar, ProgressBarLabel)
        {
        }

        public void CheckForDuplicateParcelIds(bool workingGrandList)
        {
            try
            {
                worksheetValues1 = SelectInputFile(ExcelClass.NemrcExportsFolder, "NEMRC_GrandListSortedByTaxMapId", "xls", ColumnName.TaxMapId, workingGrandList);
                if (worksheetValues1 != null)
                {
                    CheckForDuplicates();
                }
                worksheetValues1.CloseWorkbook();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CheckForDuplicates()
        {
            progressBar.Visible = true;
            progressBar.Value = 0;
            progressBar.Step = 1;
            progressBar.Maximum = worksheetValues1.numRows;
            int rowIndex = 2;
            string previoustaxMapId = "";
            try
            {
                while (rowIndex < worksheetValues1.numRows)
                {
                    string taxMapId = GetCellValue(worksheetValues1.inputWorksheet, rowIndex, worksheetValues1.GetNemrcColumnNum(ColumnName.TaxMapId));
                    if (String.IsNullOrEmpty(taxMapId))
                    {
                        MessageBox.Show("Empty Cell");
                    }
                    if (taxMapId == previoustaxMapId)
                    {
                        MessageBox.Show("Duplicate taxMapId: " + taxMapId);
                    }
                    previoustaxMapId = taxMapId;
                    rowIndex++;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
