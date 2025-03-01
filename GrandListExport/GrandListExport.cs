using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GrandListExport
{
    public partial class GrandListExport : Form
    {
        private int GrandListYear = 0;
        public GrandListExport()
        {
            InitializeComponent();
            this.button_FloodZoneOwnership.Visible = true;
        }
        public void CombineButton_Click(object sender, EventArgs e)
        {
            if (CheckRadioButtons(true))
            {
                using (CombineActivesInactives excelWorkbook = new CombineActivesInactives(this.progressBar, ProgressBarLabel))        // Combine Active/Inactives for Arcmap
                {
                    excelWorkbook.Combine(Working_radioButton.Checked, GrandListYear);
                }
            }
        }
        public void CreateNewParcelIds_Click(object sender, EventArgs e)
        {
            using (CreateListOfParcelIds excelWorkbook = new CreateListOfParcelIds(this.progressBar, ProgressBarLabel))             // create new parcel Ids for Conversion
            {
                excelWorkbook.CreateList(Working_radioButton.Checked);
            }
        }
        public void GrandListExport_Click(object sender, EventArgs e)
        {
            //Working_radioButton.Checked = true;
            if (CheckRadioButtons(true))
            {
                using (GrandListExportNemrc excelWorkbook = new GrandListExportNemrc(this.progressBar, ProgressBarLabel))
                {
                    excelWorkbook.CreateList(Working_radioButton.Checked, GrandListYear);
                }
            }
        }
        public void FloodZoneOwnershipCheck_Click(object sender, EventArgs e)
        {
            Working_radioButton.Checked = true;
            {
                using (FloodZoneOwnership excelWorkbook = new FloodZoneOwnership(this.progressBar, ProgressBarLabel))
                {
                    excelWorkbook.FloodZoneOwnershipSetup(Working_radioButton.Checked, false);
                }
            }
        }
        public void PatriotOwnershipCheck_Click(object sender, EventArgs e)
        {
            if (CheckRadioButtons(true))
            {
                using (PatriotOwnershipCheck excelWorkbook = new PatriotOwnershipCheck(this.progressBar, ProgressBarLabel))
                {
                    excelWorkbook.PatriotOwnershipDifferences(Working_radioButton.Checked, false, GrandListYear);
                }
            }
        }
        public void PatriotInactiveOwnershipCheck_Click(object sender, EventArgs e)
        {
            if (CheckRadioButtons(true))
            {
                using (PatriotOwnershipCheck excelWorkbook = new PatriotOwnershipCheck(this.progressBar, ProgressBarLabel))
                {
                    excelWorkbook.PatriotInactiveDifferences(Working_radioButton.Checked, GrandListYear);
                }
            }
        }
        public void PatriotOwnershipValueCheck_Click(object sender, EventArgs e)
        {
            if (CheckRadioButtons(true))
            {
                using (PatriotOwnershipCheck excelWorkbook = new PatriotOwnershipCheck(this.progressBar, ProgressBarLabel))
                {
                    excelWorkbook.PatriotOwnershipDifferences(Working_radioButton.Checked, true, GrandListYear);
                }
            }
        }
        public void MatchContiguous_Click(object sender, EventArgs e)
        {
            if (CheckRadioButtons(true))
            {
                using (MatchContiguousParcels excelWorkbook = new MatchContiguousParcels(this.progressBar, ProgressBarLabel))        // Combine Active/Inactives for Arcmap
                {
                    excelWorkbook.MatchContiguous(Working_radioButton.Checked, GrandListYear);
                }
            }
        }
        public void ActiveDifferenceButton_Click(object sender, EventArgs e)
        {
            if (CheckRadioButtons(true))
            {
                using (NemrcOwnershipCheck excelWorkbook = new NemrcOwnershipCheck(this.progressBar, ProgressBarLabel))
                {
                    excelWorkbook.NemrcTaxMapDifferences(true, Working_radioButton.Checked, GrandListYear);
                }
            }
        }
        public void InactiveDifferenceButton_Click(object sender, EventArgs e)
        {
            if (CheckRadioButtons(true))
            {
                using (NemrcOwnershipCheck excelWorkbook = new NemrcOwnershipCheck(this.progressBar, ProgressBarLabel))
                {
                    excelWorkbook.NemrcTaxMapDifferences(false, Working_radioButton.Checked, GrandListYear);
                }
            }
        }
        public void DuplicateParcelIds_Click(object sender, EventArgs e)
        {
            using (DifferenceListOfParcelIds excelWorkbook = new DifferenceListOfParcelIds(this.progressBar, ProgressBarLabel))
            {
                excelWorkbook.CheckForDuplicateParcelIds(Working_radioButton.Checked);
            }
        }
        public void MergeTwoWorksheets_Click(object sender, EventArgs e)
        {
            using (MergeTwoWorksheets excelWorkbook = new MergeTwoWorksheets(this.progressBar, ProgressBarLabel))
            {
                excelWorkbook.MergeWorksheets(true, Working_radioButton.Checked);
            }
        }
        private bool CheckRadioButtons(bool getYear=false)
        {
            using (WorkingOrAsBilledForm workingOrAsBilledForm = new WorkingOrAsBilledForm(getYear))
            {
                workingOrAsBilledForm.ShowDialog();
                GrandListYear = workingOrAsBilledForm.year;
                string grandList = workingOrAsBilledForm.GetGrandList();
                if (string.IsNullOrEmpty(grandList))
                {
                    return false;
                }
                if (grandList.ToLower() == "working")
                {
                    Working_radioButton.Checked = true; 
                    return true;
                }
                AsBilled_radioButton.Checked = true; 
                return true;
            }
        }
        private bool CheckRadioButtonsOld()
        {
            DialogResult result = MessageBox.Show("Yes: Working Grand List     No: As Billed     Cancel: Cancel Report", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            switch (result)
            {
                case DialogResult.Yes: Working_radioButton.Checked = true; return true;
                case DialogResult.No: AsBilled_radioButton.Checked = true; return true;
                default: return false;
            }
        }
    }
}
