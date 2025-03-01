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
    public partial class WorkingOrAsBilledForm : Form
    {
        string GrandList = "";
        public int year = 0;
        public WorkingOrAsBilledForm(bool GetYear)
        {
            InitializeComponent();
            if (GetYear)
            {
                Year_textBox.Visible = true;
                Year_label.Visible = true;
                Year_textBox.Focus();
            }
        }
        public string GetGrandList()
        {
            return GrandList;
        }
        public void WorkingClicked(object sender, EventArgs e)
        {
            GrandList = "Working";
            this.Close();
        }
        public void AsBilledClicked(object sender, EventArgs e)
        {
            GrandList = "AsBilled";
            if (String.IsNullOrEmpty(Year_textBox.Text) || Year_textBox.Text.Length != 4)
            {
                MessageBox.Show("Please Enter a Valid Year");
            }
            else
            {
                year = int.Parse(Year_textBox.Text);
                this.Close();
            }
        }
        public void CancelClicked(object sender, EventArgs e)
        {
            GrandList = "";
            this.Close();
        }
    }
}
