namespace GrandListExport
{
    partial class WorkingOrAsBilledForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button_Working = new System.Windows.Forms.Button();
            this.button_AsBilled = new System.Windows.Forms.Button();
            this.button_Cancel = new System.Windows.Forms.Button();
            this.Year_textBox = new System.Windows.Forms.TextBox();
            this.Year_label = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button_Working
            // 
            this.button_Working.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button_Working.Location = new System.Drawing.Point(42, 161);
            this.button_Working.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.button_Working.Name = "button_Working";
            this.button_Working.Size = new System.Drawing.Size(151, 44);
            this.button_Working.TabIndex = 0;
            this.button_Working.Text = "Working";
            this.button_Working.UseVisualStyleBackColor = true;
            this.button_Working.Click += new System.EventHandler(this.WorkingClicked);
            // 
            // button_AsBilled
            // 
            this.button_AsBilled.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button_AsBilled.Location = new System.Drawing.Point(322, 161);
            this.button_AsBilled.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.button_AsBilled.Name = "button_AsBilled";
            this.button_AsBilled.Size = new System.Drawing.Size(151, 44);
            this.button_AsBilled.TabIndex = 1;
            this.button_AsBilled.Text = "As Billed";
            this.button_AsBilled.UseVisualStyleBackColor = true;
            this.button_AsBilled.Click += new System.EventHandler(this.AsBilledClicked);
            // 
            // button_Cancel
            // 
            this.button_Cancel.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button_Cancel.Location = new System.Drawing.Point(594, 161);
            this.button_Cancel.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.button_Cancel.Name = "button_Cancel";
            this.button_Cancel.Size = new System.Drawing.Size(151, 44);
            this.button_Cancel.TabIndex = 2;
            this.button_Cancel.Text = "Cancel";
            this.button_Cancel.UseVisualStyleBackColor = true;
            this.button_Cancel.Click += new System.EventHandler(this.CancelClicked);
            // 
            // Year_textBox
            // 
            this.Year_textBox.Location = new System.Drawing.Point(386, 91);
            this.Year_textBox.Name = "Year_textBox";
            this.Year_textBox.Size = new System.Drawing.Size(87, 31);
            this.Year_textBox.TabIndex = 3;
            this.Year_textBox.Visible = false;
            // 
            // Year_label
            // 
            this.Year_label.AutoSize = true;
            this.Year_label.Location = new System.Drawing.Point(317, 94);
            this.Year_label.Name = "Year_label";
            this.Year_label.Size = new System.Drawing.Size(58, 25);
            this.Year_label.TabIndex = 4;
            this.Year_label.Text = "Year";
            this.Year_label.Visible = false;
            // 
            // WorkingOrAsBilledForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(804, 269);
            this.Controls.Add(this.Year_label);
            this.Controls.Add(this.Year_textBox);
            this.Controls.Add(this.button_Cancel);
            this.Controls.Add(this.button_AsBilled);
            this.Controls.Add(this.button_Working);
            this.Location = new System.Drawing.Point(400, 300);
            this.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.Name = "WorkingOrAsBilledForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Working Or As Billed Grand List";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_Working;
        private System.Windows.Forms.Button button_AsBilled;
        private System.Windows.Forms.Button button_Cancel;
        private System.Windows.Forms.TextBox Year_textBox;
        private System.Windows.Forms.Label Year_label;
    }
}