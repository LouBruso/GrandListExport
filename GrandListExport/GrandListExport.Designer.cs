namespace GrandListExport
{
    partial class GrandListExport
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
            this.CombineButton = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.ProgressBarLabel = new System.Windows.Forms.Label();
            this.InactiveDifferenceButton = new System.Windows.Forms.Button();
            this.buttonCreateNewParcelIds = new System.Windows.Forms.Button();
            this.button_DuplicateParcelIds = new System.Windows.Forms.Button();
            this.buttonPatriotOwnershipCheck = new System.Windows.Forms.Button();
            this.buttonMatchContiguous = new System.Windows.Forms.Button();
            this.button_FloodZoneOwnership = new System.Windows.Forms.Button();
            this.buttonPatriotOwnershipValueCheck = new System.Windows.Forms.Button();
            this.ActiveDifferenceButton = new System.Windows.Forms.Button();
            this.Working_radioButton = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.AsBilled_radioButton = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonPatriotInactiveOwnershipCheck = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // CombineButton
            // 
            this.CombineButton.Location = new System.Drawing.Point(38, 66);
            this.CombineButton.Name = "CombineButton";
            this.CombineButton.Size = new System.Drawing.Size(230, 23);
            this.CombineButton.TabIndex = 0;
            this.CombineButton.Text = "Combine Actives and Inactives For Tax Maps";
            this.CombineButton.UseVisualStyleBackColor = true;
            this.CombineButton.Click += new System.EventHandler(this.CombineButton_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(38, 422);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(230, 23);
            this.progressBar.TabIndex = 1;
            this.progressBar.Visible = false;
            // 
            // ProgressBarLabel
            // 
            this.ProgressBarLabel.AutoSize = true;
            this.ProgressBarLabel.Location = new System.Drawing.Point(35, 406);
            this.ProgressBarLabel.Name = "ProgressBarLabel";
            this.ProgressBarLabel.Size = new System.Drawing.Size(37, 13);
            this.ProgressBarLabel.TabIndex = 2;
            this.ProgressBarLabel.Text = "Active";
            this.ProgressBarLabel.Visible = false;
            // 
            // InactiveDifferenceButton
            // 
            this.InactiveDifferenceButton.Location = new System.Drawing.Point(38, 326);
            this.InactiveDifferenceButton.Name = "InactiveDifferenceButton";
            this.InactiveDifferenceButton.Size = new System.Drawing.Size(230, 23);
            this.InactiveDifferenceButton.TabIndex = 3;
            this.InactiveDifferenceButton.Text = "Nemrc TaxMap Inactive Differences";
            this.InactiveDifferenceButton.UseVisualStyleBackColor = true;
            this.InactiveDifferenceButton.Click += new System.EventHandler(this.InactiveDifferenceButton_Click);
            // 
            // buttonCreateNewParcelIds
            // 
            this.buttonCreateNewParcelIds.Location = new System.Drawing.Point(38, 12);
            this.buttonCreateNewParcelIds.Name = "buttonCreateNewParcelIds";
            this.buttonCreateNewParcelIds.Size = new System.Drawing.Size(230, 23);
            this.buttonCreateNewParcelIds.TabIndex = 4;
            this.buttonCreateNewParcelIds.Text = "Create New Parcel Ids";
            this.buttonCreateNewParcelIds.UseVisualStyleBackColor = true;
            this.buttonCreateNewParcelIds.Visible = false;
            this.buttonCreateNewParcelIds.Click += new System.EventHandler(this.CreateNewParcelIds_Click);
            // 
            // button_DuplicateParcelIds
            // 
            this.button_DuplicateParcelIds.Location = new System.Drawing.Point(38, 539);
            this.button_DuplicateParcelIds.Name = "button_DuplicateParcelIds";
            this.button_DuplicateParcelIds.Size = new System.Drawing.Size(230, 23);
            this.button_DuplicateParcelIds.TabIndex = 5;
            this.button_DuplicateParcelIds.Text = "Merge Two Worksheets";
            this.button_DuplicateParcelIds.UseVisualStyleBackColor = true;
            this.button_DuplicateParcelIds.Visible = false;
            this.button_DuplicateParcelIds.Click += new System.EventHandler(this.MergeTwoWorksheets_Click);
            // 
            // buttonPatriotOwnershipCheck
            // 
            this.buttonPatriotOwnershipCheck.Location = new System.Drawing.Point(38, 106);
            this.buttonPatriotOwnershipCheck.Name = "buttonPatriotOwnershipCheck";
            this.buttonPatriotOwnershipCheck.Size = new System.Drawing.Size(230, 23);
            this.buttonPatriotOwnershipCheck.TabIndex = 6;
            this.buttonPatriotOwnershipCheck.Text = "Patriot Nemrc Name/Address Diffs";
            this.buttonPatriotOwnershipCheck.UseVisualStyleBackColor = true;
            this.buttonPatriotOwnershipCheck.Click += new System.EventHandler(this.PatriotOwnershipCheck_Click);
            // 
            // buttonMatchContiguous
            // 
            this.buttonMatchContiguous.Location = new System.Drawing.Point(38, 246);
            this.buttonMatchContiguous.Name = "buttonMatchContiguous";
            this.buttonMatchContiguous.Size = new System.Drawing.Size(230, 23);
            this.buttonMatchContiguous.TabIndex = 7;
            this.buttonMatchContiguous.Text = "TaxMap Match Inactive Contiguous IDs";
            this.buttonMatchContiguous.UseVisualStyleBackColor = true;
            this.buttonMatchContiguous.Click += new System.EventHandler(this.MatchContiguous_Click);
            // 
            // button_FloodZoneOwnership
            // 
            this.button_FloodZoneOwnership.Location = new System.Drawing.Point(38, 366);
            this.button_FloodZoneOwnership.Name = "button_FloodZoneOwnership";
            this.button_FloodZoneOwnership.Size = new System.Drawing.Size(230, 23);
            this.button_FloodZoneOwnership.TabIndex = 14;
            this.button_FloodZoneOwnership.Text = "Grand List Export";
            this.button_FloodZoneOwnership.UseVisualStyleBackColor = true;
            this.button_FloodZoneOwnership.Click += new System.EventHandler(this.GrandListExport_Click);
            // 
            // buttonPatriotOwnershipValueCheck
            // 
            this.buttonPatriotOwnershipValueCheck.Location = new System.Drawing.Point(38, 146);
            this.buttonPatriotOwnershipValueCheck.Name = "buttonPatriotOwnershipValueCheck";
            this.buttonPatriotOwnershipValueCheck.Size = new System.Drawing.Size(230, 23);
            this.buttonPatriotOwnershipValueCheck.TabIndex = 8;
            this.buttonPatriotOwnershipValueCheck.Text = "Patriot Nemrc Value Differences";
            this.buttonPatriotOwnershipValueCheck.UseVisualStyleBackColor = true;
            this.buttonPatriotOwnershipValueCheck.Click += new System.EventHandler(this.PatriotOwnershipValueCheck_Click);
            // 
            // ActiveDifferenceButton
            // 
            this.ActiveDifferenceButton.Location = new System.Drawing.Point(38, 286);
            this.ActiveDifferenceButton.Name = "ActiveDifferenceButton";
            this.ActiveDifferenceButton.Size = new System.Drawing.Size(230, 23);
            this.ActiveDifferenceButton.TabIndex = 9;
            this.ActiveDifferenceButton.Text = "Nemrc TaxMap Active Differences";
            this.ActiveDifferenceButton.UseVisualStyleBackColor = true;
            this.ActiveDifferenceButton.Click += new System.EventHandler(this.ActiveDifferenceButton_Click);
            // 
            // Working_radioButton
            // 
            this.Working_radioButton.AutoSize = true;
            this.Working_radioButton.Location = new System.Drawing.Point(6, 29);
            this.Working_radioButton.Name = "Working_radioButton";
            this.Working_radioButton.Size = new System.Drawing.Size(65, 17);
            this.Working_radioButton.TabIndex = 10;
            this.Working_radioButton.TabStop = true;
            this.Working_radioButton.Text = "Working";
            this.Working_radioButton.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.AsBilled_radioButton);
            this.groupBox1.Controls.Add(this.Working_radioButton);
            this.groupBox1.Location = new System.Drawing.Point(38, 468);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(230, 65);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Grand List";
            // 
            // AsBilled_radioButton
            // 
            this.AsBilled_radioButton.AutoSize = true;
            this.AsBilled_radioButton.Location = new System.Drawing.Point(139, 29);
            this.AsBilled_radioButton.Name = "AsBilled_radioButton";
            this.AsBilled_radioButton.Size = new System.Drawing.Size(65, 17);
            this.AsBilled_radioButton.TabIndex = 11;
            this.AsBilled_radioButton.TabStop = true;
            this.AsBilled_radioButton.Text = "As Billed";
            this.AsBilled_radioButton.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(81, 224);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(146, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "Nemrc - Tax Map Differences";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(80, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(134, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Nemrc - Patriot Differences";
            // 
            // buttonPatriotInactiveOwnershipCheck
            // 
            this.buttonPatriotInactiveOwnershipCheck.Location = new System.Drawing.Point(38, 186);
            this.buttonPatriotInactiveOwnershipCheck.Name = "buttonPatriotInactiveOwnershipCheck";
            this.buttonPatriotInactiveOwnershipCheck.Size = new System.Drawing.Size(230, 23);
            this.buttonPatriotInactiveOwnershipCheck.TabIndex = 15;
            this.buttonPatriotInactiveOwnershipCheck.Text = "Patriot Inactive Name/Address Diffs";
            this.buttonPatriotInactiveOwnershipCheck.UseVisualStyleBackColor = true;
            this.buttonPatriotInactiveOwnershipCheck.Click += new System.EventHandler(this.PatriotInactiveOwnershipCheck_Click);
            //
            // GrandListExport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(331, 574);
            this.Controls.Add(this.buttonPatriotInactiveOwnershipCheck);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.ActiveDifferenceButton);
            this.Controls.Add(this.buttonPatriotOwnershipValueCheck);
            this.Controls.Add(this.buttonMatchContiguous);
            this.Controls.Add(this.button_FloodZoneOwnership);
            this.Controls.Add(this.buttonPatriotOwnershipCheck);
            this.Controls.Add(this.button_DuplicateParcelIds);
            this.Controls.Add(this.buttonCreateNewParcelIds);
            this.Controls.Add(this.InactiveDifferenceButton);
            this.Controls.Add(this.ProgressBarLabel);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.CombineButton);
            this.Name = "GrandListExport";
            this.Text = "NEMRC Grand List";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button CombineButton;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label ProgressBarLabel;
        private System.Windows.Forms.Button InactiveDifferenceButton;
        private System.Windows.Forms.Button buttonCreateNewParcelIds;
        private System.Windows.Forms.Button button_DuplicateParcelIds;
        private System.Windows.Forms.Button buttonPatriotOwnershipCheck;
        private System.Windows.Forms.Button buttonMatchContiguous;
        private System.Windows.Forms.Button button_FloodZoneOwnership;
        private System.Windows.Forms.Button buttonPatriotOwnershipValueCheck;
        private System.Windows.Forms.Button ActiveDifferenceButton;
        private System.Windows.Forms.RadioButton Working_radioButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton AsBilled_radioButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonPatriotInactiveOwnershipCheck;
    }
}

