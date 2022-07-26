namespace NewVersion.Report.QuickSales_Report
{
    partial class frmQuickSales
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmQuickSales));
            this.cboReport = new System.Windows.Forms.ComboBox();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.cboShpLoc = new System.Windows.Forms.ComboBox();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.dtDate2 = new System.Windows.Forms.DateTimePicker();
            this.Label5 = new System.Windows.Forms.Label();
            this.dtDate1 = new System.Windows.Forms.DateTimePicker();
            this.cboFac = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkRL_OTH = new System.Windows.Forms.CheckBox();
            this.chkRL_OTHER = new System.Windows.Forms.CheckBox();
            this.chkRL_FA = new System.Windows.Forms.CheckBox();
            this.chkRL_EXP = new System.Windows.Forms.CheckBox();
            this.chkNR_OTH = new System.Windows.Forms.CheckBox();
            this.chkNR_EXP = new System.Windows.Forms.CheckBox();
            this.chkNR_DOM = new System.Windows.Forms.CheckBox();
            this.chkIN_DOM = new System.Windows.Forms.CheckBox();
            this.chkIN_ADP = new System.Windows.Forms.CheckBox();
            this.chkIN_DES = new System.Windows.Forms.CheckBox();
            this.btnGenreport = new System.Windows.Forms.Button();
            this.GroupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // cboReport
            // 
            this.cboReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboReport.FormattingEnabled = true;
            this.cboReport.ItemHeight = 13;
            this.cboReport.Location = new System.Drawing.Point(19, 12);
            this.cboReport.Name = "cboReport";
            this.cboReport.Size = new System.Drawing.Size(327, 21);
            this.cboReport.TabIndex = 18;
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.Label3);
            this.GroupBox1.Controls.Add(this.cboShpLoc);
            this.GroupBox1.Controls.Add(this.Label4);
            this.GroupBox1.Controls.Add(this.Label2);
            this.GroupBox1.Controls.Add(this.dtDate2);
            this.GroupBox1.Controls.Add(this.Label5);
            this.GroupBox1.Controls.Add(this.dtDate1);
            this.GroupBox1.Controls.Add(this.cboFac);
            this.GroupBox1.Location = new System.Drawing.Point(9, 48);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(347, 136);
            this.GroupBox1.TabIndex = 19;
            this.GroupBox1.TabStop = false;
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(29, 102);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(78, 13);
            this.Label3.TabIndex = 34;
            this.Label3.Text = "Shipment Loc :";
            // 
            // cboShpLoc
            // 
            this.cboShpLoc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboShpLoc.FormattingEnabled = true;
            this.cboShpLoc.Location = new System.Drawing.Point(113, 99);
            this.cboShpLoc.Name = "cboShpLoc";
            this.cboShpLoc.Size = new System.Drawing.Size(100, 21);
            this.cboShpLoc.TabIndex = 33;
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(21, 77);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(86, 13);
            this.Label4.TabIndex = 11;
            this.Label4.Text = "Invoice Date to :";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(10, 51);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(97, 13);
            this.Label2.TabIndex = 9;
            this.Label2.Text = "Invoice Date from :";
            this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dtDate2
            // 
            this.dtDate2.CustomFormat = "dd/MM/yyyy";
            this.dtDate2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDate2.Location = new System.Drawing.Point(113, 73);
            this.dtDate2.Name = "dtDate2";
            this.dtDate2.Size = new System.Drawing.Size(100, 20);
            this.dtDate2.TabIndex = 12;
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(59, 22);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(48, 13);
            this.Label5.TabIndex = 2;
            this.Label5.Text = "Factory :";
            this.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // dtDate1
            // 
            this.dtDate1.CustomFormat = "dd/MM/yyyy";
            this.dtDate1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDate1.Location = new System.Drawing.Point(113, 47);
            this.dtDate1.Name = "dtDate1";
            this.dtDate1.Size = new System.Drawing.Size(100, 20);
            this.dtDate1.TabIndex = 10;
            // 
            // cboFac
            // 
            this.cboFac.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFac.FormattingEnabled = true;
            this.cboFac.Location = new System.Drawing.Point(113, 19);
            this.cboFac.Name = "cboFac";
            this.cboFac.Size = new System.Drawing.Size(100, 21);
            this.cboFac.TabIndex = 3;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkRL_OTH);
            this.groupBox2.Controls.Add(this.chkRL_OTHER);
            this.groupBox2.Controls.Add(this.chkRL_FA);
            this.groupBox2.Controls.Add(this.chkRL_EXP);
            this.groupBox2.Controls.Add(this.chkNR_OTH);
            this.groupBox2.Controls.Add(this.chkNR_EXP);
            this.groupBox2.Controls.Add(this.chkNR_DOM);
            this.groupBox2.Controls.Add(this.chkIN_DOM);
            this.groupBox2.Controls.Add(this.chkIN_ADP);
            this.groupBox2.Controls.Add(this.chkIN_DES);
            this.groupBox2.Location = new System.Drawing.Point(10, 190);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(346, 110);
            this.groupBox2.TabIndex = 20;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Cust Group:";
            // 
            // chkRL_OTH
            // 
            this.chkRL_OTH.AutoSize = true;
            this.chkRL_OTH.Checked = true;
            this.chkRL_OTH.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRL_OTH.Location = new System.Drawing.Point(98, 87);
            this.chkRL_OTH.Name = "chkRL_OTH";
            this.chkRL_OTH.Size = new System.Drawing.Size(66, 17);
            this.chkRL_OTH.TabIndex = 43;
            this.chkRL_OTH.Text = "RL-OTH";
            this.chkRL_OTH.UseVisualStyleBackColor = true;
            // 
            // chkRL_OTHER
            // 
            this.chkRL_OTHER.AutoSize = true;
            this.chkRL_OTHER.Checked = true;
            this.chkRL_OTHER.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRL_OTHER.Location = new System.Drawing.Point(98, 64);
            this.chkRL_OTHER.Name = "chkRL_OTHER";
            this.chkRL_OTHER.Size = new System.Drawing.Size(81, 17);
            this.chkRL_OTHER.TabIndex = 42;
            this.chkRL_OTHER.Text = "RL-OTHER";
            this.chkRL_OTHER.UseVisualStyleBackColor = true;
            // 
            // chkRL_FA
            // 
            this.chkRL_FA.AutoSize = true;
            this.chkRL_FA.Checked = true;
            this.chkRL_FA.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRL_FA.Location = new System.Drawing.Point(98, 41);
            this.chkRL_FA.Name = "chkRL_FA";
            this.chkRL_FA.Size = new System.Drawing.Size(61, 17);
            this.chkRL_FA.TabIndex = 41;
            this.chkRL_FA.Text = "RL-F/A";
            this.chkRL_FA.UseVisualStyleBackColor = true;
            // 
            // chkRL_EXP
            // 
            this.chkRL_EXP.AutoSize = true;
            this.chkRL_EXP.Checked = true;
            this.chkRL_EXP.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRL_EXP.Location = new System.Drawing.Point(98, 18);
            this.chkRL_EXP.Name = "chkRL_EXP";
            this.chkRL_EXP.Size = new System.Drawing.Size(64, 17);
            this.chkRL_EXP.TabIndex = 40;
            this.chkRL_EXP.Text = "RL-EXP";
            this.chkRL_EXP.UseVisualStyleBackColor = true;
            // 
            // chkNR_OTH
            // 
            this.chkNR_OTH.AutoSize = true;
            this.chkNR_OTH.Checked = true;
            this.chkNR_OTH.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkNR_OTH.Location = new System.Drawing.Point(251, 64);
            this.chkNR_OTH.Name = "chkNR_OTH";
            this.chkNR_OTH.Size = new System.Drawing.Size(68, 17);
            this.chkNR_OTH.TabIndex = 39;
            this.chkNR_OTH.Text = "NR-OTH";
            this.chkNR_OTH.UseVisualStyleBackColor = true;
            // 
            // chkNR_EXP
            // 
            this.chkNR_EXP.AutoSize = true;
            this.chkNR_EXP.Checked = true;
            this.chkNR_EXP.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkNR_EXP.Location = new System.Drawing.Point(251, 41);
            this.chkNR_EXP.Name = "chkNR_EXP";
            this.chkNR_EXP.Size = new System.Drawing.Size(66, 17);
            this.chkNR_EXP.TabIndex = 38;
            this.chkNR_EXP.Text = "NR-EXP";
            this.chkNR_EXP.UseVisualStyleBackColor = true;
            // 
            // chkNR_DOM
            // 
            this.chkNR_DOM.AutoSize = true;
            this.chkNR_DOM.Checked = true;
            this.chkNR_DOM.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkNR_DOM.Location = new System.Drawing.Point(251, 18);
            this.chkNR_DOM.Name = "chkNR_DOM";
            this.chkNR_DOM.Size = new System.Drawing.Size(70, 17);
            this.chkNR_DOM.TabIndex = 37;
            this.chkNR_DOM.Text = "NR-DOM";
            this.chkNR_DOM.UseVisualStyleBackColor = true;
            // 
            // chkIN_DOM
            // 
            this.chkIN_DOM.AutoSize = true;
            this.chkIN_DOM.Checked = true;
            this.chkIN_DOM.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIN_DOM.Location = new System.Drawing.Point(183, 64);
            this.chkIN_DOM.Name = "chkIN_DOM";
            this.chkIN_DOM.Size = new System.Drawing.Size(65, 17);
            this.chkIN_DOM.TabIndex = 36;
            this.chkIN_DOM.Text = "IN-DOM";
            this.chkIN_DOM.UseVisualStyleBackColor = true;
            // 
            // chkIN_ADP
            // 
            this.chkIN_ADP.AutoSize = true;
            this.chkIN_ADP.Checked = true;
            this.chkIN_ADP.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIN_ADP.Location = new System.Drawing.Point(183, 18);
            this.chkIN_ADP.Name = "chkIN_ADP";
            this.chkIN_ADP.Size = new System.Drawing.Size(62, 17);
            this.chkIN_ADP.TabIndex = 34;
            this.chkIN_ADP.Text = "IN-ADP";
            this.chkIN_ADP.UseVisualStyleBackColor = true;
            // 
            // chkIN_DES
            // 
            this.chkIN_DES.AutoSize = true;
            this.chkIN_DES.Checked = true;
            this.chkIN_DES.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIN_DES.Location = new System.Drawing.Point(183, 41);
            this.chkIN_DES.Name = "chkIN_DES";
            this.chkIN_DES.Size = new System.Drawing.Size(62, 17);
            this.chkIN_DES.TabIndex = 35;
            this.chkIN_DES.Text = "IN-DES";
            this.chkIN_DES.UseVisualStyleBackColor = true;
            // 
            // btnGenreport
            // 
            this.btnGenreport.Location = new System.Drawing.Point(115, 312);
            this.btnGenreport.Name = "btnGenreport";
            this.btnGenreport.Size = new System.Drawing.Size(107, 49);
            this.btnGenreport.TabIndex = 21;
            this.btnGenreport.Text = "Get Report";
            this.btnGenreport.UseVisualStyleBackColor = true;
            this.btnGenreport.Click += new System.EventHandler(this.btnGenreport_Click);
            // 
            // frmQuickSales
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 382);
            this.Controls.Add(this.btnGenreport);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.GroupBox1);
            this.Controls.Add(this.cboReport);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(384, 420);
            this.Name = "frmQuickSales";
            this.Text = "frmQuickSales";
            this.Load += new System.EventHandler(this.frmQuickSales_Load);
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.ComboBox cboReport;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.ComboBox cboShpLoc;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.DateTimePicker dtDate2;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.DateTimePicker dtDate1;
        internal System.Windows.Forms.ComboBox cboFac;
        private System.Windows.Forms.GroupBox groupBox2;
        internal System.Windows.Forms.CheckBox chkRL_OTH;
        internal System.Windows.Forms.CheckBox chkRL_OTHER;
        internal System.Windows.Forms.CheckBox chkRL_FA;
        internal System.Windows.Forms.CheckBox chkRL_EXP;
        internal System.Windows.Forms.CheckBox chkNR_OTH;
        internal System.Windows.Forms.CheckBox chkNR_EXP;
        internal System.Windows.Forms.CheckBox chkNR_DOM;
        internal System.Windows.Forms.CheckBox chkIN_DOM;
        internal System.Windows.Forms.CheckBox chkIN_ADP;
        internal System.Windows.Forms.CheckBox chkIN_DES;
        internal System.Windows.Forms.Button btnGenreport;
    }
}