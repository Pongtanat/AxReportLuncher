namespace NewVersion.Report.POReamain
{
    partial class PORemain
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
            this.btnGenreport = new System.Windows.Forms.Button();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.cboShpLoc = new System.Windows.Forms.ComboBox();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.dtDate2 = new System.Windows.Forms.DateTimePicker();
            this.Label5 = new System.Windows.Forms.Label();
            this.dtDate1 = new System.Windows.Forms.DateTimePicker();
            this.cboFac = new System.Windows.Forms.ComboBox();
            this.cboReport = new System.Windows.Forms.ComboBox();
            this.GroupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGenreport
            // 
            this.btnGenreport.Location = new System.Drawing.Point(118, 201);
            this.btnGenreport.Name = "btnGenreport";
            this.btnGenreport.Size = new System.Drawing.Size(107, 49);
            this.btnGenreport.TabIndex = 26;
            this.btnGenreport.Text = "Get Report";
            this.btnGenreport.UseVisualStyleBackColor = true;
            this.btnGenreport.Click += new System.EventHandler(this.btnGenreport_Click);
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
            this.GroupBox1.Location = new System.Drawing.Point(12, 12);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(347, 162);
            this.GroupBox1.TabIndex = 25;
            this.GroupBox1.TabStop = false;
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(5, 129);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(102, 13);
            this.Label3.TabIndex = 34;
            this.Label3.Text = "Number Sequence :";
            // 
            // cboShpLoc
            // 
            this.cboShpLoc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboShpLoc.FormattingEnabled = true;
            this.cboShpLoc.Location = new System.Drawing.Point(114, 126);
            this.cboShpLoc.Name = "cboShpLoc";
            this.cboShpLoc.Size = new System.Drawing.Size(100, 21);
            this.cboShpLoc.TabIndex = 33;
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(36, 106);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(66, 13);
            this.Label4.TabIndex = 11;
            this.Label4.Text = "PO Date to :";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(25, 80);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(77, 13);
            this.Label2.TabIndex = 9;
            this.Label2.Text = "PO Date from :";
            this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dtDate2
            // 
            this.dtDate2.CustomFormat = "dd/MM/yyyy";
            this.dtDate2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDate2.Location = new System.Drawing.Point(114, 100);
            this.dtDate2.Name = "dtDate2";
            this.dtDate2.Size = new System.Drawing.Size(100, 20);
            this.dtDate2.TabIndex = 10;
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(59, 32);
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
            this.dtDate1.Location = new System.Drawing.Point(114, 74);
            this.dtDate1.Name = "dtDate1";
            this.dtDate1.Size = new System.Drawing.Size(100, 20);
            this.dtDate1.TabIndex = 10;
            // 
            // cboFac
            // 
            this.cboFac.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFac.FormattingEnabled = true;
            this.cboFac.Location = new System.Drawing.Point(113, 29);
            this.cboFac.Name = "cboFac";
            this.cboFac.Size = new System.Drawing.Size(100, 21);
            this.cboFac.TabIndex = 3;
            // 
            // cboReport
            // 
            this.cboReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboReport.FormattingEnabled = true;
            this.cboReport.ItemHeight = 13;
            this.cboReport.Location = new System.Drawing.Point(83, -60);
            this.cboReport.Name = "cboReport";
            this.cboReport.Size = new System.Drawing.Size(327, 21);
            this.cboReport.TabIndex = 24;
            // 
            // PORemain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(374, 269);
            this.Controls.Add(this.btnGenreport);
            this.Controls.Add(this.GroupBox1);
            this.Controls.Add(this.cboReport);
            this.Name = "PORemain";
            this.Text = "PORemain";
            this.Load += new System.EventHandler(this.PORemain_Load);
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Button btnGenreport;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.ComboBox cboShpLoc;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.DateTimePicker dtDate2;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.DateTimePicker dtDate1;
        internal System.Windows.Forms.ComboBox cboFac;
        internal System.Windows.Forms.ComboBox cboReport;
    }
}