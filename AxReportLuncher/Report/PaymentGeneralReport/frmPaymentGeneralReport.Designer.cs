namespace NewVersion.Report.PaymentGeneralReport
{
    partial class frmPaymentGeneralReport
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
            this.Label4 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.dtDate2 = new System.Windows.Forms.DateTimePicker();
            this.Label5 = new System.Windows.Forms.Label();
            this.dtDate1 = new System.Windows.Forms.DateTimePicker();
            this.cboFac = new System.Windows.Forms.ComboBox();
            this.cboReport = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtStartVoucher = new System.Windows.Forms.TextBox();
            this.txtEndVoucher = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cboGroup = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.GroupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGenreport
            // 
            this.btnGenreport.Location = new System.Drawing.Point(73, 254);
            this.btnGenreport.Name = "btnGenreport";
            this.btnGenreport.Size = new System.Drawing.Size(198, 49);
            this.btnGenreport.TabIndex = 26;
            this.btnGenreport.Text = "Get Report";
            this.btnGenreport.UseVisualStyleBackColor = true;
            this.btnGenreport.Click += new System.EventHandler(this.btnGenreport_Click);
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.Label4);
            this.GroupBox1.Controls.Add(this.Label2);
            this.GroupBox1.Controls.Add(this.dtDate2);
            this.GroupBox1.Controls.Add(this.Label5);
            this.GroupBox1.Controls.Add(this.dtDate1);
            this.GroupBox1.Controls.Add(this.cboFac);
            this.GroupBox1.Location = new System.Drawing.Point(12, 12);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(301, 116);
            this.GroupBox1.TabIndex = 25;
            this.GroupBox1.TabStop = false;
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(44, 77);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(86, 13);
            this.Label4.TabIndex = 11;
            this.Label4.Text = "Invoice Date to :";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(33, 51);
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
            this.dtDate2.Location = new System.Drawing.Point(136, 73);
            this.dtDate2.Name = "dtDate2";
            this.dtDate2.Size = new System.Drawing.Size(100, 20);
            this.dtDate2.TabIndex = 10;
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(82, 22);
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
            this.dtDate1.Location = new System.Drawing.Point(136, 47);
            this.dtDate1.Name = "dtDate1";
            this.dtDate1.Size = new System.Drawing.Size(100, 20);
            this.dtDate1.TabIndex = 10;
            // 
            // cboFac
            // 
            this.cboFac.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFac.FormattingEnabled = true;
            this.cboFac.Location = new System.Drawing.Point(136, 19);
            this.cboFac.Name = "cboFac";
            this.cboFac.Size = new System.Drawing.Size(100, 21);
            this.cboFac.TabIndex = 3;
            // 
            // cboReport
            // 
            this.cboReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboReport.FormattingEnabled = true;
            this.cboReport.ItemHeight = 13;
            this.cboReport.Location = new System.Drawing.Point(17, -70);
            this.cboReport.Name = "cboReport";
            this.cboReport.Size = new System.Drawing.Size(327, 21);
            this.cboReport.TabIndex = 24;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 27;
            this.label1.Text = "Start Voucher :";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 42);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 13);
            this.label3.TabIndex = 28;
            this.label3.Text = "End Voucher :";
            // 
            // txtStartVoucher
            // 
            this.txtStartVoucher.Location = new System.Drawing.Point(109, 13);
            this.txtStartVoucher.Name = "txtStartVoucher";
            this.txtStartVoucher.Size = new System.Drawing.Size(149, 20);
            this.txtStartVoucher.TabIndex = 29;
            // 
            // txtEndVoucher
            // 
            this.txtEndVoucher.Location = new System.Drawing.Point(109, 39);
            this.txtEndVoucher.Name = "txtEndVoucher";
            this.txtEndVoucher.Size = new System.Drawing.Size(149, 20);
            this.txtEndVoucher.TabIndex = 30;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(60, 68);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 13);
            this.label6.TabIndex = 31;
            this.label6.Text = "Group :";
            // 
            // cboGroup
            // 
            this.cboGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboGroup.FormattingEnabled = true;
            this.cboGroup.Location = new System.Drawing.Point(109, 65);
            this.cboGroup.Name = "cboGroup";
            this.cboGroup.Size = new System.Drawing.Size(149, 21);
            this.cboGroup.TabIndex = 32;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.cboGroup);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.txtStartVoucher);
            this.groupBox2.Controls.Add(this.txtEndVoucher);
            this.groupBox2.Location = new System.Drawing.Point(13, 135);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(300, 100);
            this.groupBox2.TabIndex = 33;
            this.groupBox2.TabStop = false;
            // 
            // frmPaymentGeneralReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(324, 322);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnGenreport);
            this.Controls.Add(this.GroupBox1);
            this.Controls.Add(this.cboReport);
            this.MaximumSize = new System.Drawing.Size(340, 360);
            this.MinimumSize = new System.Drawing.Size(340, 360);
            this.Name = "frmPaymentGeneralReport";
            this.Text = "PaymentGeneral Report";
            this.Load += new System.EventHandler(this.frmPaymentGeneralReport_Load);
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Button btnGenreport;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.DateTimePicker dtDate2;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.DateTimePicker dtDate1;
        internal System.Windows.Forms.ComboBox cboFac;
        internal System.Windows.Forms.ComboBox cboReport;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtStartVoucher;
        private System.Windows.Forms.TextBox txtEndVoucher;
        private System.Windows.Forms.Label label6;
        internal System.Windows.Forms.ComboBox cboGroup;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}