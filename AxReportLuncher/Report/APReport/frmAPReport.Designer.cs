namespace NewVersion.Report.APReport
{
    partial class frmAPReport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAPReport));
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.rdoSummay = new System.Windows.Forms.RadioButton();
            this.rdoReconcile = new System.Windows.Forms.RadioButton();
            this.rdoDueDate = new System.Windows.Forms.RadioButton();
            this.btnGenreport = new System.Windows.Forms.Button();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.dtTo = new System.Windows.Forms.DateTimePicker();
            this.lblDateTo = new System.Windows.Forms.Label();
            this.btnRemoveVender = new System.Windows.Forms.Button();
            this.btnAddVender = new System.Windows.Forms.Button();
            this.lstVender2 = new System.Windows.Forms.ListBox();
            this.lstVender1 = new System.Windows.Forms.ListBox();
            this.Label9 = new System.Windows.Forms.Label();
            this.dtFrom = new System.Windows.Forms.DateTimePicker();
            this.lblDateFrom = new System.Windows.Forms.Label();
            this.btnFind = new System.Windows.Forms.Button();
            this.cboFac = new System.Windows.Forms.ComboBox();
            this.Label5 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.txbVendCode = new System.Windows.Forms.TextBox();
            this.GroupBox2.SuspendLayout();
            this.GroupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // GroupBox2
            // 
            this.GroupBox2.Controls.Add(this.rdoSummay);
            this.GroupBox2.Controls.Add(this.rdoReconcile);
            this.GroupBox2.Controls.Add(this.rdoDueDate);
            this.GroupBox2.Location = new System.Drawing.Point(12, 12);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Size = new System.Drawing.Size(388, 70);
            this.GroupBox2.TabIndex = 44;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "Report";
            // 
            // rdoSummay
            // 
            this.rdoSummay.AutoSize = true;
            this.rdoSummay.Location = new System.Drawing.Point(132, 19);
            this.rdoSummay.Name = "rdoSummay";
            this.rdoSummay.Size = new System.Drawing.Size(90, 17);
            this.rdoSummay.TabIndex = 42;
            this.rdoSummay.Text = "A/P Summary";
            this.rdoSummay.UseVisualStyleBackColor = true;
            // 
            // rdoReconcile
            // 
            this.rdoReconcile.AutoSize = true;
            this.rdoReconcile.Location = new System.Drawing.Point(26, 42);
            this.rdoReconcile.Name = "rdoReconcile";
            this.rdoReconcile.Size = new System.Drawing.Size(95, 17);
            this.rdoReconcile.TabIndex = 41;
            this.rdoReconcile.Text = "A/P Reconcile";
            this.rdoReconcile.UseVisualStyleBackColor = true;
            // 
            // rdoDueDate
            // 
            this.rdoDueDate.AutoSize = true;
            this.rdoDueDate.Checked = true;
            this.rdoDueDate.Location = new System.Drawing.Point(26, 19);
            this.rdoDueDate.Name = "rdoDueDate";
            this.rdoDueDate.Size = new System.Drawing.Size(90, 17);
            this.rdoDueDate.TabIndex = 40;
            this.rdoDueDate.TabStop = true;
            this.rdoDueDate.Text = "A/P DueDate";
            this.rdoDueDate.UseVisualStyleBackColor = true;
            // 
            // btnGenreport
            // 
            this.btnGenreport.Location = new System.Drawing.Point(151, 355);
            this.btnGenreport.MaximumSize = new System.Drawing.Size(107, 49);
            this.btnGenreport.MinimumSize = new System.Drawing.Size(107, 49);
            this.btnGenreport.Name = "btnGenreport";
            this.btnGenreport.Size = new System.Drawing.Size(107, 49);
            this.btnGenreport.TabIndex = 43;
            this.btnGenreport.Text = "Get Report";
            this.btnGenreport.UseVisualStyleBackColor = true;
            this.btnGenreport.Click += new System.EventHandler(this.btnGenreport_Click);
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.dtTo);
            this.GroupBox1.Controls.Add(this.lblDateTo);
            this.GroupBox1.Controls.Add(this.btnRemoveVender);
            this.GroupBox1.Controls.Add(this.btnAddVender);
            this.GroupBox1.Controls.Add(this.lstVender2);
            this.GroupBox1.Controls.Add(this.lstVender1);
            this.GroupBox1.Controls.Add(this.Label9);
            this.GroupBox1.Controls.Add(this.dtFrom);
            this.GroupBox1.Controls.Add(this.lblDateFrom);
            this.GroupBox1.Controls.Add(this.btnFind);
            this.GroupBox1.Controls.Add(this.cboFac);
            this.GroupBox1.Controls.Add(this.Label5);
            this.GroupBox1.Controls.Add(this.Label1);
            this.GroupBox1.Controls.Add(this.txbVendCode);
            this.GroupBox1.Location = new System.Drawing.Point(12, 88);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(388, 256);
            this.GroupBox1.TabIndex = 42;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "Criteria";
            // 
            // dtTo
            // 
            this.dtTo.CausesValidation = false;
            this.dtTo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtTo.Location = new System.Drawing.Point(263, 222);
            this.dtTo.Name = "dtTo";
            this.dtTo.Size = new System.Drawing.Size(100, 20);
            this.dtTo.TabIndex = 39;
            this.dtTo.Value = new System.DateTime(2015, 11, 12, 11, 19, 35, 0);
            // 
            // lblDateTo
            // 
            this.lblDateTo.AutoSize = true;
            this.lblDateTo.Location = new System.Drawing.Point(231, 225);
            this.lblDateTo.Name = "lblDateTo";
            this.lblDateTo.Size = new System.Drawing.Size(26, 13);
            this.lblDateTo.TabIndex = 38;
            this.lblDateTo.Text = "To :";
            this.lblDateTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnRemoveVender
            // 
            this.btnRemoveVender.Location = new System.Drawing.Point(228, 125);
            this.btnRemoveVender.Name = "btnRemoveVender";
            this.btnRemoveVender.Size = new System.Drawing.Size(29, 23);
            this.btnRemoveVender.TabIndex = 36;
            this.btnRemoveVender.Text = "<";
            this.btnRemoveVender.UseVisualStyleBackColor = true;
            this.btnRemoveVender.Click += new System.EventHandler(this.btnRemoveVender_Click);
            // 
            // btnAddVender
            // 
            this.btnAddVender.Location = new System.Drawing.Point(228, 96);
            this.btnAddVender.Name = "btnAddVender";
            this.btnAddVender.Size = new System.Drawing.Size(29, 23);
            this.btnAddVender.TabIndex = 35;
            this.btnAddVender.Text = ">";
            this.btnAddVender.UseVisualStyleBackColor = true;
            this.btnAddVender.Click += new System.EventHandler(this.btnAddVender_Click);
            // 
            // lstVender2
            // 
            this.lstVender2.FormattingEnabled = true;
            this.lstVender2.Location = new System.Drawing.Point(263, 82);
            this.lstVender2.Name = "lstVender2";
            this.lstVender2.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstVender2.Size = new System.Drawing.Size(100, 134);
            this.lstVender2.TabIndex = 37;
            this.lstVender2.DoubleClick += new System.EventHandler(this.lstVender2_DoubleClick);
            // 
            // lstVender1
            // 
            this.lstVender1.FormattingEnabled = true;
            this.lstVender1.Location = new System.Drawing.Point(122, 82);
            this.lstVender1.Name = "lstVender1";
            this.lstVender1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstVender1.Size = new System.Drawing.Size(100, 134);
            this.lstVender1.TabIndex = 34;
            this.lstVender1.DoubleClick += new System.EventHandler(this.lstVender1_DoubleClick);
            // 
            // Label9
            // 
            this.Label9.AutoSize = true;
            this.Label9.Location = new System.Drawing.Point(40, 82);
            this.Label9.Name = "Label9";
            this.Label9.Size = new System.Drawing.Size(76, 13);
            this.Label9.TabIndex = 33;
            this.Label9.Text = "&Vender Group:";
            this.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // dtFrom
            // 
            this.dtFrom.CausesValidation = false;
            this.dtFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtFrom.Location = new System.Drawing.Point(122, 222);
            this.dtFrom.Name = "dtFrom";
            this.dtFrom.Size = new System.Drawing.Size(100, 20);
            this.dtFrom.TabIndex = 30;
            this.dtFrom.Value = new System.DateTime(2015, 11, 12, 11, 19, 35, 0);
            // 
            // lblDateFrom
            // 
            this.lblDateFrom.Location = new System.Drawing.Point(35, 225);
            this.lblDateFrom.Name = "lblDateFrom";
            this.lblDateFrom.Size = new System.Drawing.Size(81, 13);
            this.lblDateFrom.TabIndex = 28;
            this.lblDateFrom.Text = "Transac. as of :";
            this.lblDateFrom.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnFind
            // 
            this.btnFind.Image = ((System.Drawing.Image)(resources.GetObject("btnFind.Image")));
            this.btnFind.Location = new System.Drawing.Point(221, 54);
            this.btnFind.Name = "btnFind";
            this.btnFind.Size = new System.Drawing.Size(25, 23);
            this.btnFind.TabIndex = 27;
            this.btnFind.UseVisualStyleBackColor = true;
            this.btnFind.Click += new System.EventHandler(this.btnFind_Click);
            // 
            // cboFac
            // 
            this.cboFac.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFac.FormattingEnabled = true;
            this.cboFac.Location = new System.Drawing.Point(120, 23);
            this.cboFac.Name = "cboFac";
            this.cboFac.Size = new System.Drawing.Size(93, 21);
            this.cboFac.TabIndex = 1;
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(68, 26);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(48, 13);
            this.Label5.TabIndex = 0;
            this.Label5.Text = "Factory :";
            this.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(72, 57);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(44, 13);
            this.Label1.TabIndex = 2;
            this.Label1.Text = "Vender:";
            this.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txbVendCode
            // 
            this.txbVendCode.Location = new System.Drawing.Point(122, 54);
            this.txbVendCode.Name = "txbVendCode";
            this.txbVendCode.Size = new System.Drawing.Size(93, 20);
            this.txbVendCode.TabIndex = 3;
            // 
            // frmAPReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(413, 430);
            this.Controls.Add(this.GroupBox2);
            this.Controls.Add(this.btnGenreport);
            this.Controls.Add(this.GroupBox1);
            this.Name = "frmAPReport";
            this.Text = "APReport";
            this.Load += new System.EventHandler(this.frmAPReport_Load);
            this.GroupBox2.ResumeLayout(false);
            this.GroupBox2.PerformLayout();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.GroupBox GroupBox2;
        internal System.Windows.Forms.RadioButton rdoSummay;
        internal System.Windows.Forms.RadioButton rdoReconcile;
        internal System.Windows.Forms.RadioButton rdoDueDate;
        internal System.Windows.Forms.Button btnGenreport;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.DateTimePicker dtTo;
        internal System.Windows.Forms.Label lblDateTo;
        internal System.Windows.Forms.Button btnRemoveVender;
        internal System.Windows.Forms.Button btnAddVender;
        internal System.Windows.Forms.ListBox lstVender2;
        internal System.Windows.Forms.ListBox lstVender1;
        internal System.Windows.Forms.Label Label9;
        internal System.Windows.Forms.DateTimePicker dtFrom;
        internal System.Windows.Forms.Label lblDateFrom;
        internal System.Windows.Forms.Button btnFind;
        internal System.Windows.Forms.ComboBox cboFac;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.TextBox txbVendCode;
    }
}