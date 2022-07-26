namespace NewVersion.Report.SalesReturn.SummaryTransaction
{
    partial class frmReturnTransaction
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
            this.StatusStrip1 = new System.Windows.Forms.StatusStrip();
            this.ToolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.rdoByVoucher = new System.Windows.Forms.RadioButton();
            this.rdoBySection = new System.Windows.Forms.RadioButton();
            this.btnGenreport = new System.Windows.Forms.Button();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.Label6 = new System.Windows.Forms.Label();
            this.Label7 = new System.Windows.Forms.Label();
            this.txbVoucher1 = new System.Windows.Forms.TextBox();
            this.txbVoucher2 = new System.Windows.Forms.TextBox();
            this.rdoShipment = new System.Windows.Forms.RadioButton();
            this.rdoReceive = new System.Windows.Forms.RadioButton();
            this.btnRemoveAllSection = new System.Windows.Forms.Button();
            this.btnAddAllSection = new System.Windows.Forms.Button();
            this.Label10 = new System.Windows.Forms.Label();
            this.btnRemoveSection = new System.Windows.Forms.Button();
            this.btnAddSection = new System.Windows.Forms.Button();
            this.lstSection2 = new System.Windows.Forms.ListBox();
            this.lstSection1 = new System.Windows.Forms.ListBox();
            this.btnRemoveCategory = new System.Windows.Forms.Button();
            this.btnAddCategory = new System.Windows.Forms.Button();
            this.lstCategory2 = new System.Windows.Forms.ListBox();
            this.lstCategory1 = new System.Windows.Forms.ListBox();
            this.Label9 = new System.Windows.Forms.Label();
            this.Label8 = new System.Windows.Forms.Label();
            this.cboFac = new System.Windows.Forms.ComboBox();
            this.Label5 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.dtDate2 = new System.Windows.Forms.DateTimePicker();
            this.txbItem1 = new System.Windows.Forms.TextBox();
            this.dtDate1 = new System.Windows.Forms.DateTimePicker();
            this.txbItem2 = new System.Windows.Forms.TextBox();
            this.rdoAll = new System.Windows.Forms.RadioButton();
            this.rdoW1 = new System.Windows.Forms.RadioButton();
            this.rdoW2 = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.StatusStrip1.SuspendLayout();
            this.GroupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // StatusStrip1
            // 
            this.StatusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripStatusLabel1});
            this.StatusStrip1.Location = new System.Drawing.Point(0, 566);
            this.StatusStrip1.Name = "StatusStrip1";
            this.StatusStrip1.Size = new System.Drawing.Size(437, 22);
            this.StatusStrip1.SizingGrip = false;
            this.StatusStrip1.TabIndex = 16;
            this.StatusStrip1.Text = "StatusStrip1";
            // 
            // ToolStripStatusLabel1
            // 
            this.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1";
            this.ToolStripStatusLabel1.Size = new System.Drawing.Size(121, 17);
            this.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1";
            // 
            // rdoByVoucher
            // 
            this.rdoByVoucher.AutoSize = true;
            this.rdoByVoucher.Location = new System.Drawing.Point(12, 510);
            this.rdoByVoucher.Name = "rdoByVoucher";
            this.rdoByVoucher.Size = new System.Drawing.Size(101, 17);
            this.rdoByVoucher.TabIndex = 15;
            this.rdoByVoucher.Text = "Sort by Voucher";
            this.rdoByVoucher.UseVisualStyleBackColor = true;
            // 
            // rdoBySection
            // 
            this.rdoBySection.AutoSize = true;
            this.rdoBySection.Checked = true;
            this.rdoBySection.Location = new System.Drawing.Point(12, 487);
            this.rdoBySection.Name = "rdoBySection";
            this.rdoBySection.Size = new System.Drawing.Size(97, 17);
            this.rdoBySection.TabIndex = 14;
            this.rdoBySection.TabStop = true;
            this.rdoBySection.Text = "Sort by Section";
            this.rdoBySection.UseVisualStyleBackColor = true;
            // 
            // btnGenreport
            // 
            this.btnGenreport.Location = new System.Drawing.Point(142, 478);
            this.btnGenreport.Name = "btnGenreport";
            this.btnGenreport.Size = new System.Drawing.Size(107, 49);
            this.btnGenreport.TabIndex = 13;
            this.btnGenreport.Text = "&Get Report";
            this.btnGenreport.UseVisualStyleBackColor = true;
            this.btnGenreport.Click += new System.EventHandler(this.btnGenreport_Click);
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.groupBox2);
            this.GroupBox1.Controls.Add(this.Label6);
            this.GroupBox1.Controls.Add(this.Label7);
            this.GroupBox1.Controls.Add(this.txbVoucher1);
            this.GroupBox1.Controls.Add(this.txbVoucher2);
            this.GroupBox1.Controls.Add(this.rdoShipment);
            this.GroupBox1.Controls.Add(this.rdoReceive);
            this.GroupBox1.Controls.Add(this.btnRemoveAllSection);
            this.GroupBox1.Controls.Add(this.btnAddAllSection);
            this.GroupBox1.Controls.Add(this.Label10);
            this.GroupBox1.Controls.Add(this.btnRemoveSection);
            this.GroupBox1.Controls.Add(this.btnAddSection);
            this.GroupBox1.Controls.Add(this.lstSection2);
            this.GroupBox1.Controls.Add(this.lstSection1);
            this.GroupBox1.Controls.Add(this.btnRemoveCategory);
            this.GroupBox1.Controls.Add(this.btnAddCategory);
            this.GroupBox1.Controls.Add(this.lstCategory2);
            this.GroupBox1.Controls.Add(this.lstCategory1);
            this.GroupBox1.Controls.Add(this.Label9);
            this.GroupBox1.Controls.Add(this.Label8);
            this.GroupBox1.Controls.Add(this.cboFac);
            this.GroupBox1.Controls.Add(this.Label5);
            this.GroupBox1.Controls.Add(this.Label4);
            this.GroupBox1.Controls.Add(this.Label1);
            this.GroupBox1.Controls.Add(this.Label3);
            this.GroupBox1.Controls.Add(this.Label2);
            this.GroupBox1.Controls.Add(this.dtDate2);
            this.GroupBox1.Controls.Add(this.txbItem1);
            this.GroupBox1.Controls.Add(this.dtDate1);
            this.GroupBox1.Controls.Add(this.txbItem2);
            this.GroupBox1.Location = new System.Drawing.Point(12, 6);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(413, 450);
            this.GroupBox1.TabIndex = 12;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "Criteria";
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(23, 371);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(53, 13);
            this.Label6.TabIndex = 27;
            this.Label6.Text = "&Voucher :";
            this.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label7
            // 
            this.Label7.AutoSize = true;
            this.Label7.Location = new System.Drawing.Point(205, 371);
            this.Label7.Name = "Label7";
            this.Label7.Size = new System.Drawing.Size(23, 13);
            this.Label7.TabIndex = 29;
            this.Label7.Text = "To:";
            // 
            // txbVoucher1
            // 
            this.txbVoucher1.Location = new System.Drawing.Point(82, 368);
            this.txbVoucher1.Name = "txbVoucher1";
            this.txbVoucher1.Size = new System.Drawing.Size(100, 20);
            this.txbVoucher1.TabIndex = 28;
            // 
            // txbVoucher2
            // 
            this.txbVoucher2.Location = new System.Drawing.Point(255, 368);
            this.txbVoucher2.Name = "txbVoucher2";
            this.txbVoucher2.Size = new System.Drawing.Size(100, 20);
            this.txbVoucher2.TabIndex = 30;
            // 
            // rdoShipment
            // 
            this.rdoShipment.AutoSize = true;
            this.rdoShipment.Location = new System.Drawing.Point(153, 48);
            this.rdoShipment.Name = "rdoShipment";
            this.rdoShipment.Size = new System.Drawing.Size(69, 17);
            this.rdoShipment.TabIndex = 26;
            this.rdoShipment.Text = "&Shipment";
            this.rdoShipment.UseVisualStyleBackColor = true;
            // 
            // rdoReceive
            // 
            this.rdoReceive.AutoSize = true;
            this.rdoReceive.Checked = true;
            this.rdoReceive.Location = new System.Drawing.Point(82, 48);
            this.rdoReceive.Name = "rdoReceive";
            this.rdoReceive.Size = new System.Drawing.Size(65, 17);
            this.rdoReceive.TabIndex = 6;
            this.rdoReceive.TabStop = true;
            this.rdoReceive.Text = "&Receive";
            this.rdoReceive.UseVisualStyleBackColor = true;
            // 
            // btnRemoveAllSection
            // 
            this.btnRemoveAllSection.Location = new System.Drawing.Point(203, 323);
            this.btnRemoveAllSection.Name = "btnRemoveAllSection";
            this.btnRemoveAllSection.Size = new System.Drawing.Size(29, 23);
            this.btnRemoveAllSection.TabIndex = 16;
            this.btnRemoveAllSection.Text = "|<";
            this.btnRemoveAllSection.UseVisualStyleBackColor = true;
            this.btnRemoveAllSection.Click += new System.EventHandler(this.btnRemoveAllSection_Click);
            // 
            // btnAddAllSection
            // 
            this.btnAddAllSection.Location = new System.Drawing.Point(203, 207);
            this.btnAddAllSection.Name = "btnAddAllSection";
            this.btnAddAllSection.Size = new System.Drawing.Size(29, 23);
            this.btnAddAllSection.TabIndex = 12;
            this.btnAddAllSection.Text = ">|";
            this.btnAddAllSection.UseVisualStyleBackColor = true;
            this.btnAddAllSection.Click += new System.EventHandler(this.btnAddAllSection_Click);
            // 
            // Label10
            // 
            this.Label10.AutoSize = true;
            this.Label10.Location = new System.Drawing.Point(27, 176);
            this.Label10.Name = "Label10";
            this.Label10.Size = new System.Drawing.Size(49, 13);
            this.Label10.TabIndex = 10;
            this.Label10.Text = "Se&ction :";
            this.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // btnRemoveSection
            // 
            this.btnRemoveSection.Location = new System.Drawing.Point(203, 294);
            this.btnRemoveSection.Name = "btnRemoveSection";
            this.btnRemoveSection.Size = new System.Drawing.Size(29, 23);
            this.btnRemoveSection.TabIndex = 15;
            this.btnRemoveSection.Text = "<";
            this.btnRemoveSection.UseVisualStyleBackColor = true;
            this.btnRemoveSection.Click += new System.EventHandler(this.btnRemoveSection_Click);
            // 
            // btnAddSection
            // 
            this.btnAddSection.Location = new System.Drawing.Point(203, 236);
            this.btnAddSection.Name = "btnAddSection";
            this.btnAddSection.Size = new System.Drawing.Size(29, 23);
            this.btnAddSection.TabIndex = 13;
            this.btnAddSection.Text = ">";
            this.btnAddSection.UseVisualStyleBackColor = true;
            this.btnAddSection.Click += new System.EventHandler(this.btnAddSection_Click);
            // 
            // lstSection2
            // 
            this.lstSection2.FormattingEnabled = true;
            this.lstSection2.Location = new System.Drawing.Point(255, 176);
            this.lstSection2.Name = "lstSection2";
            this.lstSection2.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstSection2.Size = new System.Drawing.Size(100, 186);
            this.lstSection2.TabIndex = 17;
            // 
            // lstSection1
            // 
            this.lstSection1.FormattingEnabled = true;
            this.lstSection1.Location = new System.Drawing.Point(82, 176);
            this.lstSection1.Name = "lstSection1";
            this.lstSection1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstSection1.Size = new System.Drawing.Size(100, 186);
            this.lstSection1.TabIndex = 11;
            // 
            // btnRemoveCategory
            // 
            this.btnRemoveCategory.Location = new System.Drawing.Point(204, 118);
            this.btnRemoveCategory.Name = "btnRemoveCategory";
            this.btnRemoveCategory.Size = new System.Drawing.Size(29, 23);
            this.btnRemoveCategory.TabIndex = 8;
            this.btnRemoveCategory.Text = "<";
            this.btnRemoveCategory.UseVisualStyleBackColor = true;
            this.btnRemoveCategory.Click += new System.EventHandler(this.btnRemoveCategory_Click);
            // 
            // btnAddCategory
            // 
            this.btnAddCategory.Location = new System.Drawing.Point(204, 89);
            this.btnAddCategory.Name = "btnAddCategory";
            this.btnAddCategory.Size = new System.Drawing.Size(29, 23);
            this.btnAddCategory.TabIndex = 7;
            this.btnAddCategory.Text = ">";
            this.btnAddCategory.UseVisualStyleBackColor = true;
            this.btnAddCategory.Click += new System.EventHandler(this.btnAddCategory_Click);
            // 
            // lstCategory2
            // 
            this.lstCategory2.FormattingEnabled = true;
            this.lstCategory2.Location = new System.Drawing.Point(255, 75);
            this.lstCategory2.Name = "lstCategory2";
            this.lstCategory2.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstCategory2.Size = new System.Drawing.Size(100, 95);
            this.lstCategory2.TabIndex = 9;
            // 
            // lstCategory1
            // 
            this.lstCategory1.FormattingEnabled = true;
            this.lstCategory1.Location = new System.Drawing.Point(82, 75);
            this.lstCategory1.Name = "lstCategory1";
            this.lstCategory1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstCategory1.Size = new System.Drawing.Size(100, 95);
            this.lstCategory1.TabIndex = 6;
            // 
            // Label9
            // 
            this.Label9.AutoSize = true;
            this.Label9.Location = new System.Drawing.Point(21, 75);
            this.Label9.Name = "Label9";
            this.Label9.Size = new System.Drawing.Size(55, 13);
            this.Label9.TabIndex = 5;
            this.Label9.Text = "&Category :";
            this.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // Label8
            // 
            this.Label8.AutoSize = true;
            this.Label8.Location = new System.Drawing.Point(9, 50);
            this.Label8.Name = "Label8";
            this.Label8.Size = new System.Drawing.Size(67, 13);
            this.Label8.TabIndex = 2;
            this.Label8.Text = "&Trans Type :";
            this.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // cboFac
            // 
            this.cboFac.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFac.FormattingEnabled = true;
            this.cboFac.Location = new System.Drawing.Point(83, 21);
            this.cboFac.Name = "cboFac";
            this.cboFac.Size = new System.Drawing.Size(100, 21);
            this.cboFac.TabIndex = 1;
            this.cboFac.SelectedIndexChanged += new System.EventHandler(this.cboFac_SelectedIndexChanged);
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(45, 24);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(31, 13);
            this.Label5.TabIndex = 0;
            this.Label5.Text = "Site :";
            this.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(203, 425);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(23, 13);
            this.Label4.TabIndex = 24;
            this.Label4.Text = "To:";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(43, 397);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(33, 13);
            this.Label1.TabIndex = 18;
            this.Label1.Text = "&Item :";
            this.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(205, 397);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(23, 13);
            this.Label3.TabIndex = 20;
            this.Label3.Text = "To:";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(43, 425);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(33, 13);
            this.Label2.TabIndex = 22;
            this.Label2.Text = "&Date:";
            this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dtDate2
            // 
            this.dtDate2.CustomFormat = "dd/MM/yyyy";
            this.dtDate2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDate2.Location = new System.Drawing.Point(255, 419);
            this.dtDate2.Name = "dtDate2";
            this.dtDate2.Size = new System.Drawing.Size(100, 20);
            this.dtDate2.TabIndex = 25;
            // 
            // txbItem1
            // 
            this.txbItem1.Location = new System.Drawing.Point(82, 394);
            this.txbItem1.Name = "txbItem1";
            this.txbItem1.Size = new System.Drawing.Size(100, 20);
            this.txbItem1.TabIndex = 19;
            // 
            // dtDate1
            // 
            this.dtDate1.CustomFormat = "dd/MM/yyyy";
            this.dtDate1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDate1.Location = new System.Drawing.Point(82, 419);
            this.dtDate1.Name = "dtDate1";
            this.dtDate1.Size = new System.Drawing.Size(100, 20);
            this.dtDate1.TabIndex = 23;
            // 
            // txbItem2
            // 
            this.txbItem2.Location = new System.Drawing.Point(255, 397);
            this.txbItem2.Name = "txbItem2";
            this.txbItem2.Size = new System.Drawing.Size(100, 20);
            this.txbItem2.TabIndex = 21;
            // 
            // rdoAll
            // 
            this.rdoAll.AutoSize = true;
            this.rdoAll.Checked = true;
            this.rdoAll.Location = new System.Drawing.Point(16, 17);
            this.rdoAll.Name = "rdoAll";
            this.rdoAll.Size = new System.Drawing.Size(36, 17);
            this.rdoAll.TabIndex = 31;
            this.rdoAll.TabStop = true;
            this.rdoAll.Text = "All";
            this.rdoAll.UseVisualStyleBackColor = true;
            // 
            // rdoW1
            // 
            this.rdoW1.AutoSize = true;
            this.rdoW1.Location = new System.Drawing.Point(53, 17);
            this.rdoW1.Name = "rdoW1";
            this.rdoW1.Size = new System.Drawing.Size(50, 17);
            this.rdoW1.TabIndex = 32;
            this.rdoW1.TabStop = true;
            this.rdoW1.Text = "WH1";
            this.rdoW1.UseVisualStyleBackColor = true;
            // 
            // rdoW2
            // 
            this.rdoW2.AutoSize = true;
            this.rdoW2.Location = new System.Drawing.Point(100, 17);
            this.rdoW2.Name = "rdoW2";
            this.rdoW2.Size = new System.Drawing.Size(50, 17);
            this.rdoW2.TabIndex = 33;
            this.rdoW2.TabStop = true;
            this.rdoW2.Text = "WH2";
            this.rdoW2.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rdoAll);
            this.groupBox2.Controls.Add(this.rdoW2);
            this.groupBox2.Controls.Add(this.rdoW1);
            this.groupBox2.Location = new System.Drawing.Point(239, 7);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(165, 44);
            this.groupBox2.TabIndex = 34;
            this.groupBox2.TabStop = false;
            // 
            // frmReturnTransaction
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(437, 588);
            this.Controls.Add(this.StatusStrip1);
            this.Controls.Add(this.rdoByVoucher);
            this.Controls.Add(this.rdoBySection);
            this.Controls.Add(this.btnGenreport);
            this.Controls.Add(this.GroupBox1);
            this.Name = "frmReturnTransaction";
            this.Text = "frmReturnTransaction";
            this.Load += new System.EventHandler(this.frmReturnTransaction_Load);
            this.StatusStrip1.ResumeLayout(false);
            this.StatusStrip1.PerformLayout();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.StatusStrip StatusStrip1;
        internal System.Windows.Forms.ToolStripStatusLabel ToolStripStatusLabel1;
        internal System.Windows.Forms.RadioButton rdoByVoucher;
        internal System.Windows.Forms.RadioButton rdoBySection;
        internal System.Windows.Forms.Button btnGenreport;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.Label Label7;
        internal System.Windows.Forms.TextBox txbVoucher1;
        internal System.Windows.Forms.TextBox txbVoucher2;
        internal System.Windows.Forms.RadioButton rdoShipment;
        internal System.Windows.Forms.RadioButton rdoReceive;
        internal System.Windows.Forms.Button btnRemoveAllSection;
        internal System.Windows.Forms.Button btnAddAllSection;
        internal System.Windows.Forms.Label Label10;
        internal System.Windows.Forms.Button btnRemoveSection;
        internal System.Windows.Forms.Button btnAddSection;
        internal System.Windows.Forms.ListBox lstSection2;
        internal System.Windows.Forms.ListBox lstSection1;
        internal System.Windows.Forms.Button btnRemoveCategory;
        internal System.Windows.Forms.Button btnAddCategory;
        internal System.Windows.Forms.ListBox lstCategory2;
        internal System.Windows.Forms.ListBox lstCategory1;
        internal System.Windows.Forms.Label Label9;
        internal System.Windows.Forms.Label Label8;
        internal System.Windows.Forms.ComboBox cboFac;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.DateTimePicker dtDate2;
        internal System.Windows.Forms.TextBox txbItem1;
        internal System.Windows.Forms.DateTimePicker dtDate1;
        internal System.Windows.Forms.TextBox txbItem2;
        private System.Windows.Forms.RadioButton rdoW2;
        private System.Windows.Forms.RadioButton rdoW1;
        private System.Windows.Forms.RadioButton rdoAll;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}