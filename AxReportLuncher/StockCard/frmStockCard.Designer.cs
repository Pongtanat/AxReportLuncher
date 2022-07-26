namespace NewVersion.StockCard
{
    partial class frmStockCard
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
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.txtItemID = new System.Windows.Forms.TextBox();
            this.txtGroupID = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.dtDate2 = new System.Windows.Forms.DateTimePicker();
            this.Label5 = new System.Windows.Forms.Label();
            this.dtDate1 = new System.Windows.Forms.DateTimePicker();
            this.cboFac = new System.Windows.Forms.ComboBox();
            this.btnGenreport = new System.Windows.Forms.Button();
            this.GroupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.txtItemID);
            this.GroupBox1.Controls.Add(this.txtGroupID);
            this.GroupBox1.Controls.Add(this.label1);
            this.GroupBox1.Controls.Add(this.Label6);
            this.GroupBox1.Controls.Add(this.Label4);
            this.GroupBox1.Controls.Add(this.Label2);
            this.GroupBox1.Controls.Add(this.dtDate2);
            this.GroupBox1.Controls.Add(this.Label5);
            this.GroupBox1.Controls.Add(this.dtDate1);
            this.GroupBox1.Controls.Add(this.cboFac);
            this.GroupBox1.Location = new System.Drawing.Point(13, 10);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(344, 174);
            this.GroupBox1.TabIndex = 27;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Enter += new System.EventHandler(this.GroupBox1_Enter);
            // 
            // txtItemID
            // 
            this.txtItemID.Location = new System.Drawing.Point(116, 76);
            this.txtItemID.Name = "txtItemID";
            this.txtItemID.Size = new System.Drawing.Size(148, 20);
            this.txtItemID.TabIndex = 38;
            // 
            // txtGroupID
            // 
            this.txtGroupID.Location = new System.Drawing.Point(116, 49);
            this.txtGroupID.Name = "txtGroupID";
            this.txtGroupID.Size = new System.Drawing.Size(148, 20);
            this.txtGroupID.TabIndex = 37;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(62, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 36;
            this.label1.Text = "Item Id :";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(32, 49);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(77, 13);
            this.Label6.TabIndex = 35;
            this.Label6.Text = "Item Group Id :";
            this.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(61, 135);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(48, 13);
            this.Label4.TabIndex = 11;
            this.Label4.Text = "Date to :";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(49, 111);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(59, 13);
            this.Label2.TabIndex = 9;
            this.Label2.Text = "Date from :";
            this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dtDate2
            // 
            this.dtDate2.CustomFormat = "dd/MM/yyyy";
            this.dtDate2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDate2.Location = new System.Drawing.Point(113, 133);
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
            this.dtDate1.Location = new System.Drawing.Point(113, 107);
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
            // btnGenreport
            // 
            this.btnGenreport.Location = new System.Drawing.Point(23, 203);
            this.btnGenreport.Name = "btnGenreport";
            this.btnGenreport.Size = new System.Drawing.Size(327, 49);
            this.btnGenreport.TabIndex = 26;
            this.btnGenreport.Text = "Get Report";
            this.btnGenreport.UseVisualStyleBackColor = true;
            this.btnGenreport.Click += new System.EventHandler(this.btnGenreport_Click);
            // 
            // frmStockCard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(374, 266);
            this.Controls.Add(this.GroupBox1);
            this.Controls.Add(this.btnGenreport);
            this.Name = "frmStockCard";
            this.Text = "frmStockCard";
            this.Load += new System.EventHandler(this.frmStockCard_Load);
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.DateTimePicker dtDate2;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.DateTimePicker dtDate1;
        internal System.Windows.Forms.ComboBox cboFac;
        internal System.Windows.Forms.Button btnGenreport;
        private System.Windows.Forms.TextBox txtItemID;
        private System.Windows.Forms.TextBox txtGroupID;
        internal System.Windows.Forms.Label label1;
    }
}