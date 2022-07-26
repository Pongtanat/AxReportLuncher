namespace NewVersion.Report.ARReconcile
{
    partial class frmARReconcile
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label5 = new System.Windows.Forms.Label();
            this.dtDate2 = new System.Windows.Forms.DateTimePicker();
            this.cboFac = new System.Windows.Forms.ComboBox();
            this.btnGenreport = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Label2);
            this.groupBox1.Controls.Add(this.Label5);
            this.groupBox1.Controls.Add(this.dtDate2);
            this.groupBox1.Controls.Add(this.cboFac);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(352, 94);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Criteria";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(32, 55);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(100, 13);
            this.Label2.TabIndex = 13;
            this.Label2.Text = "Invoice Date as of :";
            this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(81, 26);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(48, 13);
            this.Label5.TabIndex = 11;
            this.Label5.Text = "Factory :";
            this.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // dtDate2
            // 
            this.dtDate2.CustomFormat = "dd/MM/yyyy";
            this.dtDate2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDate2.Location = new System.Drawing.Point(135, 51);
            this.dtDate2.Name = "dtDate2";
            this.dtDate2.Size = new System.Drawing.Size(100, 20);
            this.dtDate2.TabIndex = 14;
            // 
            // cboFac
            // 
            this.cboFac.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFac.FormattingEnabled = true;
            this.cboFac.Location = new System.Drawing.Point(135, 23);
            this.cboFac.Name = "cboFac";
            this.cboFac.Size = new System.Drawing.Size(100, 21);
            this.cboFac.TabIndex = 12;
            // 
            // btnGenreport
            // 
            this.btnGenreport.Location = new System.Drawing.Point(141, 127);
            this.btnGenreport.Name = "btnGenreport";
            this.btnGenreport.Size = new System.Drawing.Size(107, 49);
            this.btnGenreport.TabIndex = 24;
            this.btnGenreport.Text = "Get Report";
            this.btnGenreport.UseVisualStyleBackColor = true;
            this.btnGenreport.Click += new System.EventHandler(this.btnGenreport_Click);
            // 
            // frmARReconcile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(377, 208);
            this.Controls.Add(this.btnGenreport);
            this.Controls.Add(this.groupBox1);
            this.MaximumSize = new System.Drawing.Size(393, 246);
            this.MinimumSize = new System.Drawing.Size(393, 246);
            this.Name = "frmARReconcile";
            this.Text = "frmARReconcile";
            this.Load += new System.EventHandler(this.frmARReconcile_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.DateTimePicker dtDate2;
        internal System.Windows.Forms.ComboBox cboFac;
        internal System.Windows.Forms.Button btnGenreport;
    }
}