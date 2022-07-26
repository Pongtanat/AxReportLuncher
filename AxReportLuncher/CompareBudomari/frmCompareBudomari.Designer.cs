namespace NewVersion.CompareBudomari
{
    partial class frmCompareBudomari
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
            this.lblMessage = new System.Windows.Forms.Label();
            this.btnImport = new System.Windows.Forms.Button();
            this.dtLast = new System.Windows.Forms.DateTimePicker();
            this.dtThis = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblMessage
            // 
            this.lblMessage.AutoSize = true;
            this.lblMessage.Location = new System.Drawing.Point(31, 69);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(25, 13);
            this.lblMessage.TabIndex = 1;
            this.lblMessage.Text = "Null";
            // 
            // btnImport
            // 
            this.btnImport.Enabled = false;
            this.btnImport.Image = global::NewVersion.Properties.Resources.excel;
            this.btnImport.Location = new System.Drawing.Point(252, 61);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(111, 29);
            this.btnImport.TabIndex = 0;
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // dtLast
            // 
            this.dtLast.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtLast.Location = new System.Drawing.Point(75, 31);
            this.dtLast.Name = "dtLast";
            this.dtLast.Size = new System.Drawing.Size(92, 20);
            this.dtLast.TabIndex = 4;
            this.dtLast.ValueChanged += new System.EventHandler(this.dtLast_ValueChanged);
            // 
            // dtThis
            // 
            this.dtThis.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtThis.Location = new System.Drawing.Point(246, 31);
            this.dtThis.Name = "dtThis";
            this.dtThis.Size = new System.Drawing.Size(79, 20);
            this.dtThis.TabIndex = 5;
            this.dtThis.ValueChanged += new System.EventHandler(this.dtThis_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Last Month";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(180, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "This Month";
            // 
            // frmCompareBudomari
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(389, 102);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dtThis);
            this.Controls.Add(this.dtLast);
            this.Controls.Add(this.lblMessage);
            this.Controls.Add(this.btnImport);
            this.Name = "frmCompareBudomari";
            this.Text = "CompareBudomari";
            this.Load += new System.EventHandler(this.frmCompareBudomari_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.DateTimePicker dtLast;
        private System.Windows.Forms.DateTimePicker dtThis;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}