namespace NewVersion.Material
{
    partial class frmMaterial
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMaterial));
            this.Opf = new System.Windows.Forms.OpenFileDialog();
            this.btImport = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblFile = new System.Windows.Forms.Label();
            this.lblType = new System.Windows.Forms.Label();
            this.lblMessage = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.gbSection = new System.Windows.Forms.GroupBox();
            this.rdoGMO = new System.Windows.Forms.RadioButton();
            this.rdoRP = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.gbSection.SuspendLayout();
            this.SuspendLayout();
            // 
            // Opf
            // 
            this.Opf.FileName = "Opf";
            // 
            // btImport
            // 
            this.btImport.Image = global::NewVersion.Properties.Resources.excel;
            this.btImport.Location = new System.Drawing.Point(295, 82);
            this.btImport.Name = "btImport";
            this.btImport.Size = new System.Drawing.Size(36, 31);
            this.btImport.TabIndex = 0;
            this.btImport.UseVisualStyleBackColor = true;
            this.btImport.Click += new System.EventHandler(this.btImport_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(12, 125);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(375, 31);
            this.btnGenerate.TabIndex = 2;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblFile);
            this.groupBox1.Controls.Add(this.lblType);
            this.groupBox1.Controls.Add(this.lblMessage);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(276, 100);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Process";
            // 
            // lblFile
            // 
            this.lblFile.AutoSize = true;
            this.lblFile.Location = new System.Drawing.Point(16, 78);
            this.lblFile.Name = "lblFile";
            this.lblFile.Size = new System.Drawing.Size(0, 13);
            this.lblFile.TabIndex = 5;
            // 
            // lblType
            // 
            this.lblType.AutoSize = true;
            this.lblType.Location = new System.Drawing.Point(61, 27);
            this.lblType.Name = "lblType";
            this.lblType.Size = new System.Drawing.Size(0, 13);
            this.lblType.TabIndex = 4;
            // 
            // lblMessage
            // 
            this.lblMessage.AutoSize = true;
            this.lblMessage.Location = new System.Drawing.Point(10, 51);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(0, 13);
            this.lblMessage.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label2.Location = new System.Drawing.Point(7, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "Type:";
            // 
            // gbSection
            // 
            this.gbSection.Controls.Add(this.rdoGMO);
            this.gbSection.Controls.Add(this.rdoRP);
            this.gbSection.Location = new System.Drawing.Point(295, 13);
            this.gbSection.Name = "gbSection";
            this.gbSection.Size = new System.Drawing.Size(92, 56);
            this.gbSection.TabIndex = 4;
            this.gbSection.TabStop = false;
            this.gbSection.Validated += new System.EventHandler(this.gbSection_Validated);
            // 
            // rdoGMO
            // 
            this.rdoGMO.AutoSize = true;
            this.rdoGMO.Location = new System.Drawing.Point(7, 33);
            this.rdoGMO.Name = "rdoGMO";
            this.rdoGMO.Size = new System.Drawing.Size(50, 17);
            this.rdoGMO.TabIndex = 1;
            this.rdoGMO.Text = "GMO";
            this.rdoGMO.UseVisualStyleBackColor = true;
            // 
            // rdoRP
            // 
            this.rdoRP.AutoSize = true;
            this.rdoRP.Checked = true;
            this.rdoRP.Location = new System.Drawing.Point(7, 11);
            this.rdoRP.Name = "rdoRP";
            this.rdoRP.Size = new System.Drawing.Size(40, 17);
            this.rdoRP.TabIndex = 0;
            this.rdoRP.TabStop = true;
            this.rdoRP.Text = "RP";
            this.rdoRP.UseVisualStyleBackColor = true;
            // 
            // frmMaterial
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(411, 168);
            this.Controls.Add(this.gbSection);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btImport);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmMaterial";
            this.Text = "Material";
            this.Load += new System.EventHandler(this.frmMaterial_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.gbSection.ResumeLayout(false);
            this.gbSection.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog Opf;
        private System.Windows.Forms.Button btImport;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblType;
        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblFile;
        private System.Windows.Forms.GroupBox gbSection;
        private System.Windows.Forms.RadioButton rdoGMO;
        private System.Windows.Forms.RadioButton rdoRP;
    }
}