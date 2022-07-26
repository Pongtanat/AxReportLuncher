using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.APReport
{
    public partial class frmAPReport : Form
    {
        public frmAPReport()
        {
            InitializeComponent();
        }

        private void frmAPReport_Load(object sender, EventArgs e)
        {
            this.Text = "A/P Due Date Report";
            cboFac.Items.Add("-All FACTORY-");
            cboFac.Items.Add("HO");
            cboFac.Items.Add("PO");
            cboFac.Items.Add("RP");
            cboFac.Items.Add("GMO");
            cboFac.SelectedIndex = 1;

            string strFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\HOYA\AXReport.ini";
                
            if(System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(strFile))==false){
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(strFile));
             }

            Common.iniFile iniFile = new Common.iniFile(strFile);
            cboFac.Text = iniFile.GetString("APReport", "fac").Trim();
            txbVendCode.Text = iniFile.GetString("APReport", "vend").Trim();
            dtFrom.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            dtTo.Value = DateTime.Now;

            if (iniFile.GetString("APReport", "d1").Trim() != "")
            {
                dtFrom.Value = new DateTime(Convert.ToInt32(iniFile.GetString("APReport", "d1").Trim())
               ,Convert.ToInt32(iniFile.GetString("APReport", "d1").Trim())
               ,Convert.ToInt32( iniFile.GetString("APReport", "d1").Trim()));
            }

            if (iniFile.GetString("APReport", "d2").Trim() != "")
            {
                dtTo.Value = new DateTime(Convert.ToInt32(iniFile.GetString("APReport", "d2").Trim())
            , Convert.ToInt32(iniFile.GetString("APReport", "d2").Trim())
            , Convert.ToInt32(iniFile.GetString("APReport", "d2").Trim()));

            }

            switch (iniFile.GetString("APReport", "report"))
            {
                case "duedate": rdoDueDate.Checked = true;
                break;
                case "reconcile": rdoDueDate.Checked = true;
                break;
                case "summary": rdoDueDate.Checked = true;
                break;

            }




            //DateTime last_date = new DateTime(dtTo.Value.Year, dtTo.Value.Month, DateTime.DaysInMonth(dtTo.Value.Year, dtTo.Value.Month));
            //DateTime firstDayOfMonth = new DateTime(dtFrom.Value.Year, dtFrom.Value.Month, 1);

           // dtFrom.Value = firstDayOfMonth;
            //dtTo.Value = last_date;

            ReArrangeControl();
            
            APReportBLL getVendGroup = new APReportBLL();




            foreach (DataRow dr in getVendGroup.getVendorGroup(iniFile.GetString("APReport", "vendgroup").Trim().Replace(",", "','")).Rows)
            {
                lstVender1.Items.Add(dr["VendGroup"]);
            }


          //  foreach (string[] str in iniFile.GetString("APReport", "vendgroup").Trim().Split(","))
          //  {
               // lstVender1.Items.Add(dr["VendGroup"]);
          //  }


            lstVender1.Sorted = true;


        }





        private void MoveToRight()
        {
            foreach(string cat in lstVender1.SelectedItems){
                lstVender2.Items.Add(cat);
            }
            foreach (string cat in lstVender2.SelectedItems)
            {
                lstVender1.Items.Add(cat);
            }
        }


        private void MoveToLeft()
        {
            foreach (string cat in lstVender2.SelectedItems)
            {
                lstVender1.Items.Add(cat);
            }
            foreach (string cat in lstVender1.SelectedItems)
            {
                lstVender2.Items.Add(cat);
            }
        }


        private void ReArrangeControl(){
            if (rdoSummay.Checked)
            {
                    lblDateFrom.Text="Date from :";
                    dtTo.Left = lstVender2.Left;
                    dtFrom.Visible = true;
                    lblDateTo.Visible = true;
            }
            else
            {
                dtFrom.Visible = false;
                dtTo.Left = dtFrom.Left;
                lblDateFrom.Text = "Transac. as of :";
                lblDateTo.Visible = false;
            }

        }

        private void btnAddVender_Click(object sender, EventArgs e)
        {
            MoveToRight();
        }

        private void btnRemoveVender_Click(object sender, EventArgs e)
        {
            MoveToLeft();
        }

        private void lstVender1_DoubleClick(object sender, EventArgs e)
        {
            MoveToRight();
        }

        private void lstVender2_DoubleClick(object sender, EventArgs e)
        {
            MoveToLeft();
        }

        private void btnFind_Click(object sender, EventArgs e)
        {

        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {
            APReportBLL APReportBLL = new APReportBLL();
            APReportOBJ APReportOBJ =new APReportOBJ();

            if (cboFac.Text != "")
            {
                this.Cursor = Cursors.WaitCursor;
                btnGenreport.Enabled = false;

                if (cboFac.Text == "-ALL FACTORY")
                {
                    APReportOBJ.Factory = "";
                }
                else
                {
                    APReportOBJ.Factory = cboFac.Text;
                }

                if (lstVender2.Items.Count > 0)
                {
                    foreach (string cat in lstVender2.Items)
                    {
                        APReportOBJ.venderGroup += cat + ",";

                    }
                    APReportOBJ.venderGroup = APReportOBJ.venderGroup.Substring(0, APReportOBJ.venderGroup.Length - 1);

                }

                APReportOBJ.vendercode = txbVendCode.Text;
                APReportOBJ.DateFrom = dtFrom.Value;
                APReportOBJ.DateTo = dtTo.Value;

                string strReport;
                string strProcess;
                if (rdoDueDate.Checked)
                {
                    strReport = "duedate";
                    strProcess = APReportBLL.getAPDueDate(APReportOBJ);

                }
                else if (rdoReconcile.Checked)
                {
                    strReport = "reconcile";
                    strProcess = APReportBLL.getAPSummary(APReportOBJ);
                }
                else
                {
                    strReport = "summary";
                    strProcess = APReportBLL.getAPSummary(APReportOBJ);

                }

                if (strProcess == "")
                {

                }
                else
                {
                    MessageBox.Show(strProcess);

                }
                btnGenreport.Enabled = true;
                this.Cursor = Cursors.Default;


            }



        }//end btnGenerate


    }
}
