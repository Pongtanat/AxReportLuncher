using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.RequisitionReport
{
    public partial class frmAnnualReport : Form
    {

        RequisitionDAL RequisitionDAL = new RequisitionDAL();
        RequisitionOBJ RequisitionOBJ = new RequisitionOBJ();
        RequistionBLL RequisitionBLL = new RequistionBLL();

        public frmAnnualReport()
        {
            InitializeComponent();
            
        }

        private void frmAnnualReport_Load(object sender, EventArgs e)
        {
            this.MinimumSize = new Size(406, 336);
            this.MaximumSize = new Size(406, 336);

            //string[] arrShpLoc = { "None", "Internal", "External", "Material", "Kata", "Trading", "Internal-External" };



            //DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);
            dtDate1.Value = firstDayOfMonth;
            dtDate1.CustomFormat = "MMMM";
           // dtDate2.Value = last_date;

            this.Text = "Annual Reports";


            string[] arrReport = new string[] { "Annual Report" };

            cboReport.DataSource = arrReport;

            string[] arrFactory = { "GMO", "RP", "PO", "NP1","HO","FOS" };
            cboFac.DataSource = arrFactory;
        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {


            RequisitionOBJ.Factory = cboFac.Text;
            RequisitionOBJ.DateFrom = dtDate1.Value;
           

            if (chkAll.Checked == true)
            {

                RequisitionOBJ.Section = "All";
            }
            else
            {
                RequisitionOBJ.Section = cboSection.Text;

            }
            //   RequisitionOBJ.ShipmentLocation = cboShpLoc.SelectedIndex;

            if (cboReport.SelectedIndex == 0)
            {
                btnGenreport.Enabled = false;

                RequisitionBLL.getAnnualReport(RequisitionOBJ);

            }
            btnGenreport.Enabled = true;



        }
        private void cboFac_SelectedIndexChanged(object sender, EventArgs e)
        {
            RequisitionOBJ.Factory = cboFac.Text;
             
            cboSection.DataSource =  RequisitionDAL.getSection(RequisitionOBJ);
            cboSection.DisplayMember = "Section";
            cboSection.ValueMember = "Section";
            
           
           
            
           
        }
    }
}
