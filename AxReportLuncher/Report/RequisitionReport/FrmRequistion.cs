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
    public partial class FrmRequistion : Form
    {
        public FrmRequistion()
        {
            InitializeComponent();
        }

        private void FrmRequistion_Load(object sender, EventArgs e)
        {
            this.MinimumSize = new Size(387, 300);
            this.MaximumSize = new Size(387, 300);

            //string[] arrShpLoc = { "None", "Internal", "External", "Material", "Kata", "Trading", "Internal-External" };



            DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);
            dtDate1.Value = firstDayOfMonth;
            dtDate2.Value = last_date;

            this.Text = "Requisition Reports";
          



            string[] arrReport = new string[] { "Requisition List"};

            cboReport.DataSource = arrReport;

            string[] arrFactory = { "GMO", "RP", "PO","FOS"};
            cboFac.DataSource = arrFactory;

            //cboShpLoc.DataSource = arrShpLoc;



        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {

            RequisitionOBJ RequisitionOBJ = new RequisitionOBJ();
            RequistionBLL RequisitionBLL = new RequistionBLL();

            RequisitionOBJ.Factory = cboFac.Text;
            RequisitionOBJ.DateFrom = dtDate1.Value;
            RequisitionOBJ.DateTo = dtDate2.Value;
         //   RequisitionOBJ.ShipmentLocation = cboShpLoc.SelectedIndex;

               if (RequisitionOBJ.DateFrom >= RequisitionOBJ.DateTo && dtDate2.Enabled == true)
            {
                MessageBox.Show("Invalid date range selected.", "error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else
            {

                RequisitionOBJ.NumberSequenceGroup = RequisitionBLL.getNumberSequenceGroup(RequisitionOBJ.Factory, RequisitionOBJ.ShipmentLocation);
                if (RequisitionOBJ.NumberSequenceGroup == "")
                {
                    // MessageBox.Show(String.Format("Number Sequence Group not found.{0}{0}Factory : {1}{0}{0}Shipment loc : {2}", InvoiceReportOBJ.Factory, cboShpLoc.SelectedItem.ToString), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }


                //Choose report
                if (cboReport.SelectedIndex == 0)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Warehouse" || FrmMain._SECTION == "ALL")
                    {
                         RequisitionBLL.getRequisitionList(RequisitionOBJ);
                    }
                    else
                    {
                        MessageBox.Show("Invalid date range selected.", "error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                      

                    }
                }
                btnGenreport.Enabled = true;
                }
               
          



        }

        private void cboReport_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


    }//end button
}
