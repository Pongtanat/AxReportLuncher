using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.QuickSales_Report
{
    public partial class frmQuickSales : Form
    {
        public frmQuickSales()
        {
            InitializeComponent();
        }

        private void frmQuickSales_Load(object sender, EventArgs e)
        {
            this.MinimumSize = new Size(387, 440);
            this.MaximumSize = new Size(387, 440);

            string[] arrShpLoc = { "None"};


            DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);
            dtDate1.Value = firstDayOfMonth;
            dtDate2.Value = last_date;

            this.Text = "QuickSale Reports";
          

            cboShpLoc.DataSource = arrShpLoc;

            string[] arrReport = { "Quick sales report","Support quick sales"};
            cboReport.DataSource = arrReport;

            string[] arrFactory = { "All","GMO", "RP", "PO" };
            cboFac.DataSource = arrFactory;
           
        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {
            QuickSaleReportBLL QuickSaleReportBLL = new QuickSaleReportBLL();
            QuickSaleReportOBJ QuickSaleReportOBJ = new QuickSaleReportOBJ();


            QuickSaleReportOBJ.Factory = cboFac.Text;
            QuickSaleReportOBJ.DateFrom = dtDate1.Value;
            QuickSaleReportOBJ.DateTo = dtDate2.Value;
            QuickSaleReportOBJ.ShipmentLocation = cboShpLoc.SelectedIndex;

            if (QuickSaleReportOBJ.DateFrom > QuickSaleReportOBJ.DateTo && dtDate2.Enabled == true)
            {
                MessageBox.Show("Invalid date range selected.", "error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else
            {

                QuickSaleReportOBJ.NumberSequenceGroup = QuickSaleReportBLL.getNumberSequenceGroup(QuickSaleReportOBJ.Factory, QuickSaleReportOBJ.ShipmentLocation);
                if (QuickSaleReportOBJ.NumberSequenceGroup == "")
                {
                    // MessageBox.Show(String.Format("Number Sequence Group not found.{0}{0}Factory : {1}{0}{0}Shipment loc : {2}", InvoiceReportOBJ.Factory, cboShpLoc.SelectedItem.ToString), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }


                string strChkCustGroup = "";
                bool boolChk = false;
                foreach (CheckBox check in groupBox2.Controls)
                {
                    if (check.Checked)
                    {
                        strChkCustGroup += check.Text + ",";
                        boolChk = true;
                    }
                }

                if (boolChk)
                {
                    strChkCustGroup = strChkCustGroup.Substring(0, strChkCustGroup.Length - 1);
                }
                QuickSaleReportOBJ.CustomerGroup = strChkCustGroup.Replace(",", "','");

                //Choose report
                if (cboReport.SelectedIndex == 1)
                {
                    btnGenreport.Enabled = false;
                    QuickSaleReportBLL.getQuickSalesReportSupport(QuickSaleReportOBJ);
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 0)
                {
                    if (QuickSaleReportOBJ.Factory == "All")
                    {
                        btnGenreport.Enabled = false;
                        QuickSaleReportBLL.getQuickSalesReportAll(QuickSaleReportOBJ);
                        btnGenreport.Enabled = true;

                    }
                    else
                    {
                        btnGenreport.Enabled = false;
                        QuickSaleReportBLL.getQuickSalesReport(QuickSaleReportOBJ);
                       btnGenreport.Enabled = true;

                   }
                   
                }



            }




        }// end onload


      





    }//end class
}
