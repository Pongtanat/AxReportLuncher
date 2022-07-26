using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.SalesReturn
{
    public partial class frmSalesReturn : Form
    {
        public frmSalesReturn()
        {
            InitializeComponent();
        }

        private void frmSalesReturn_Load(object sender, EventArgs e)
        {

            this.MinimumSize = new Size(387, 440);
            this.MaximumSize = new Size(387, 440);

            string[] arrShpLoc = { "None", "RTRD"};

            this.Text = "Sales Return Reports";

            DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);

            dtDate1.Value = firstDayOfMonth;
            dtDate2.Value = last_date;


            cboShpLoc.DataSource = arrShpLoc;

            string[] arrReport = { "Sales Return Book", "Sale Return by Items", "Sale Return by Customer", "Defective Receive & Remain Report", "Summary Transaction" };
            cboReport.DataSource = arrReport;

            string[] arrFactory = { "GMO", "RP", "PO" };
            cboFac.DataSource = arrFactory;



        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {
            SaleReturnBLL SaleReturnBLL = new SaleReturnBLL();
            SalesReturnOBJ SaleReturnOBJ = new SalesReturnOBJ();


            SaleReturnOBJ.Factory = cboFac.Text;
            SaleReturnOBJ.DateFrom = dtDate1.Value;
            SaleReturnOBJ.DateTo = dtDate2.Value;
            SaleReturnOBJ.ShipmentLocation = cboShpLoc.SelectedIndex;

            if (SaleReturnOBJ.DateFrom > SaleReturnOBJ.DateTo && dtDate2.Enabled == true)
            {
                MessageBox.Show("Invalid date range selected.", "error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }

            //SaleReturnOBJ.NumberSequenceGroup = SaleReturnBLL.getNumberSequenceGroup(SaleReturnOBJ.Factory, SaleReturnOBJ.ShipmentLocation);
            if (SaleReturnOBJ.NumberSequenceGroup == "")
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




            if (chkShowWH.Checked)
            {
                SaleReturnOBJ.ShowWH = chkShowWH.Checked;

            }
            else
            {
                SaleReturnOBJ.ShowWH = chkShowWH.Checked;
            }



            if (boolChk)
            {
                strChkCustGroup = strChkCustGroup.Substring(0, strChkCustGroup.Length - 1);
            }
            SaleReturnOBJ.CustomerGroup = strChkCustGroup.Replace(",", "','");

            //Choose report
            if (cboReport.SelectedIndex == 0)
            {
                btnGenreport.Enabled = false;
                SaleReturnBLL.getSaleReturnBook(SaleReturnOBJ);
                btnGenreport.Enabled = true;
            }
            else if (cboReport.SelectedIndex == 1)
            {
                btnGenreport.Enabled = false;
                SaleReturnBLL.getSaleReturnByItem(SaleReturnOBJ);
                btnGenreport.Enabled = true;

            }
            else if (cboReport.SelectedIndex == 2)
            {
                btnGenreport.Enabled = false;
                SaleReturnBLL.getSaleReturnByCustomer(SaleReturnOBJ);
                btnGenreport.Enabled = true;
            }
            else if (cboReport.SelectedIndex == 3)
            {
                btnGenreport.Enabled = false;
                SaleReturnBLL.getSalesReturnRemain(SaleReturnOBJ);
                btnGenreport.Enabled = true;
            }
          
        }

        private void cboReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboReport.SelectedIndex == 4)
            {
                SalesReturn.SummaryTransaction.frmReturnTransaction frmReturnTransaction = new SalesReturn.SummaryTransaction.frmReturnTransaction();
                frmReturnTransaction.Show();

            }
        }

    }
}
