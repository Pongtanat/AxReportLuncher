using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.InvioceReport
{
    public partial class frmInvoiceReport : Form
    {
        public frmInvoiceReport()
        {
            InitializeComponent();
        }

        private void frmInvoiceReport_Load(object sender, EventArgs e)
        {

            this.MinimumSize = new Size(387, 440);
            this.MaximumSize = new Size(387, 440);

            string[] arrShpLoc = {"None", "Internal", "External", "Material", "Kata", "Trading", "Internal-External" };


            DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);
            dtDate1.Value = firstDayOfMonth;
            dtDate2.Value = last_date;


            this.Text = "Invoice Reports";
    

            cboShpLoc.DataSource = arrShpLoc;



           

            string[] arrReport = new string[] {
                "Summary sale by item"
                , "Detail by group code"
                , "Summary by lentype"
                , "Total sales by currency"
                , "Total sales by customer and currency"
                , "Invoice Detail","Invoice by Date"
                ,"Invoice by Customer"
                ,"Invoice by Item"
                ,"Invoice by Invoice"
                ,"Sales by Customer"};

            cboReport.DataSource = arrReport;
            string[] arrFactory = { "GMO", "RP", "PO", "ALL" };
            cboFac.DataSource = arrFactory;



        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {

            InvoiceBLL InvoiceBLL = new InvoiceBLL();
            InvoiceOBJ InvoiceOBJ = new InvoiceOBJ();


            InvoiceOBJ.Factory = cboFac.Text;
            InvoiceOBJ.DateFrom = dtDate1.Value;
            InvoiceOBJ.DateTo = dtDate2.Value;
            InvoiceOBJ.ShipmentLocation = cboShpLoc.SelectedIndex;

            if (InvoiceOBJ.DateFrom >= InvoiceOBJ.DateTo && dtDate2.Enabled == true)
            {
                MessageBox.Show("Invalid date range selected.", "error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else
            {

                InvoiceOBJ.NumberSequenceGroup = InvoiceBLL.getNumberSequenceGroup(InvoiceOBJ.Factory, InvoiceOBJ.ShipmentLocation);
                if (InvoiceOBJ.NumberSequenceGroup == "")
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
                    InvoiceOBJ.ShowWH = chkShowWH.Checked;

                }
                else
                {
                    InvoiceOBJ.ShowWH = chkShowWH.Checked;
                }



                if (boolChk)
                {
                    strChkCustGroup = strChkCustGroup.Substring(0, strChkCustGroup.Length - 1);
                }
                InvoiceOBJ.CustomerGroup = strChkCustGroup.Replace(",", "','");

                //Choose report
                if (cboReport.SelectedIndex == 0)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION=="ALL")
                    {
                    InvoiceBLL.getSummaryByItem(InvoiceOBJ);
                    }

                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 1)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION == "ALL")
                    {
                        InvoiceBLL.getDetailByGroupCode(InvoiceOBJ);
                    }

                    
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 2)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION == "ALL")
                    {
                        InvoiceBLL.getSummaryByLenType(InvoiceOBJ);
                    }
                   
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 3)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION == "ALL")
                    {
                        InvoiceBLL.getSaleByCurrency(InvoiceOBJ);
                    }
                    
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 4)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION == "ALL")
                    {
                        InvoiceBLL.getSaleByCustomerAndCurrency(InvoiceOBJ);
                    }
                    
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 5)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "INA" || FrmMain._SECTION == "ALL")
                    {
                        InvoiceBLL.getInvoiceDetail(InvoiceOBJ);
                    }
                    
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 6)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION == "ALL" || FrmMain._SECTION == "Sales")
                    {
                        InvoiceBLL.getInvoiceByDate(InvoiceOBJ);
                    }

                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 7)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION == "ALL" || FrmMain._SECTION == "Sales")
                    {
                        InvoiceBLL.getInvoiceByCustomer(InvoiceOBJ);
                    }

                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 8)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION == "ALL")
                    {
                        InvoiceBLL.getInvoiceByItem(InvoiceOBJ);
                    }

                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 9)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION == "ALL")
                    {
                        InvoiceBLL.getInvoiceByInvoice(InvoiceOBJ);
                    }

                    btnGenreport.Enabled = true;
                }

                else if (cboReport.SelectedIndex == 10)
                {
                    btnGenreport.Enabled = false;
                    if (FrmMain._SECTION == "Control" || FrmMain._SECTION == "ALL" || FrmMain._SECTION == "Sales")
                    {
                        InvoiceBLL.getSaleByCustomer(InvoiceOBJ);
                    }

                    btnGenreport.Enabled = true;
                }






            }

        }//end buttom generate
    }
}
