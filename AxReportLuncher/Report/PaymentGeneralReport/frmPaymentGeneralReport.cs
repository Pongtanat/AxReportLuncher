using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.PaymentGeneralReport
{
    public partial class frmPaymentGeneralReport : Form
    {
        public frmPaymentGeneralReport()
        {
            InitializeComponent();
        }

        private void frmPaymentGeneralReport_Load(object sender, EventArgs e)
        {

            

            this.Text = "Payment General Report";


            DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);
            dtDate1.Value = firstDayOfMonth;
            dtDate2.Value = last_date;



            string[] arrFactory = { "HO","GMO", "RP", "PO" };
            string[] arrGroup = { "Domestic", "Import", "Material", "Payment" };
            cboFac.DataSource = arrFactory;
            cboGroup.DataSource = arrGroup;


        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {
            PaymentGeneralBLL PaymentGeneralBLL = new PaymentGeneralBLL();
            PaymentGeneralOBJ PaymentGeneralOBJ = new PaymentGeneralOBJ ();


            PaymentGeneralOBJ.Factory = cboFac.Text;
            PaymentGeneralOBJ.DateFrom = dtDate1.Value;
            PaymentGeneralOBJ.DateTo = dtDate2.Value;

            PaymentGeneralOBJ.GroupVoucher = cboGroup.Text;
            PaymentGeneralOBJ.StartVoucher = txtStartVoucher.Text.ToString();
            PaymentGeneralOBJ.EndVoucher = txtEndVoucher.Text.ToString();


            btnGenreport.Enabled = false;
            PaymentGeneralBLL.getPaymentGeneralReport(PaymentGeneralOBJ);
            btnGenreport.Enabled = true;




        }

  
    }
}
