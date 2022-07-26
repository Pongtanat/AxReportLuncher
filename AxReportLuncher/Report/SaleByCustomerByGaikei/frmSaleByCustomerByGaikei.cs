using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.SaleByCustomerByGaikei
{
    public partial class frmSaleByCustomerByGaikei : Form
    {
        public frmSaleByCustomerByGaikei()
        {
            InitializeComponent();
        }

        private void frmSaleByCustomerByGaikei_Load(object sender, EventArgs e)
        {
            this.MinimumSize = new Size(384, 168);
            this.MaximumSize = new Size(384, 168);

            dtDate.Format = DateTimePickerFormat.Custom;
            dtDate.CustomFormat = "MMMM yyyy";
            dtDate.ShowUpDown = true;

            string[] arrFactory = { "RP1", "RP2" };
            cboFac.DataSource = arrFactory;
        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {

            SalesByGaikeiOBJ SalesByGaikeiOBJ = new SalesByGaikeiOBJ();
            SaleByGaikeiBLL SaleByGaikeiBLL = new SaleByGaikeiBLL();

            SalesByGaikeiOBJ.Factory = cboFac.Text;
            SalesByGaikeiOBJ.dtDate = String.Format("{0:yyyyMM}",dtDate.Value);


            SaleByGaikeiBLL.getRequisitionList(SalesByGaikeiOBJ);


          //  RequisitionOBJ.Factory = cboFac.Text;
          //  RequisitionOBJ.DateFrom = dtDate1.Value;
          //RequisitionOBJ.DateTo = dtDate2.Value;


        }
    }
}
