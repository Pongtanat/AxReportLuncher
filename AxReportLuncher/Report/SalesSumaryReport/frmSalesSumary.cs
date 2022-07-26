using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.Sales_Sumary_Report
{
    public partial class frmSalesSumary : Form
    {
        public frmSalesSumary()
        {
            InitializeComponent();
        }

        private void frmSalesSumary_Load(object sender, EventArgs e)
        {
            this.Text = "Sales Summary Report";

            string[] arrFactory = { "GMO", "RP", "PO" };
            cboFac.DataSource = arrFactory;

            DateTime dt =  DateTime.Now;
            int [] arrYear = new int[3];


            cboPeriod.Text = dt.Year.ToString();
            for (int i = 0; i < 3; i++)
            {
                arrYear[i] = dt.Year-i;

            }

            cboPeriod.DataSource = arrYear;

           
        }
        private void btnGenreport_Click(object sender, EventArgs e)
        {
            SalesSummaryBLL SalesSummaryBLL = new SalesSummaryBLL();
            SalesSummaryOBJ SalesSummaryOBJ = new SalesSummaryOBJ();
                
            this.Cursor = Cursors.WaitCursor;
            btnGenreport.Enabled = false;


            SalesSummaryOBJ.Factory = cboFac.SelectedItem.ToString();
            SalesSummaryOBJ.DateFrom = new DateTime((int)cboPeriod.SelectedItem, 4, 1);
            SalesSummaryOBJ.DateTo = new DateTime((int)cboPeriod.SelectedItem + 1, 3, 31);
            string strProcess = "";
            btnGenreport.Enabled = false;
            strProcess = SalesSummaryBLL.getSalesSummary(SalesSummaryOBJ);
            btnGenreport.Enabled = true;
            this.Cursor = Cursors.Default;

            
            
        }
    }
}
