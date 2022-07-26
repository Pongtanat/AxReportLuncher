using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.StockCard
{
    public partial class frmStockCard : Form
    {
        public frmStockCard()
        {
            InitializeComponent();
        }

        private void frmStockCard_Load(object sender, EventArgs e)
        {
            this.MinimumSize = new Size(387, 304);
            this.MaximumSize = new Size(387, 304);

            string[] arrFactory = { "RP", "GMO", "PO","HO" };

            DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);
            dtDate1.Value = firstDayOfMonth;
            dtDate2.Value = last_date;
            this.Text = "StockCard Reports";


            cboFac.DataSource = arrFactory;
 
         

        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {
            StockCardBLL StockCardBLL = new StockCardBLL();
            StockCardOBJ StockCardOBJ = new StockCardOBJ();


            StockCardOBJ.Factory = cboFac.Text;
            StockCardOBJ.DateFrom = dtDate1.Value;
            StockCardOBJ.DateTo = dtDate2.Value;

            StockCardOBJ.GroupID = txtGroupID.Text;
            StockCardOBJ.ItemID = txtItemID.Text;


            btnGenreport.Enabled = false;
            StockCardBLL.getReceiveReport(StockCardOBJ);
            btnGenreport.Enabled = true;





        }

        private void GroupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
