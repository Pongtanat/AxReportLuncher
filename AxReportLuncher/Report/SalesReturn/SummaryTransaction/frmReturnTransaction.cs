using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.SalesReturn.SummaryTransaction
{
    public partial class frmReturnTransaction : Form
    {
        public frmReturnTransaction()
        {
            InitializeComponent();
        }

        private void frmReturnTransaction_Load(object sender, EventArgs e)
        {

            this.Text = "Return Summary Transaction Report"; 

            DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);

            dtDate1.Value = firstDayOfMonth;
            dtDate2.Value = last_date;

            string[] arrFactory = { "GMO", "RP", "PO" };
            cboFac.DataSource = arrFactory;

            ReturnTransactionBLL ReturnTransactionBLL = new ReturnTransactionBLL();

            foreach (DataRow dr in ReturnTransactionBLL.getCategoryByType().Rows)
            {
                lstCategory1.Items.Add(dr["ITEMGROUPID"]);
            }

            lstCategory1.Sorted = true;
            lstCategory2.Sorted = true;
            lstSection1.Sorted = true;
            lstSection2.Sorted = true;

        }

        private void cboFac_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstSection1.Items.Clear();
            lstSection2.Items.Clear();

            ReturnTransactionBLL ReturnTransactionBLL = new ReturnTransactionBLL();
            foreach (DataRow dr in ReturnTransactionBLL.getAllSubSectionByFactory(cboFac.Text).Rows)
            {
                lstSection1.Items.Add(dr["SubSection"]);
            }

        }

        private void btnAddCategory_Click(object sender, EventArgs e)
        {
            foreach (string cat in lstCategory1.SelectedItems)
            {
                lstCategory2.Items.Add(cat);
            }

            foreach (string cat in lstCategory2.Items)
            {
                lstCategory1.Items.Remove(cat);
            }
        }

        private void btnRemoveCategory_Click(object sender, EventArgs e)
        {
            foreach (string cat in lstCategory2.SelectedItems)
            {
                lstCategory1.Items.Add(cat);
            }

            foreach (string cat in lstCategory1.Items)
            {
                lstCategory2.Items.Remove(cat);
            }
        }

        private void btnAddSection_Click(object sender, EventArgs e)
        {
            foreach (string cat in lstSection1.SelectedItems)
            {
                lstSection2.Items.Add(cat);
            }

            foreach (string cat in lstSection2.Items)
            {
                lstSection1.Items.Remove(cat);
            }
        }




        private void btnRemoveSection_Click(object sender, EventArgs e)
        {
            foreach (string cat in lstSection2.SelectedItems)
            {
                lstSection1.Items.Add(cat);
            }

            foreach (string cat in lstSection1.Items)
            {
                lstSection2.Items.Remove(cat);
            }
        }

        private void btnAddAllSection_Click(object sender, EventArgs e)
        {
            foreach (string cat in lstSection1.Items)
            {
                lstSection2.Items.Add(cat);
            }

            lstSection1.Items.Clear();
        }

        private void btnRemoveAllSection_Click(object sender, EventArgs e)
        {
            foreach (string cat in lstSection2.Items)
            {
                lstSection1.Items.Add(cat);
            }

            lstSection2.Items.Clear();
        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {
            if (lstCategory2.Items.Count == 0 || lstCategory2.Items.Count==0)
            {
                this.Close();
            }

            ReturnTransactionBLL ReturnTransactionBLL = new ReturnTransactionBLL();
            ReturnTransactionOBJ ReturnTransactionOBJ = new ReturnTransactionOBJ();

            ReturnTransactionOBJ.Factory = cboFac.Text;

            if (rdoReceive.Checked)
            {
                ReturnTransactionOBJ.TransType = 3;
            }
            else
            {
                ReturnTransactionOBJ.TransType = 4;
            }


            foreach (string cat in lstCategory2.Items)
            {
                ReturnTransactionOBJ.Category += cat + ",";
            }


            ReturnTransactionOBJ.Category = ReturnTransactionOBJ.Category.Substring(0, ReturnTransactionOBJ.Category.Length - 1);
          
            foreach (string sec in lstSection2.Items)
            {
                ReturnTransactionOBJ.Section += sec + ",";
            }

            ReturnTransactionOBJ.Section = ReturnTransactionOBJ.Section.Substring(0, ReturnTransactionOBJ.Section.Length - 1);
            ReturnTransactionOBJ.ItemFrom = txbItem1.Text;
            ReturnTransactionOBJ.ItemTo = txbItem2.Text;
            ReturnTransactionOBJ.VoucherFrom = txbVoucher1.Text;
            ReturnTransactionOBJ.VoucherTo = txbVoucher2.Text;
            ReturnTransactionOBJ.DateFrom = dtDate1.Value;
            ReturnTransactionOBJ.DateTo = dtDate2.Value;


            foreach (Control control in this.groupBox2.Controls)
            {
                if (control is RadioButton)
                {
                    RadioButton radio = control as RadioButton;
                    if (radio.Checked)
                    {
                        ReturnTransactionOBJ.WareHouse = radio.Text;
                    }
                }
            }



            if (rdoBySection.Checked)
            {
                btnGenreport.Enabled = false;
                ReturnTransactionBLL.ProcessBySection(ReturnTransactionOBJ);
                btnGenreport.Enabled = true;
            }
            else
            {
                btnGenreport.Enabled = false;
                ReturnTransactionBLL.ProcessByVoucher(ReturnTransactionOBJ);
                btnGenreport.Enabled = true;
            }





        
        
        }




    }
}
