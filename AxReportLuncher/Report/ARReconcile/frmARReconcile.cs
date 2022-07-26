using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.ARReconcile
{
    public partial class frmARReconcile : Form
    {
        public frmARReconcile()
        {
            InitializeComponent();
        }

        private void frmARReconcile_Load(object sender, EventArgs e)
        {
            this.Text = "A/R Reconcile Report";
            string[] arrFactory = { "GMO", "PO", "RP" };
            cboFac.DataSource = arrFactory;

            dtDate2.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1).AddDays(-1);
        

        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {
            ARReconcileBLL ARReconcileBLL = new ARReconcileBLL();
            ARReconcileOBJ ARReconcileOBJ = new ARReconcileOBJ();

            string strOutput; 
            if (cboFac.Text!=""){

                this.Cursor = Cursors.WaitCursor;
                btnGenreport.Enabled = false;

                ARReconcileOBJ.Factory = cboFac.Text;
                ARReconcileOBJ.DateTo = dtDate2.Value;

                string strProcess = "";
                strProcess = ARReconcileBLL.getARReconcile(ARReconcileOBJ);
                if (strProcess == "")
                {
                    MessageBox.Show(strProcess);

                }
                btnGenreport.Enabled = true;
                this.Cursor = Cursors.Default;
            
            
            
            }


        }

   
    }
}
