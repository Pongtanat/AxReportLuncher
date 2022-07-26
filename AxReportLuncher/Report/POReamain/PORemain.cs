using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewVersion.Report.POReamain
{
    public partial class PORemain : Form
    {
        public PORemain()
        {
            InitializeComponent();
        }

        private void PORemain_Load(object sender, EventArgs e)
        {
            this.MinimumSize = new Size(390, 308);
            this.MaximumSize = new Size(390, 308);

    
            string[] arrFactory = { "GMO", "RP", "PO", "FOS","HO" };
            string[] arrReport ={ "PORemain Report" };
            string[] arrnumberSequengroup = { "DM", "IM" };

         

            DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);
            dtDate1.Value = firstDayOfMonth;
            dtDate2.Value = last_date;


            this.Text = "PORemain Report";

        
            cboFac.DataSource = arrFactory;
            cboReport.DataSource = arrReport;
            cboShpLoc.DataSource = arrnumberSequengroup;

          
          


        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {
            PORemainBLL PORemainBLL = new PORemainBLL();
            PORemainOBJ PORemainOBJ = new PORemainOBJ();


            PORemainOBJ.Factory = cboFac.Text;
            PORemainOBJ.DateFrom = dtDate1.Value;
            PORemainOBJ.DateTo = dtDate2.Value;
            PORemainOBJ.NumberSequenceGroup = cboShpLoc.Text;

            try
            {
                btnGenreport.Enabled = false;
                PORemainBLL.getPORemain(PORemainOBJ);
                btnGenreport.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                btnGenreport.Enabled = false;
            }

        }
    }
}
