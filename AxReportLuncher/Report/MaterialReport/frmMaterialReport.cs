using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace NewVersion.Report.MaterialReport
{
    public partial class frmMaterialReport : Form
    {
        string filePath;
        OpenFileDialog dlg = new OpenFileDialog();
        

        public frmMaterialReport()
        {
            InitializeComponent();
        }

        private void frmMaterialReport_Load(object sender, EventArgs e)
        {
            this.MinimumSize = new Size(387, 380);
            this.MaximumSize = new Size(387,380);

            string[] arrShpLoc = { "None", "RTRD" };
            string[] arrFactory = { "RP", "GMO" ,"PO"};
            string[] arrReport = { "Receive Report", "Shipment Report", "Summary Of Materail Used", "Materail Balance By Item", "Loss Suri", "Material Report", "Summary Materail Balance", "Material Compare", "Movement By item", "Group Material Compare", "Stock Compare", "Materail Purchase"};
            string[] arrCategory = { "All", "Z", "Y", "O", "I"};




            DateTime last_date = new DateTime(dtDate2.Value.Year, dtDate2.Value.Month, DateTime.DaysInMonth(dtDate2.Value.Year, dtDate2.Value.Month));
            DateTime firstDayOfMonth = new DateTime(dtDate1.Value.Year, dtDate1.Value.Month, 1);
            dtDate1.Value = firstDayOfMonth;
            dtDate2.Value = last_date;
            this.Text = "Material Reports";
           

            cboReport.DataSource = arrReport;
            cboFac.DataSource = arrFactory;
            cboCategory.DataSource = arrCategory;
         

        }

        private void btnGenreport_Click(object sender, EventArgs e)
        {
            MaterialBLL MaterialBLL = new MaterialBLL();
            MaterialOBJ MaterialOBJ = new MaterialOBJ();


            MaterialOBJ.Factory = cboFac.Text;
            MaterialOBJ.DateFrom = dtDate1.Value;
            MaterialOBJ.DateTo = dtDate2.Value;
            MaterialOBJ.Category = cboCategory.Text;
           // MaterialOBJ.ShipmentLocation = cboShpLoc.SelectedIndex;

            if (MaterialOBJ.DateFrom > MaterialOBJ.DateTo && dtDate2.Enabled == true && cboReport.SelectedIndex != 3)
            {
               MessageBox.Show("Invalid date range selected.", "error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else
            {

                MaterialOBJ.NumberSequenceGroup = MaterialBLL.getNumberSequenceGroup(MaterialOBJ.Factory, MaterialOBJ.ShipmentLocation);
                if (MaterialOBJ.NumberSequenceGroup == "")
                {
                    // MessageBox.Show(String.Format("Number Sequence Group not found.{0}{0}Factory : {1}{0}{0}Shipment loc : {2}", InvoiceReportOBJ.Factory, cboShpLoc.SelectedItem.ToString), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                /*

                */
                //Choose report
                if (cboReport.SelectedIndex == 0)
                {
                    btnGenreport.Enabled = false;
                    MaterialBLL.getReceiveReport(MaterialOBJ);

                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 1)
                {
                     btnGenreport.Enabled = false;
                     MaterialBLL.getShiptmentReport(MaterialOBJ);
                     btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 2)
                {
                    btnGenreport.Enabled = false;
                    if (MaterialOBJ.Factory == "GMO")
                    {
                       MaterialBLL.getDetailMaterialUSEDForGMO(MaterialOBJ);
                       // MaterialBLL.getMaterialReportMO(MaterialOBJ);
                       // MaterialBLL.getSummaryMaterialMO(MaterialOBJ);
                    }
                    else if(MaterialOBJ.Factory == "RP")
                    {

                        MaterialBLL.getDetailMaterialUSED(MaterialOBJ);


                    }
                    else if (MaterialOBJ.Factory == "PO")
                    {
                        MaterialBLL.getDetailMaterialUSEDForPO(MaterialOBJ);
                    }

                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 3)
                {
                    btnGenreport.Enabled = false;
                    MaterialBLL.getMaterialBalanceByItem(MaterialOBJ);
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 4)
                {
                    btnGenreport.Enabled = false;
                    MaterialBLL.getLossSuri(MaterialOBJ);
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 5)
                {
                    btnGenreport.Enabled = false;
                    if (MaterialOBJ.Factory != "GMO")
                    {
                        MaterialBLL.getMaterialReport(MaterialOBJ);
                    }
                    else
                    {

                        MaterialBLL.getMaterialReportMO(MaterialOBJ);

                    }
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 6)
                {
                    btnGenreport.Enabled = false;
                    MaterialBLL.getSummaryMaterialBalance(MaterialOBJ);
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 7)
                {
                    btnGenreport.Enabled = false;
                    MaterialBLL.getSummaryMaterialCompare(MaterialOBJ);
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 8)
                {
                    btnGenreport.Enabled = false;
                    MaterialBLL.getMaterialMoveMentByItem(MaterialOBJ);
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 9)
                {
                    btnGenreport.Enabled = false;
                    MaterialBLL.getGroupMateCompare(MaterialOBJ);
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 10)
                {
                    btnGenreport.Enabled = false;
                    MaterialBLL.getStockCompare(MaterialOBJ, Path.GetFileName(filePath).ToString());
                    btnGenreport.Enabled = true;
                }
                else if (cboReport.SelectedIndex == 11)
                {
                    btnGenreport.Enabled = false;
                    if (MaterialOBJ.Factory == "RP")
                    {
                        MaterialBLL.getMaterialPurchase(MaterialOBJ);
                    }

                    else if (MaterialOBJ.Factory == "PO")
                    {
                        MaterialBLL.getMaterialPurchaseForPO(MaterialOBJ);

                    }
                    else if (MaterialOBJ.Factory == "GMO")
                    {
                        MaterialBLL.getMaterialPurchaseForGMO(MaterialOBJ);

                    }
                    btnGenreport.Enabled = true;
                }else if(cboReport.SelectedIndex == 12){
                    btnGenreport.Enabled = false;
                    MaterialBLL.getSummaryMaterialMO(MaterialOBJ);
                    btnGenreport.Enabled = true;
                }

            }
        }

        private void btnBrows_Click(object sender, EventArgs e)
        {
            btnBrows.Enabled = false;
            string strSystemPath = System.IO.Directory.GetCurrentDirectory();

            string Folderpath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);


            dlg.Filter = "Excel Files|*.;*.xlsx;*.xls;";
            // dlg.Filter = "CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml|Excel Files|*.;*.xlsx";
            dlg.Multiselect = true;

            try
            {

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    filePath = dlg.FileName;
                    string org, copy;
                    System.IO.StreamReader file = new System.IO.StreamReader(filePath);
                    System.IO.File.Copy(filePath, @"\\192.1.87.244\Mat\Material.xlsx", true);
                    file.Close();
                    btnBrows.Enabled = true;
                }

            

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
/*
        private void cboReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboReport.SelectedIndex == 3 )
            {
                dtDate1.Enabled = false;
            }
            else
            {
                dtDate1.Enabled = true;
            }
        }// end Generate button*/
    }
}
