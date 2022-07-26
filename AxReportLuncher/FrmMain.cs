using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Net.NetworkInformation;


namespace NewVersion
{


    public partial class FrmMain : Form
    {
        public static string _SECTION;
        public static string _ROLLFAC;
        private string _Username = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
        public static string _Database;

        
        public FrmMain()
        {
            InitializeComponent();
        }


        private void FrmMain_Load(object sender, EventArgs e)
        {

            this.Left = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Right - this.Width;
            string strReportVersion = "";

            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                Version ver = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                strReportVersion = string.Format("{4}, Version: {0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision, Assembly.GetEntryAssembly().GetName().Name);
            }
            else
            {
                var ver = Assembly.GetExecutingAssembly().GetName().Version;
                strReportVersion = string.Format("{4}, Version: {0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision, Assembly.GetEntryAssembly().GetName().Name);
            }


            foreach (ToolStripMenuItem tsm in menuStrip1.Items)
            {
                tsm.Enabled = false;

            }



            string Domain = IPGlobalProperties.GetIPGlobalProperties().DomainName;
           

            AXREportLancherBLL AXReportLauncherBLL = new AXREportLancherBLL();
            DataTable dtMenu = AXReportLauncherBLL.getMenuByUser(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString());
            DataTable dtRoleuser = AXReportLauncherBLL.getRoleuser(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString());

            NewVersion.Material.MaterialOBJ MaterialOBJ = new NewVersion.Material.MaterialOBJ();


            _ROLLFAC = dtRoleuser.Rows[0][1].ToString();


            MaterialOBJ._ROLLFAC = dtRoleuser.Rows[0][1].ToString();
            this.Text = String.Format("{0} - {1} - {2}{3}",_Database, strReportVersion, Domain, _Username);

            try
            {

                if (dtMenu.Rows[0]["WORKERTASKID"].ToString().ToUpper() == "REPORT ALL")
                {
                    foreach (ToolStripMenuItem tsm in menuStrip1.Items)
                    {

                        tsm.Enabled = true;
                        _SECTION = "ALL";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            ToolStripMenuItem ToolstripMenuItem = new ToolStripMenuItem();

            foreach(DataRow dr in dtMenu.Rows){
            switch(dr["WORKERTASKID"].ToString()){

            case "Report Control":
                    accountToolStripMenuItem.Enabled = true;
                    _SECTION = "Control";
                    break;    
            case "Report Finance":
                    financeToolStripMenuItem.Enabled=true  ;
                    _SECTION = "Finance";
                    break;
            case "Report Purchase":
                    purchaseToolStripMenuItem.Enabled=true;
                    _SECTION = "Purchase";
                    break;            
            case "Report INA":
                   internalAuditToolStripMenuItem.Enabled=true;
                    _SECTION = "INA";
                     break;
             case "Report Sales":
                    salesToolStripMenuItem.Enabled=true;
                    _SECTION = "Sales";
                    break;           
            case "Report Warehouse":
                   warehouseToolStripMenuItem.Enabled=true;
                   _SECTION = "Warehouse";
                   break;   
             
               }

           

            }

           





        }//end load


        private void salesReturnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.SalesReturn.frmSalesReturn frmSaleReturn = new NewVersion.Report.SalesReturn.frmSalesReturn();
            frmSaleReturn.Show();
        }

        private void salesSummaryReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.Sales_Sumary_Report.frmSalesSumary frmSalesSummaryReport = new NewVersion.Report.Sales_Sumary_Report.frmSalesSumary();
            frmSalesSummaryReport.Show();
        }

  

        private void quickSalesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.QuickSales_Report.frmQuickSales frmQuickSales = new NewVersion.Report.QuickSales_Report.frmQuickSales();
            frmQuickSales.Show();
        }

        private void invoiceReportToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            NewVersion.Report.InvioceReport.frmInvoiceReport frmInvoiceReport = new NewVersion.Report.InvioceReport.frmInvoiceReport();
            frmInvoiceReport.Show();
        }

        private void saleSummaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.Sales_Sumary_Report.frmSalesSumary frmSaleSummary = new NewVersion.Report.Sales_Sumary_Report.frmSalesSumary();
            frmSaleSummary.Show();
        }


        private void materialReportsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.MaterialReport.frmMaterialReport frmMaterialReport = new NewVersion.Report.MaterialReport.frmMaterialReport();
            frmMaterialReport.Show();
        }

        private void importMaterialToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Material.frmMaterial frmMaterialImport = new NewVersion.Material.frmMaterial();
            frmMaterialImport.Show();
        }

        private void invoiceRetportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.InvioceReport.frmInvoiceReport frmInvoiceReport = new NewVersion.Report.InvioceReport.frmInvoiceReport();
            frmInvoiceReport.Show();
        }

        private void materialToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  NewVersion.Report.InvioceReport.frmInvoiceReport frmInvoiceReport = new NewVersion.Report.InvioceReport.frmInvoiceReport();
           // frmInvoiceReport.Show();
        }

        private void requisitionListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.RequisitionReport.FrmRequistion frmRequisition = new NewVersion.Report.RequisitionReport.FrmRequistion();
            frmRequisition.Show();
        }

        private void aRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.ARReconcile.frmARReconcile frmARReconcile = new NewVersion.Report.ARReconcile.frmARReconcile();
            frmARReconcile.Show();
        }

        private void cToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.CompareBudomari.frmCompareBudomari frmCompareBudomari = new NewVersion.CompareBudomari.frmCompareBudomari();
            frmCompareBudomari.Show();
        }

        private void paymentGeneralReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.PaymentGeneralReport.frmPaymentGeneralReport frmPaymentGeneralReport = new NewVersion.Report.PaymentGeneralReport.frmPaymentGeneralReport();
            frmPaymentGeneralReport.Show();
        }

        private void salesReturnToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            NewVersion.Report.SalesReturn.frmSalesReturn frmSalesReturn  = new NewVersion.Report.SalesReturn.frmSalesReturn();
            frmSalesReturn.Show();
        }

        private void accountPayableReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //NewVersion.Report.APReport.frmAPReport frmAPReport = new NewVersion.Report.APReport.frmAPReport();
           // frmAPReport.Show();
        }

        private void paymentGeneralReportToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            NewVersion.Report.PaymentGeneralReport.frmPaymentGeneralReport frmPaymentGeneralReport = new NewVersion.Report.PaymentGeneralReport.frmPaymentGeneralReport();
            frmPaymentGeneralReport.Show();
        }

        private void stockCardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.StockCard.frmStockCard frmStockCard = new NewVersion.StockCard.frmStockCard();
            frmStockCard.Show();
        }

        private void annualReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewVersion.Report.RequisitionReport.frmAnnualReport frmAnnualReport = new NewVersion.Report.RequisitionReport.frmAnnualReport();
            frmAnnualReport.Show();
        }
         
        private void invoiceReportToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            NewVersion.Report.InvioceReport.frmInvoiceReport frmInvoiceReport = new NewVersion.Report.InvioceReport.frmInvoiceReport();
            frmInvoiceReport.Show();
        }



        private void stockCardToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            NewVersion.StockCard.frmStockCard frmStockCard = new NewVersion.StockCard.frmStockCard();
            frmStockCard.Show();
        }

        private void saleByCustomerByGaikeiToolStripMenuItem_Click(object sender, EventArgs e)
        {
           // SaleByCustomerByGaikei
            NewVersion.Report.SaleByCustomerByGaikei.frmSaleByCustomerByGaikei frmSaleByCustomerByGaikei = new NewVersion.Report.SaleByCustomerByGaikei.frmSaleByCustomerByGaikei();

            frmSaleByCustomerByGaikei.Show();


        }

        private void stockCardToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            NewVersion.StockCard.frmStockCard frmStockCard = new NewVersion.StockCard.frmStockCard();
            frmStockCard.Show();
        }

        private void pORemainToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //NewVersion.Report.POReamain.PORemain frmPORemain = new NewVersion.Report.POReamain.PORemain();
            //frmPORemain.Show();
        }

     
    
 

   

    
    }
}
