using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;

namespace NewVersion.CompareBudomari
{
    public partial class frmCompareBudomari : Form

    {

        string filePath;
        OpenFileDialog dlg = new OpenFileDialog();



        public frmCompareBudomari()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            string strSystemPath = System.IO.Directory.GetCurrentDirectory();
            string Folderpath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
           
            CompareBudomari.BudomariBLL BudomariBLL = new BudomariBLL();

            BudomariOBJ BudomariOBJ = new BudomariOBJ();
            BudomariOBJ.DateFrom = dtLast.Value;
            BudomariOBJ.DateTo = dtThis.Value;

            //BudomariBLL.GetBodomari(BudomariOBJ);

            dlg.Filter = "Excel Files|*.;*.xlsx;*.xls;";

            try
            {

                foreach (Process clsProcess in Process.GetProcesses())
                {
                    if (clsProcess.ProcessName.Equals("EXCEL"))
                    {
                        clsProcess.Kill();
                        break;
                    }
                }

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    lblMessage.Text = "Loading ....";

                    filePath = dlg.FileName;
                    System.IO.StreamReader file = new System.IO.StreamReader(filePath);
                    StreamReader objReader = new StreamReader(filePath.ToString());
                    string sLine = "";
                
                    //  CallExcel(arrText, true);


                    if (File.Exists(@"\\192.1.87.242\CompareBudomari\Budomari.xlsx"))
                    {
                        File.Delete(@"\\192.1.87.242\CompareBudomari\Budomari.xlsx");
                    }


                    System.IO.File.Copy(filePath, @"\\192.1.87.242\CompareBudomari\Budomari.xlsx", true);

                    Excel.Application xlsApp = new Excel.Application();
                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;


                    Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(dlg.FileName);
                    Excel._Worksheet xlsWorkSheet = xlsBookTemplate.Sheets[1];

                    xlsWorkSheet = xlsBookTemplate.Sheets[1];

                    BudomariOBJ.GetSheet1 = xlsWorkSheet.Name;
                    xlsWorkSheet = xlsBookTemplate.Sheets[2];
                    BudomariOBJ.GetSheet2 = xlsWorkSheet.Name;

                    file.Close();



                

                   // Generate();
                  //  BudomariBLL.GetBodomari();
                   // Generate();
                    lblMessage.Text = "Processing...";
                    string ERROR = BudomariBLL.GetBodomari(BudomariOBJ);
                    lblMessage.Text = ERROR;
    
                    xlsBookTemplate.Close();
                    xlsApp.Workbooks.Close();
                  



                }

            }
            catch (Exception ex)
            {

                lblMessage.Text = "Error loading";
                foreach (Process clsProcess in Process.GetProcesses())
                {
                    if (clsProcess.ProcessName.Equals("EXCEL"))
                    {
                        clsProcess.Kill();
                        break;
                    }
                }
            }

        }



        private void btnUpdate_Click(object sender, EventArgs e)
        {

            OpenFileDialog dlg = new OpenFileDialog();
            string strSystemPath = System.IO.Directory.GetCurrentDirectory();
            string Folderpath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            Excel.Application xlsApp = new Excel.Application();

            dlg.Filter = "Excel Files|*.;*.xlsx;*.xls;";

            try
            {

                foreach (Process clsProcess in Process.GetProcesses())
                {
                    if (clsProcess.ProcessName.Equals("EXCEL"))
                    {
                        clsProcess.Kill();
                        break;
                    }
                }

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                   // lblMessage.Text = "Loading ....";

                    filePath = dlg.FileName;
                    System.IO.StreamReader file = new System.IO.StreamReader(filePath);
                    StreamReader objReader = new StreamReader(filePath.ToString());
                    string sLine = "";

                    //  CallExcel(arrText, true);


                    if (File.Exists(@"\\192.1.87.242\CompareBudomari\MasterGroup.xlsx"))
                    {
                        File.Delete(@"\\192.1.87.242\CompareBudomari\MasterGroup.xlsx");
                    }


                    System.IO.File.Copy(filePath, @"\\192.1.87.242\CompareBudomari\MasterGroup.xlsx", true);


                    file.Close();

               
                   lblMessage.Text = "Master Updated";

                }

            }
            catch (Exception ex)
            {

                lblMessage.Text = "Error Export";
                foreach (Process clsProcess in Process.GetProcesses())
                {
                    if (clsProcess.ProcessName.Equals("EXCEL"))
                    {
                        clsProcess.Kill();
                        break;
                    }
                }
            }





        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            Excel.Application xlsApp = new Excel.Application();
            xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
            xlsApp.SheetsInNewWorkbook = 1;
            xlsApp.DisplayAlerts = false;
            xlsApp.Visible = false;

            try
            {

                Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(@"\\192.1.87.242\CompareBudomari\MasterGroup.xlsx");
                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[2].delete();
                xlsSheet = xlsBook.Sheets[1];
                xlsSheet.Name = "MasterGroup";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }


            xlsApp.DisplayAlerts = true;
            xlsApp.Visible = true;

        }

        private void dtLast_ValueChanged(object sender, EventArgs e)
        {
            btnImport.Enabled = true;
        }

        private void dtThis_ValueChanged(object sender, EventArgs e)
        {
            btnImport.Enabled = true;
        }

        private void frmCompareBudomari_Load(object sender, EventArgs e)
        {
            dtThis.Format = DateTimePickerFormat.Custom;
            dtThis.CustomFormat = "MM/yyyy";

            dtLast.Format = DateTimePickerFormat.Custom;
            dtLast.CustomFormat = "MM/yyyy";

        }

     






    }//end class
}
