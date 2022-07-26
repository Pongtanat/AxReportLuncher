using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;


namespace NewVersion.Report.RequisitionReport
{
    class RequistionBLL
    {
        RequisitionDAL InvoiceDAL = new RequisitionDAL();

        public string getNumberSequenceGroup(string strFac, int intShipmentLocation)
        {
            string strNumberSequenceGroup = "";
            DataTable dt = InvoiceDAL.getNumberSequenceGroup(strFac, intShipmentLocation);

            if (dt.Rows.Count > 0)
            {
                strNumberSequenceGroup = dt.Rows[0][0].ToString();

            }
            return strNumberSequenceGroup;
        }

        public string getRequisitionList(RequisitionOBJ RequistionOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = RequistionOBJ.DateFrom;



                rsSum = InvoiceDAL.getRequistionList(RequistionOBJ); //External

                if (rsSum.RecordCount > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();


                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    Excel.Range rangeSource, rangeDest;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    int intStartRow = 5;  //StartRow
                    Excel.Workbook xlsBookTemplate;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Requisition\RequisitionList.xlsx");



                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Name = "Requisition List";
                    //xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", RequistionOBJ.Factory, RequistionOBJ.DateFrom, RequistionOBJ.DateTo, RequistionOBJ.CustomerGroup.Replace("','", ", "));

                    xlsSheet.Cells[2, 1] = String.Format("{0:dd-MMM-yyyy} to {1:dd-MMM-yyyy}", RequistionOBJ.DateFrom, RequistionOBJ.DateTo);


                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 6]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 6]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 5]].EntireRow.delete();

                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    //xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 6]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";

                    xlsSheet.Range["A:E"].EntireColumn.AutoFit();


                    rsSum = InvoiceDAL.SummaryByItem(RequistionOBJ); //External

                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Summary By Item";
                    //xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", RequistionOBJ.Factory, RequistionOBJ.DateFrom, RequistionOBJ.DateTo, RequistionOBJ.CustomerGroup.Replace("','", ", "));

                    xlsSheet.Cells[2, 1] = String.Format("{0:dd-MMM-yyyy} to {1:dd-MMM-yyyy}", RequistionOBJ.DateFrom, RequistionOBJ.DateTo);


                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 6]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 6]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 5]].EntireRow.delete();

                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    //xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 6]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";

                    xlsSheet.Range["B:E"].EntireColumn.AutoFit();






                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Invoice Detail



        public string getAnnualReport(RequisitionOBJ RequistionOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = RequistionOBJ.DateFrom;



                rsSum = InvoiceDAL.getAnnualReport(RequistionOBJ); //External

                if (rsSum.RecordCount > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();


                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    Excel.Range rangeSource, rangeDest;
                    Excel.Worksheet xlsSheet;

                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    int intStartRow = 2;  //StartRow
                    Excel.Workbook xlsBookTemplate;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Requisition\Annual.xlsx");

                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsSheet = (Excel.Worksheet)xlsBookTemplate.Sheets[2];
                    //xlsSheet2.Delete();

                    //Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    //Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    //Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    //xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    //xlsBookTemplate.Close();
                  //  xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                   // xlsSheet = xlsBook.Sheets[2];
                    
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    //xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 6]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";

                    xlsSheet.Range["A:G"].EntireColumn.AutoFit();


                   // rsSum = InvoiceDAL.SummaryByItem(RequistionOBJ); //External

                    xlsSheet = (Excel.Worksheet)xlsBookTemplate.Sheets[1];
                    xlsSheet.Name = "AnnualReport(" +RequistionOBJ.Factory+")";
                    //xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", RequistionOBJ.Factory, RequistionOBJ.DateFrom, RequistionOBJ.DateTo, RequistionOBJ.CustomerGroup.Replace("','", ", "));

                    xlsSheet.Cells[2, 1] = String.Format("{0:dd-MMM-yyyy} to {1:dd-MMM-yyyy}", RequistionOBJ.DateFrom, RequistionOBJ.DateTo);

                    xlsBookTemplate.RefreshAll();
                   

                    
                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;    

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Invoice Detail

    }
}
