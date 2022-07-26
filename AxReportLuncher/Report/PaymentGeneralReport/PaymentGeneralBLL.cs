using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace NewVersion.Report.PaymentGeneralReport
{
    class PaymentGeneralBLL
    {
        PaymentGeneralDAL PaymentGeneralDAL = new PaymentGeneralDAL();

        public string getNumberSequenceGroup(string strFac, int intShipmentLocation)
        {
            string strNumberSequenceGroup = "";
            DataTable dt = PaymentGeneralDAL.getNumberSequenceGroup(strFac, intShipmentLocation);

            if (dt.Rows.Count > 0)
            {
                strNumberSequenceGroup = dt.Rows[0][0].ToString();

            }
            return strNumberSequenceGroup;
        }




        public string getPaymentGeneralReport(PaymentGeneralOBJ PaymentGeneralOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = PaymentGeneralOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < PaymentGeneralOBJ.DateTo);



                rsSum = PaymentGeneralDAL.getPaymentGeneral(PaymentGeneralOBJ); //External

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
                    int intStartRow = 3;  //StartRow
                    int Column = 18;
                    Excel.Workbook xlsBookTemplate;
                  
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\PaymentGeneral\PaymentGeneralReport.xlsx");


                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = PaymentGeneralOBJ.GroupVoucher.ToString();
                    xlsSheet.Cells[1, 1] = String.Format("{0} {1} : {2:dd-MMM-yyyy} to {3:dd-MMM-yyyy}", PaymentGeneralOBJ.Factory, PaymentGeneralOBJ.GroupVoucher.ToString(),PaymentGeneralOBJ.DateFrom, PaymentGeneralOBJ.DateTo);
             

     
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();

              
                  
                    int temp = intStartRow + rsSum.RecordCount;
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);


                    xlsSheet.Range["H:V"].EntireColumn.AutoFit();


                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Sale By Group code





    }
}
