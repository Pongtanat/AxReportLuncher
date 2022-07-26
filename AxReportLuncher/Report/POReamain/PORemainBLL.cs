using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace NewVersion.Report.POReamain
{
    class PORemainBLL
    {
        PORemainDAL PORemainDAL = new PORemainDAL();

        public string getPORemain(PORemainOBJ PORemainOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = PORemainOBJ.DateFrom;



                rsSum = PORemainDAL.getPORemain(PORemainOBJ); //External

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
                    int intStartRow = 4;  //StartRow
                    Excel.Workbook xlsBookTemplate;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\PORemain\PORemain.xls");



                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "PORemain-"+PORemainOBJ.Factory;
                    xlsSheet.Cells[2, 1] = String.Format("{0}-{1} : {2:dd-MMM-yyyy} to {3:dd-MMM-yyyy}", PORemainOBJ.Factory, PORemainOBJ.NumberSequenceGroup,PORemainOBJ.DateFrom, PORemainOBJ.DateFrom);


                    //Row
                    //rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 15]];
                    //rangeSource.EntireRow.Copy();
                    //rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 15]];
                    //rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                   // xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 15]].EntireRow.delete();

                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    // xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 14]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";
                    xlsSheet.Range["A" + (intStartRow), "Y" + (intStartRow+rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Range["B:Y"].EntireColumn.AutoFit();

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
