using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;


namespace NewVersion.StockCard
{
    class StockCardBLL
    {
        StockCardDAL StockCardDAL = new StockCardDAL();

        public string getReceiveReport(StockCardOBJ StockCardOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));


                rsSum = StockCardDAL.getStockCard(StockCardOBJ); //Com

                if (rsSum.RecordCount > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 6;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;

                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\StockCard\StockCard.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();
                    Excel.Range rangeSource, rangeDest;

                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "StockCard "+StockCardOBJ.GroupID;
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    // xlsSheet.Cells[1, 1] = "Receive Report";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0:dd-MMM-yyyy} to {1:dd-MMM-yyyy}", StockCardOBJ.DateFrom, StockCardOBJ.DateTo);



                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 11]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), 11]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    intStartRow += rsSum.RecordCount + 1;

                    xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 2), 11]].EntireRow.delete();



                    xlsSheet.Range["A:K"].Columns.EntireColumn.AutoFit();

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }
    }
}
