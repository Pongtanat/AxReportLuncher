using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace NewVersion.Report.APReport
{
    class APReportBLL
    {
        APReportDAL APReportDAL = new APReportDAL();



        public string getAPDueDate(APReportOBJ APReportOBJ)
        {
            try
            {

                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                ADODB.Recordset rsSum = new ADODB.Recordset();
                //DataTable dt = ARReconcileDAL.getInvoiceAccountCurr(ARReconcileOBJ);
                string strSystemPath = System.IO.Directory.GetCurrentDirectory();
                DataTable dt = new DataTable();
                dt = APReportDAL.getVendor(APReportOBJ);

                if (dt.Rows.Count > 0)
                {
                    Excel.Application xlsApp = new Excel.Application();
                    string[] arrfac = {""};
                    if (APReportOBJ.Factory == "")
                    {
                        xlsApp.SheetsInNewWorkbook = 4;
                        arrfac[0] = "HO";
                        arrfac[1] = "RP";
                        arrfac[2] = "PO";
                        arrfac[3] = "GMO";

                    }
                    else
                    {
                        arrfac[0] = APReportOBJ.Factory;
                        xlsApp.SheetsInNewWorkbook = 1;
                    }

                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;

                    //************** new blank workbook *************//
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    //***********************************************//

                    for (int iFac = 0; iFac < (arrfac.Length - 1); iFac++)
                    {
                        xlsSheet = xlsBook.Sheets[(iFac + 1)];

                        xlsSheet.Name = arrfac[iFac];
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;
                        xlsSheet.Range["A1:J1"].Merge();
                        xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND) LTD. - {0} ", arrfac[iFac]);
                        xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].EntireColumn.NumberFormat = "dd/mm/yyyy";
                        xlsSheet.Range[xlsSheet.Cells[1, 4], xlsSheet.Cells[1, 5]].EntireColumn.NumberFormat = "dd/mm/yyyy";
                        xlsSheet.Range[xlsSheet.Cells[1, 10], xlsSheet.Cells[1, 10]].EntireColumn.NumberFormat = "dd/mm/yyyy";
                        xlsSheet.Range[xlsSheet.Cells[1, 7], xlsSheet.Cells[1, 9]].EntireColumn.NumberFormat = "#,##0.00";
                        xlsSheet.Range[xlsSheet.Cells[1, 8], xlsSheet.Cells[1, 8]].EntireColumn.NumberFormat = "#,##0.00000";

                        int iRow = 3;
                        int iRowStartGrandTotal = iRow;
                        foreach(DataRow dr in dt.Rows){
                            APReportOBJ.vendercode = dr["ACCOUNTNUM"].ToString();
                            APReportOBJ.Factory = arrfac[iFac];
                            rsSum = APReportDAL.getAPDueDate(APReportOBJ);

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["A" + iRow + ":J" + iRow].Merge();
                                xlsSheet.Cells[iRow, 1] = APReportOBJ.vendercode + ":" + dr["NAME"];
                                drawHeader((iRow + 1), xlsSheet);
                                xlsSheet.Range[xlsSheet.Cells[(iRow+2+1), 1], xlsSheet.Cells[(iRow+2+rsSum.RecordCount), 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                xlsSheet.Range[xlsSheet.Cells[(iRow + 2 + rsSum.RecordCount), 1], xlsSheet.Cells[(iRow + 2 + rsSum.RecordCount+1), 1]].EntireRow.Delete();
                                xlsSheet.Range["A" + (iRow+2)].CopyFromRecordset(rsSum);
                                iRow += (iRow + 8 + rsSum.RecordCount);
                            }

                            if (dr.Table.Rows.IndexOf(dr) < dt.Rows.Count - 1 && dt.Rows[dr.Table.Rows.IndexOf(dr) + 1]["VendGroup"] != dr["VendGroup"]||dr.Table.Rows.IndexOf(dr)==dt.Rows.Count-1)
                            {
                                drawTotal(iRow, iRowStartGrandTotal, xlsSheet);
                                iRow += 6;
                                iRowStartGrandTotal = iRow + 2;
                            }

                        }// end foreach

                }// end if dr

                    rsSum.Close();
                    rsSum = null;
                    xlsApp.SheetsInNewWorkbook = 3;
                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;
                    return "";

                }
                else
                {
                    return "Not found.";

                }



            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return null;
        }

        public string getAPReconcile(APReportOBJ APReportOBJ)
        {
            try
            {
                    //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    ADODB.Recordset rsSum = new ADODB.Recordset();
                    ADODB.Recordset rsRate = new ADODB.Recordset();
            
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();
                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\APReconcile\APReconcile.xls");
                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[2].delete();
                    xlsSheet = xlsBook.Sheets[1];

                    xlsSheet.Name = String.Format("{0}", "MMMyyyy");
                    xlsSheet.Cells[2, 1] = APReportOBJ.Factory + " FACTORY : ACCOUNT PAYABLE BALANCE";
                   
                rsRate = APReportDAL.getCurrRate(APReportOBJ);
                
                APReportOBJ.venderGroup = "NR-TOV";
                rsSum = APReportDAL.getAPReconcile(APReportOBJ);
                if (rsSum.RecordCount > 0)
                {
                    xlsSheet.Range["N24"].CopyFromRecordset(rsRate);
                    xlsSheet.Range[xlsSheet.Cells[(21), 1], xlsSheet.Cells[(21+rsSum.RecordCount), 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(21+rsSum.RecordCount), 1], xlsSheet.Cells[(21+rsSum.RecordCount+1), 1]].EntireRow.Delete();
                    xlsSheet.Range["N21"].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(21), 7], xlsSheet.Cells[(21 + rsSum.RecordCount -1), 7]].Formula = "=IF(E21=0,0,F21/E21)";
                    xlsSheet.Range[xlsSheet.Cells[(21), 9], xlsSheet.Cells[(21 + rsSum.RecordCount - 1), 9]].Formula = "0";
                    xlsSheet.Range[xlsSheet.Cells[(21), 10], xlsSheet.Cells[(21 + rsSum.RecordCount - 1), 10]].Formula = "0";
                    xlsSheet.Range[xlsSheet.Cells[(21), 11], xlsSheet.Cells[(21 + rsSum.RecordCount - 1), 11]].Formula = "=IF(I21=0,0,J21/I21)";
                    xlsSheet.Range[xlsSheet.Cells[(21), 16], xlsSheet.Cells[(21 + rsSum.RecordCount - 1), 16]].Formula = "=+E21-I21";
                    xlsSheet.Range[xlsSheet.Cells[(21), 17], xlsSheet.Cells[(21 + rsSum.RecordCount - 1), 17]].Formula = "=+F21-J21";
                    xlsSheet.Range[xlsSheet.Cells[(21), 18], xlsSheet.Cells[(21 + rsSum.RecordCount - 1), 18]].Formula = "=IF(P21=0,0,Q21/P21)"; 
                }
                rsSum.Close();


                APReportOBJ.venderGroup = "RL-TOV";
                rsSum = APReportDAL.getAPReconcile(APReportOBJ);
                if (rsSum.RecordCount > 0)
                {
                    xlsSheet.Range["N9"].CopyFromRecordset(rsRate);
                    xlsSheet.Range[xlsSheet.Cells[(6+1), 1], xlsSheet.Cells[(6 + rsSum.RecordCount), 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(6 + rsSum.RecordCount), 1], xlsSheet.Cells[(6 + rsSum.RecordCount + 1), 1]].EntireRow.Delete();
                    xlsSheet.Range["B6"].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(6), 7], xlsSheet.Cells[(6 + rsSum.RecordCount - 1), 7]].Formula = "=IF(E6=0,0,F6/E6)";
                    xlsSheet.Range[xlsSheet.Cells[(6), 9], xlsSheet.Cells[(6 + rsSum.RecordCount - 1), 9]].Formula = "0";
                    xlsSheet.Range[xlsSheet.Cells[(6), 10], xlsSheet.Cells[(6 + rsSum.RecordCount - 1), 10]].Formula = "0";
                    xlsSheet.Range[xlsSheet.Cells[(6), 11], xlsSheet.Cells[(6 + rsSum.RecordCount - 1), 11]].Formula = "=IF(I6=0,0,J6/I6)";
                    xlsSheet.Range[xlsSheet.Cells[(6), 16], xlsSheet.Cells[(6 + rsSum.RecordCount - 1), 16]].Formula = "=+E6-I6";
                    xlsSheet.Range[xlsSheet.Cells[(6), 17], xlsSheet.Cells[(6 + rsSum.RecordCount - 1), 17]].Formula = "=+F6-J6";
                    xlsSheet.Range[xlsSheet.Cells[(6), 18], xlsSheet.Cells[(6 + rsSum.RecordCount - 1), 18]].Formula = "=IF(P6=0,0,Q6/P6)";
                }
                rsSum.Close();
                rsSum = null;
                xlsApp.SheetsInNewWorkbook = 3;
                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;
                return "";


            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return null;
        }

        public string getAPSummary(APReportOBJ APReportOBJ)
        {
            try
            {
                //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                ADODB.Recordset rsSum = new ADODB.Recordset();
                //ADODB.Recordset rsRate = new ADODB.Recordset();

                string strSystemPath = System.IO.Directory.GetCurrentDirectory();
                Excel.Application xlsApp = new Excel.Application();
                System.Globalization.CultureInfo oldCI;
                oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\APReconcile\APSummary.xls");
                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[2].delete();
                xlsSheet = xlsBook.Sheets[1];

           
                rsSum = APReportDAL.getAPSummary(APReportOBJ);

                if (rsSum.RecordCount > 0)
                {
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;
                    xlsSheet.Name = "Summary AP";

                    xlsSheet.Cells[1, 1] = APReportOBJ.Factory + " - Summary A/P " +String.Format("{0:dd/MM/yyyy} to {1:dd/MM/yyyy}", APReportOBJ.DateFrom, APReportOBJ.DateTo);
                    xlsSheet.Range["A1:H1"].Merge();
                    xlsSheet.Range["A1:A1"].HorizontalAlignment = Excel.Constants.xlCenter;
                    xlsSheet.Range["A2:H2"].HorizontalAlignment = Excel.Constants.xlCenter;
                    xlsSheet.Range["A2:H2"].Interior.Color = 15652797;

               
                    xlsSheet.Range[xlsSheet.Cells[(3+1), 1], xlsSheet.Cells[(3+ rsSum.RecordCount), 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(3 + rsSum.RecordCount), 1], xlsSheet.Cells[(3 + rsSum.RecordCount + 1), 1]].EntireRow.Delete();
                    xlsSheet.Range["A3"].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(3), 4], xlsSheet.Cells[(3 + rsSum.RecordCount - 1), 8]].EntireColumn.NumberFormat = "#,##0.00_);[Red](#,##0.00)";

                 
                }
                rsSum.Close();
                rsSum = null;
                xlsApp.SheetsInNewWorkbook = 3;
                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;
                return "";


            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return null;
        }


        public DataTable findVendor(string FieldSearch, string ValueToSearch)
        {
            return APReportDAL.findVendor(FieldSearch, ValueToSearch);
        }

        public DataTable getVendor(APReportOBJ APReportOBJ)
        {
            return APReportDAL.getVendor(APReportOBJ);
        }

        public DataTable getVendorGroup(string strVendGroup)
        {
            return APReportDAL.getVenderGroup(strVendGroup);
        }


        void drawTotal(int Row, int RowStart, Excel.Worksheet xlsSheet)
        {

          //  Color RGB = new Color();
          //  myRgbColor = Color.FromRgb(0, 255, 0);

            Color Color = Color.FromArgb(255, 204, 255);

            xlsSheet.Cells[(Row + 0), 6] = "Grand Total CNY";
            xlsSheet.Cells[(Row + 1), 6] = "Grand Total JPY";
            xlsSheet.Cells[(Row + 2), 6] = "Grand Total THB";
            xlsSheet.Cells[(Row + 3), 6] = "Grand Total USD";
            xlsSheet.Cells[(Row + 4), 6] = "Grand Total";

            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row+4), 1]].Interior.Color = Color;
            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 4), 10]].Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 4), 10]].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 4), 10]].Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 4), 10]].Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 4), 10]].Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;


             xlsSheet.Range[xlsSheet.Cells[Row+0, 7], xlsSheet.Cells[(Row + 0), 7]].Formula = @"=SUMIF($F$" +( RowStart + 2) + ":$F$" +( Row - 2 )+ @",""Total CNY"",$G$" +( RowStart + 2) + ":$G$" +( Row - 2) + ")";//CNY
             xlsSheet.Range[xlsSheet.Cells[(Row + 0), 9], xlsSheet.Cells[(Row + 0), 9]].Formula = @"=SUMIF($F$" + (RowStart + 2) + ":$F$" + (Row - 2) + @",""Total CNY"",$I$" + (RowStart + 2) + ":$I$" + (Row - 2) + ")";//CNY->THB
             xlsSheet.Range[xlsSheet.Cells[(Row + 1), 7], xlsSheet.Cells[(Row + 1), 7]].Formula = @"=SUMIF($F$" + (RowStart + 2) + ":$F$" + (Row - 2) + @",""Total JPY"",$G$" + (RowStart + 2) + ":$G$" + (Row - 2) + ")";//JPY
             xlsSheet.Range[xlsSheet.Cells[(Row + 1), 9], xlsSheet.Cells[(Row + 1), 9]].Formula = @"=SUMIF($F$" + (RowStart + 2) + ":$F$" + (Row - 2) + @",""Total JPY"",$I$" + (RowStart + 2) + ":$I$" + (Row - 2) + ")";//JPY->THB
             xlsSheet.Range[xlsSheet.Cells[(Row + 2), 7], xlsSheet.Cells[(Row + 2), 7]].Formula = @"=SUMIF($F$" + (RowStart + 2) + ":$G$" + (Row - 2) + @",""Total THB"",$G$" + (RowStart + 2) + ":$G$" + (Row - 2) + ")";//THB
             xlsSheet.Range[xlsSheet.Cells[(Row + 2), 9], xlsSheet.Cells[(Row + 2), 9]].Formula = @"=SUMIF($F$" + (RowStart + 2) + ":$G$" + (Row - 2) + @",""Total THB"",$I$" + (RowStart + 2) + ":$I$" + (Row - 2) + ")";//THB -> THB
             xlsSheet.Range[xlsSheet.Cells[(Row + 3), 7], xlsSheet.Cells[(Row + 3), 7]].Formula = @"=SUMIF($F$" + (RowStart + 2) + ":$F$" + (Row - 2) + @",""Total USD"",$G$" + (RowStart + 2) + ":$G$" + (Row - 2) + ")";//USD
             xlsSheet.Range[xlsSheet.Cells[(Row + 3), 9], xlsSheet.Cells[(Row + 3), 9]].Formula = @"=SUMIF($F$" + (RowStart + 2) + ":$F$" + (Row - 2) + @",""Total USD"",$I$" + (RowStart + 2) + ":$I$" + (Row - 2) + ")";//USD -> THB

             xlsSheet.Range[xlsSheet.Cells[(Row + 4), 9], xlsSheet.Cells[(Row + 4), 9]].Formula = @"=SUMIF($F$" + (RowStart + 2) + ":$F$" + (Row - 2) + @",""Total"",$I$" + (RowStart + 2) + ":$I$" + (Row - 2) + ")";
        }

        void drawHeader(int Row, Excel.Worksheet xlsSheet)
        {

            //  Color RGB = new Color();
            //  myRgbColor = Color.FromRgb(0, 255, 0);

            Color Color = Color.FromArgb(252, 213, 180);

            xlsSheet.Cells[(Row), 1] = "Inv.Date";
            xlsSheet.Cells[(Row), 2] = "Invoice No.";
            xlsSheet.Cells[(Row), 3] = "Voucher";
            xlsSheet.Cells[(Row), 4] = "Rcpt.Date";
            xlsSheet.Cells[(Row), 5] = "AWB Date";
            xlsSheet.Cells[(Row), 6] = "Cur.";
            xlsSheet.Cells[(Row), 7] = "Amount Cur.";
            xlsSheet.Cells[(Row), 8] = "Exch Rate";
            xlsSheet.Cells[(Row), 9] = "Amount Baht";
            xlsSheet.Cells[(Row), 10] = "Due Date";

            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row), 10]].Interior.Color = Color;

            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 7), 10]].Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 7), 10]].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 7), 10]].Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 7), 10]].Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            xlsSheet.Range[xlsSheet.Cells[Row, 1], xlsSheet.Cells[(Row + 7), 10]].Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;


            xlsSheet.Cells[(Row + 3), 6] = "Total CNY";
            xlsSheet.Cells[(Row + 4), 6] = "Total JPY";
            xlsSheet.Cells[(Row + 5), 6] = "Total THB";
            xlsSheet.Cells[(Row + 6), 6] = "Total USD";
            xlsSheet.Cells[(Row + 7), 6] = "Total";


            xlsSheet.Range[xlsSheet.Cells[Row + 3, 7], xlsSheet.Cells[(Row + 3), 7]].Formula = @"=SUMIF($F$" + (Row + 1) + ":$F$" + (Row + 2) + @",""CNY"",$G$" + (Row + 1) + ":$G$" + (Row +2) + ")";//CNY
            xlsSheet.Range[xlsSheet.Cells[(Row + 3), 9], xlsSheet.Cells[(Row + 3), 9]].Formula = @"=SUMIF($F$" + (Row + 1) + ":$F$" + (Row + 2) + @",""CNY"",$I$" + (Row + 1) + ":$I$" + (Row + 2) + ")";//CNY->THB
            xlsSheet.Range[xlsSheet.Cells[(Row + 4), 7], xlsSheet.Cells[(Row + 4), 7]].Formula = @"=SUMIF($F$" + (Row + 1) + ":$F$" + (Row + 2) + @",""JPY"",$G$" + (Row + 1) + ":$G$" + (Row + 2) + ")";//JPY
            xlsSheet.Range[xlsSheet.Cells[(Row + 4), 9], xlsSheet.Cells[(Row + 4), 9]].Formula = @"=SUMIF($F$" + (Row + 1) + ":$F$" + (Row + 2) + @",""JPY"",$I$" + (Row + 1) + ":$I$" + (Row + 2) + ")";//JPY->THB
            xlsSheet.Range[xlsSheet.Cells[(Row + 5), 7], xlsSheet.Cells[(Row + 5), 7]].Formula = @"=SUMIF($F$" + (Row + 1) + ":$G$" + (Row + 2) + @",""THB"",$G$" + (Row + 1) + ":$G$" + (Row + 2) + ")";//THB
            xlsSheet.Range[xlsSheet.Cells[(Row + 5), 9], xlsSheet.Cells[(Row + 5), 9]].Formula = @"=SUMIF($F$" + (Row + 1) + ":$G$" + (Row + 2) + @",""THB"",$I$" + (Row + 1) + ":$I$" + (Row + 2) + ")";//THB -> THB
            xlsSheet.Range[xlsSheet.Cells[(Row + 6), 7], xlsSheet.Cells[(Row + 6), 7]].Formula = @"=SUMIF($F$" + (Row + 1) + ":$F$" + (Row + 2) + @",""USD"",$G$" + (Row + 1) + ":$G$" + (Row + 2) + ")";//USD
            xlsSheet.Range[xlsSheet.Cells[(Row + 6), 9], xlsSheet.Cells[(Row + 6), 9]].Formula = @"=SUMIF($F$" + (Row + 1) + ":$F$" + (Row + 2) + @",""USD"",$I$" + (Row + 1) + ":$I$" + (Row + 2) + ")";//USD -> THB

            xlsSheet.Range[xlsSheet.Cells[(Row + 7), 9], xlsSheet.Cells[(Row + 7), 9]].Formula = "=SUM(R[-6]C:R[-5]C)";
        }


    }// end class
}
