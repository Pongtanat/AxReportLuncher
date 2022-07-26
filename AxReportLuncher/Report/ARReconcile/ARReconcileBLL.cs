using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace NewVersion.Report.ARReconcile
{
    class ARReconcileBLL
    {
        ARReconcileDAL ARReconcileDAL = new ARReconcileDAL();

        public string getARReconcile(ARReconcileOBJ ARReconcileOBJ)
        {
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dt = ARReconcileDAL.getInvoiceAccountCurr(ARReconcileOBJ);
                string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                if (dt.Rows.Count > 0)
                {
                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\ARReconcile\ARReconcile.xls");
                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();
                    xlsSheet = xlsBook.Sheets[3];


                    //==============Detail====================
                   // xlsSheetTemplate = xlsBook.Sheets[3];
                    foreach (DataRow dr in dt.Rows)
                    {
                         ARReconcileOBJ.InvoiceAccount = dr["INVOICEACCOUNT"].ToString();
                         ARReconcileOBJ.CurrencyISO = dr["Curr"].ToString();

                         rsSum = ARReconcileDAL.getInvoiceDetail3Edit(ARReconcileOBJ);

                         if (rsSum.RecordCount > 0)
                         {
                             // xlsSheet.Copy(After: xlsBook.Sheets.Count);
                             // xlsSheet = xlsBook.Sheets[xlsBook.Sheets.Count];

                              xlsSheet = xlsBook.Sheets[xlsBook.Sheets.Count-1];
                              xlsSheet.Copy(After: xlsBook.Sheets[xlsBook.Sheets.Count-1]);





                              //xlsSheet = xlsBook.Sheets[xlsBook.Sheets.Count];
                              //xlsSheet.Copy(Before: xlsBook.Sheets[3]);


                           
                              xlsSheet.Cells.Font.Name = "Arial";
                              xlsSheet.Cells.Font.Size = 8;
                              xlsSheet.Name = ARReconcileOBJ.InvoiceAccount + "-" + dr["Curr"].ToString();
                              xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND) LTD. - {0} ",ARReconcileOBJ.Factory);
                              xlsSheet.Cells[2, 1] = String.Format("A/R Reconcile : Date as of {0:MMMM-yyyy} ", ARReconcileOBJ.DateTo);
                              xlsSheet.Cells[3, 1] = String.Format("{0} - {1}", dr["NAME"].ToString(), dr["Curr"].ToString());



                              xlsSheet.Range[xlsSheet.Cells[(5+1), 1], xlsSheet.Cells[(5+rsSum.RecordCount+1), 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                             
                              
                             xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                             xlsSheet.Range[xlsSheet.Cells[(5 + rsSum.RecordCount), 1], xlsSheet.Cells[(5 + rsSum.RecordCount + 2), 1]].EntireRow.Delete();


                             xlsSheet.Range["B:H"].Columns.EntireColumn.AutoFit();

                         } //end if
                         rsSum.Close();
                    } //end for
                  //  xlsSheetTemplate.Delete();
                    //xlsBook.Sheets[2].Delete();


                    //=============================Header====================
                    xlsSheet = xlsBook.Sheets[2];
                    xlsBook.Sheets[1].activate();
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;
                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND) LTD. - {0} ", ARReconcileOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("A/R Reconcile : Date as of {0:MMMM-yyyy} ", ARReconcileOBJ.DateTo);


                    dt.Clear();
             
                    dt = ARReconcileDAL.getInvoiceDueDate2(ARReconcileOBJ);
                    //rsSum = ARReconcileDAL.getReconcileSummary(dt, ARReconcileOBJ);
                  
                    Excel.Range rangeSource, rangeDest, rowSum;
                    int colGrandTotal = 0;
                    rangeSource = xlsSheet.Range["C3:G7"];
                    xlsSheet.Cells[3, 3] = dt.Rows[0][0];
                    for (int i = 1; i <= (dt.Rows.Count - 1);i++ )
                    {
                        rangeSource.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[3, 8 + ((i - 1) * 5)], xlsSheet.Cells[3, 8 + ((i - 1) * 5)]];
                        rangeDest.Insert(Shift:Excel.XlInsertShiftDirection.xlShiftToRight);
                        //xlsSheet.Cells[3, 3 + (i* 5)] = dt.Rows[i][0];

                        xlsSheet.Cells[3, 3 + (i * 5)] = dt.Rows[i][0];

                        colGrandTotal = 8 + (i * 5);
                    }
              
                       xlsSheet.Range[xlsSheet.Cells[3, colGrandTotal], xlsSheet.Cells[3, (colGrandTotal+4)]].EntireColumn.delete();
                    

                
                    //Rows 


                    rsSum = ARReconcileDAL.getReconcileSummaryDetail(dt, ARReconcileOBJ, "4");
                    rangeSource = xlsSheet.Range["B29:C29"];
                    rowSum = xlsSheet.Range[xlsSheet.Cells[29, 2], xlsSheet.Cells[29, 2]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(29 + 1), 1], xlsSheet.Cells[(29 + rsSum.RecordCount), 1]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(29 + rsSum.RecordCount), 1], xlsSheet.Cells[(29 + rsSum.RecordCount + 1), 1]].EntireRow.Delete();
                    xlsSheet.Range["B" + 29].CopyFromRecordset(rsSum);


                    rsSum = ARReconcileDAL.getReconcileSummaryDetail(dt, ARReconcileOBJ, "3");
                    rangeSource = xlsSheet.Range["B23:C23"];
                    rowSum = xlsSheet.Range[xlsSheet.Cells[23, 2], xlsSheet.Cells[23, 2]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(23 + 1), 1], xlsSheet.Cells[(23 + rsSum.RecordCount), 1]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(23+rsSum.RecordCount), 1], xlsSheet.Cells[(23 +rsSum.RecordCount+1), 1]].EntireRow.Delete();
                    xlsSheet.Range["B" + 23].CopyFromRecordset(rsSum);

                    rsSum = ARReconcileDAL.getReconcileSummaryDetail(dt, ARReconcileOBJ, "2");
                    rangeSource = xlsSheet.Range["B17:C17"];
                    rowSum = xlsSheet.Range[xlsSheet.Cells[17, 2], xlsSheet.Cells[17, 2]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(17 + 1), 1], xlsSheet.Cells[(17 + rsSum.RecordCount), 1]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(17 + rsSum.RecordCount), 1], xlsSheet.Cells[(17 + rsSum.RecordCount + 1), 1]].EntireRow.Delete();
                    xlsSheet.Range["B" + 17].CopyFromRecordset(rsSum);

                    rsSum = ARReconcileDAL.getReconcileSummaryDetail(dt, ARReconcileOBJ, "1");
                    rangeSource = xlsSheet.Range["B11:C11"];
                    rowSum = xlsSheet.Range[xlsSheet.Cells[11, 2], xlsSheet.Cells[11, 2]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(11 + 1), 1], xlsSheet.Cells[(11 + rsSum.RecordCount), 1]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(11 + rsSum.RecordCount), 1], xlsSheet.Cells[(11 + rsSum.RecordCount + 1), 1]].EntireRow.Delete();
                    xlsSheet.Range["B" + 11].CopyFromRecordset(rsSum);

                  //  Excel._Worksheet xlsWorkSheet;

                    rsSum = ARReconcileDAL.getReconcileSummary(dt, ARReconcileOBJ);


                    rangeSource = xlsSheet.Range[xlsSheet.Cells[5, 1], xlsSheet.Cells[5, colGrandTotal]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(6), 1], xlsSheet.Cells[(5 + rsSum.RecordCount), colGrandTotal]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(5 + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(5 + rsSum.RecordCount), colGrandTotal]].EntireRow.delete();
                  

                   // xlsSheet.Range[xlsSheet.Cells[(5 + 1), 1], xlsSheet.Cells[(5+rsSum.RecordCount), 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                   // xlsSheet.Range[xlsSheet.Cells[(5 + rsSum.RecordCount), 1], xlsSheet.Cells[(5 + rsSum.RecordCount + 1), 1]].EntireRow.Delete();

   

                    xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                    xlsSheet.Range["B:CM"].Columns.EntireColumn.AutoFit();
                    //xlsSheet.Range[xlsSheet.Cells[(5), 2], xlsSheet.Cells[(5 + rsSum.RecordCount - 1), 3]].Copy(xlsSheet.Range[xlsSheet.Cells[(rowSum.Row), 2], xlsSheet.Cells[(rowSum.Row + rsSum.RecordCount - 1), 3]]);
                    //xlsSheet.Range["A:H"].Columns.EntireColumn.AutoFit();


                    xlsBook.Sheets[xlsBook.Sheets.Count-1].delete();


                    //Cover
                    xlsSheet = xlsBook.Sheets[xlsBook.Sheets.Count];

                    // ARReconcil Not sales
                     dt = ARReconcileDAL.getInvoiceAccountNum(ARReconcileOBJ);
                    foreach (DataRow dr in dt.Rows)
                    {
                        ARReconcileOBJ.InvoiceAccount = dr["ACCOUNTNUMM"].ToString();
                        //ARReconcileOBJ.CurrencyISO = dr["Curr"].ToString();

                        rsSum = ARReconcileDAL.getInvoiceDetail4New(ARReconcileOBJ);

                        if (rsSum.RecordCount > 0)
                        {
                            // xlsSheet.Copy(After: xlsBook.Sheets.Count);
                            // xlsSheet = xlsBook.Sheets[xlsBook.Sheets.Count];

                            xlsSheet = xlsBook.Sheets[xlsBook.Sheets.Count];
                            xlsSheet.Copy(After: xlsBook.Sheets[xlsBook.Sheets.Count]);

                            xlsSheet.Cells.Font.Name = "Arial";
                            xlsSheet.Cells.Font.Size = 8;
                            xlsSheet.Name = ARReconcileOBJ.InvoiceAccount; //+ "-" + dr["Curr"].ToString();
                            xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND) LTD. - {0} ", ARReconcileOBJ.Factory);
                            xlsSheet.Cells[2, 1] = String.Format("A/R Reconcile : Date as of {0:MMMM-yyyy} ", ARReconcileOBJ.DateTo);
                            //xlsSheet.Cells[3, 1] = String.Format("{0} - {1}", dr["NAME"].ToString(), dr["Curr"].ToString());




                            rangeSource = xlsSheet.Range[xlsSheet.Cells[5, 1], xlsSheet.Cells[5 , 13]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(5+1), 1], xlsSheet.Cells[(5 + rsSum.RecordCount), 13]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);


                         //   xlsSheet.Range[xlsSheet.Cells[(5 + 1), 1], xlsSheet.Cells[(5 + rsSum.RecordCount + 1), 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);


                            xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                            xlsSheet.Range[xlsSheet.Cells[(5 + rsSum.RecordCount), 1], xlsSheet.Cells[(5 + rsSum.RecordCount+1 ), 1]].EntireRow.Delete();


                          
                            int Rows = rsSum.RecordCount + 12;
                            colGrandTotal = 13;
                            rsSum = ARReconcileDAL.getInvoiceAccountCURR(ARReconcileOBJ);
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[Rows, 1], xlsSheet.Cells[Rows, colGrandTotal]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(Rows+1), 1], xlsSheet.Cells[(Rows+rsSum.RecordCount), colGrandTotal]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                            xlsSheet.Range[xlsSheet.Cells[(Rows+rsSum.RecordCount), 1], xlsSheet.Cells[(Rows+rsSum.RecordCount)+1, 1]].EntireRow.Delete();


                            xlsSheet.Range["G" + Rows].CopyFromRecordset(rsSum);
                  



                            xlsSheet.Range["B:H"].Columns.EntireColumn.AutoFit();

                        } //end if
                        rsSum.Close();
                    } //end for
                    //  xlsSheetTemplate.Delete();
                    //xlsBook.Sheets[2].Delete();

                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsSheet = xlsBook.Sheets[1];
                    xlsBook.Sheets[1].activate();
                    xlsSheet.Cells[15, 2] = String.Format("{0} FACTORY", ARReconcileOBJ.Factory);
                    xlsSheet.Cells[18, 2] = String.Format("{0:MMMM-yyyy} ", ARReconcileOBJ.DateTo);

                   // xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;
                  
                    

                }//end dt

            }
            catch (Exception ex)
            {
                return ex.Message;

            }//end tyr

            return null;
        }

    }//end class
}
