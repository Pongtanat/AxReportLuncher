using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using Common.cExcel;

namespace NewVersion.Report.InvioceReport
{
    class InvoiceBLL
    {
        InvoiceDAL InvoiceDAL = new InvoiceDAL();

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


        public string getSummaryByItem(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < InvoiceOBJ.DateTo);

                rsSum = InvoiceDAL.getSummaryByItem(InvoiceOBJ, true); //External

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
                    int indexMonth = 3;
                    int Column = 0;

                    

                    Excel.Workbook xlsBookTemplate;
                    if (InvoiceOBJ.Factory == "GMO")
                    {
                        xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleByItem\SummarySaleByItemGMO.xlsx");

                    }
                    else
                    {
                        xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleByItem\SummarySaleByItem.xlsx");
                    }
                    
                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "SummaryByItem";
                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                   
                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[17, 10]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 3];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }
                        xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 18]].EntireColumn.delete();


                        Column=  (11 + (dtMonthRange.Rows.Count * 8))- ((11 + (dtMonthRange.Rows.Count * 8)) - 16);
                        Column = Column + 11;
                        xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 8) + 4] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);
                        //xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);


                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 11], xlsSheet.Cells[14, 28]].EntireColumn.delete();
                        Column = 10;
                    }


                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1,3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();
                   

                
                   // DateTime thisDate1 = new DateTime(dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1].ToString());
                    //Console.WriteLine("Today is " + thisDate1.ToString("MMMM dd, yyyy") + ".");

                    


                    xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                    intStartRow += rsSum.RecordCount;

                    rsSum = InvoiceDAL.getSummaryByItem(InvoiceOBJ, false); //Internal
                    int temp = 0;

                    if (rsSum.RecordCount > 0)
                    {
                        //Row
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();
                        temp += 2;
                    }
                    else
                    {
                        temp += 4;

                    }

                   // intStartRow += rsSum.RecordCount;
                     temp += intStartRow + rsSum.RecordCount;
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                  

                    rsSum = InvoiceDAL.getSummaryByItemNocome(InvoiceOBJ); //Nocome
                    if (rsSum.RecordCount > 0)
                    {
                        xlsSheet.Range["A" + (temp)].CopyFromRecordset(rsSum);

                    }


                    temp = temp +1;
                    rsSum = InvoiceDAL.getSummaryByItemReturn(InvoiceOBJ); //Return
                    if (rsSum.RecordCount > 0)
                    {
                        xlsSheet.Range["A" + (temp)].CopyFromRecordset(rsSum);

                    }


                    if (InvoiceOBJ.Factory == "GMO")
                    {
                        temp = temp + 2;
                        rsSum = InvoiceDAL.getSummaryByItemTrading(InvoiceOBJ,false); //Trading
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["B" + (temp)].CopyFromRecordset(rsSum);

                        }
                        temp = temp + 1;
                        rsSum = InvoiceDAL.getSummaryByItemTrading(InvoiceOBJ, true); //Trading Return
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["B" + (temp)].CopyFromRecordset(rsSum);

                        }
                    }



                    int salePricePCS = 6;
                    int bahtPCS = 9;
                    int salePriceSET = 7;
                    int bahtSET = 10;
                    temp = temp + 1;
                    
                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[3, indexMonth] = drr[0];
                        indexMonth += 8;

                        xlsSheet.Range[xlsSheet.Cells[5, bahtPCS], xlsSheet.Cells[temp, bahtPCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-5],0)";
                        bahtPCS += 8;

                        xlsSheet.Range[xlsSheet.Cells[5, salePricePCS], xlsSheet.Cells[temp, salePricePCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        salePricePCS += 8;

                        xlsSheet.Range[xlsSheet.Cells[5, salePriceSET], xlsSheet.Cells[temp, salePriceSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-4],0)";
                        salePriceSET += 8;

                        xlsSheet.Range[xlsSheet.Cells[5, bahtSET], xlsSheet.Cells[temp, bahtSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-7],0)";
                        bahtSET += 8;

                    }


                    //============================ Summary by Group Code

                      rsSum = InvoiceDAL.getDetailbyByGroupCode(InvoiceOBJ, true); //External

                      if (rsSum.RecordCount > 0)
                      {
                         // string strSystemPath = System.IO.Directory.GetCurrentDirectory();


                         // Excel.Application xlsApp = new Excel.Application();
                         // System.Globalization.CultureInfo oldCI;
                          oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                          System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                         // Excel.Range rangeSource, rangeDest;


                          xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                          xlsApp.SheetsInNewWorkbook = 1;
                          xlsApp.DisplayAlerts = false;
                          xlsApp.Visible = false;
                           intStartRow = 5;  //StartRow
                           indexMonth = 7;
                           Column = 0;
                         

                          // xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                          //xlsBook = xlsApp.Workbooks.Add();
                         // xlsSheet = xlsBook.Worksheets[4];
                          //xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                         // xlsBookTemplate.Close();
                         // xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                          xlsSheet = xlsBook.Sheets[2];
                          xlsSheet.Name = "Detail By GroupCode";
                          xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                          if (dtMonthRange.Rows.Count > 1)
                          {


                              for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                              {
                                  rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[15, 14]];
                                  rangeSource.EntireColumn.Copy();
                                  rangeDest = xlsSheet.Cells[3, 7];
                                  rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                              }

                              xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[14, 22]].EntireColumn.delete();

                              Column = (15 + (dtMonthRange.Rows.Count * 8)) - ((15 + (dtMonthRange.Rows.Count * 8)) - 16);
                              Column = Column + 15;
                              xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 8) + 8] = "Compare " +String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);

                          }
                          else
                          {
                              xlsSheet.Range[xlsSheet.Cells[3, 15], xlsSheet.Cells[14, 31]].EntireColumn.delete();
                              Column = 14;
                          }


                          //Row
                          rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                          rangeSource.EntireRow.Copy();
                          rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                          rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                          xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();



                          // DateTime thisDate1 = new DateTime(dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1].ToString());
                          //Console.WriteLine("Today is " + thisDate1.ToString("MMMM dd, yyyy") + ".");


                          xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                          intStartRow += rsSum.RecordCount;

                          rsSum = InvoiceDAL.getDetailbyByGroupCode(InvoiceOBJ, false); //NOCOME

                          if (rsSum.RecordCount > 0)
                          {
                              //Row
                              rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                              rangeSource.EntireRow.Copy();
                              rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                              rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                              xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();

                          }

                          // intStartRow += rsSum.RecordCount;
                           temp = intStartRow + rsSum.RecordCount;
                          xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                         



                          if (InvoiceOBJ.Factory == "GMO")
                          {
                              rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, false, false, "DetailbyLentype"); //Return
                              if (rsSum.RecordCount > 0)
                              {
                                  xlsSheet.Range["G" + (temp+1)].CopyFromRecordset(rsSum);

                              }

                              temp = temp + 3;
                              rsSum = InvoiceDAL.getSummaryByItemTrading(InvoiceOBJ,false); //Trading

                              if (rsSum.RecordCount > 0)
                              {
                                  xlsSheet.Range["F" + (temp)].CopyFromRecordset(rsSum);

                              }

                              temp = temp + 1;
                              rsSum = InvoiceDAL.getSummaryByItemTrading(InvoiceOBJ, true); //Trding Return
                             
                              if (rsSum.RecordCount > 0)
                              {
                                  xlsSheet.Range["F" + (temp)].CopyFromRecordset(rsSum);

                              }

                          }



                           salePricePCS = 10;
                           bahtPCS = 13;
                           salePriceSET = 11;
                           bahtSET = 14;
                          temp = temp + 1;

                          foreach (DataRow drr in dtMonthRange.Rows)
                          {
                              xlsSheet.Cells[3, indexMonth] = drr[0];
                              indexMonth += 8;

                              xlsSheet.Range[xlsSheet.Cells[5, bahtPCS], xlsSheet.Cells[temp, bahtPCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-5],0)";
                              bahtPCS += 8;

                              xlsSheet.Range[xlsSheet.Cells[5, salePricePCS], xlsSheet.Cells[temp, salePricePCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                              salePricePCS += 8;

                              xlsSheet.Range[xlsSheet.Cells[5, salePriceSET], xlsSheet.Cells[temp, salePriceSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-4],0)";
                              salePriceSET += 8;

                              xlsSheet.Range[xlsSheet.Cells[5, bahtSET], xlsSheet.Cells[temp, bahtSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-7],0)";
                              bahtSET += 8;

                          }

                      }

                      //============================ Summary by LenType==========
                      if (InvoiceOBJ.Factory == "GMO")
                      {
                          rsSum = InvoiceDAL.getSummarybyLenType(InvoiceOBJ, true); //Com

                          if (rsSum.RecordCount > 0)
                          {
                              //string strSystemPath = System.IO.Directory.GetCurrentDirectory();


                              //Excel.Application xlsApp = new Excel.Application();
                              //System.Globalization.CultureInfo oldCI;
                              //oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                              //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                              //Excel.Range rangeSource, rangeDest;


                              xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                              xlsApp.SheetsInNewWorkbook = 1;
                              xlsApp.DisplayAlerts = false;
                              xlsApp.Visible = false;
                              intStartRow = 5;  //StartRow
                              indexMonth = 2;
                              Column = 0;

                              //  Excel.Workbook xlsBookTemplate;

                              //xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleByLenType\SummarySaleByLentype.xlsx");

                              //Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                              //Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                              //Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                              // xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                              // xlsBookTemplate.Close();
                              // xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                              xlsSheet = xlsBook.Sheets[3];
                              xlsSheet.Name = "SummaryByLenType";
                              xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));

                              if (dtMonthRange.Rows.Count > 1)
                              {
                                  //Column
                                  for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                                  {
                                      rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 2], xlsSheet.Cells[15, 4]];
                                      rangeSource.EntireColumn.Copy();

                                      rangeDest = xlsSheet.Cells[3, 2];
                                      rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                                  }

                                  xlsSheet.Range[xlsSheet.Cells[3, 2], xlsSheet.Cells[10, 7]].EntireColumn.delete();


                                  Column = (5 + (dtMonthRange.Rows.Count * 3)) - ((5 + (dtMonthRange.Rows.Count * 3)) - 6);
                                  Column = Column + 5;
                                  xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 3) + 3] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);


                              }
                              else
                              {
                                  xlsSheet.Range[xlsSheet.Cells[3, 5], xlsSheet.Cells[15, 11]].EntireColumn.delete();
                                  Column = 4;
                              }


                              //Row
                              rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                              rangeSource.EntireRow.Copy();
                              rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                              rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                              xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();

                              xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                              intStartRow += rsSum.RecordCount;


                              rsSum = InvoiceDAL.getSummarybyLenType(InvoiceOBJ, false); //Nocome
                              if (rsSum.RecordCount > 0)
                              {
                                  xlsSheet.Range["A" + (intStartRow)].CopyFromRecordset(rsSum);

                              }

                              temp = intStartRow + 3;


                              rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, false, false,"SummaryBylentype"); //Return
                              if (rsSum.RecordCount > 0)
                              {
                                  xlsSheet.Range["B" + (temp)].CopyFromRecordset(rsSum);

                              }


                              if (InvoiceOBJ.Factory == "GMO")
                              {
                                  temp = temp + 2;
                                  rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, true, false,"SummaryBylentype");  //Trading No return
                                  if (rsSum.RecordCount > 0)
                                  {
                                      xlsSheet.Range["B" + (temp)].CopyFromRecordset(rsSum);

                                  }

                                  temp = temp + 1;
                                  rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, true, true,"SummaryBylentype");  //Trading return
                                  if (rsSum.RecordCount > 0)
                                  {
                                      xlsSheet.Range["B" + (temp)].CopyFromRecordset(rsSum);

                                  }

                              }


                              bahtPCS = 4;
                              temp = temp + 1;

                              foreach (DataRow drr in dtMonthRange.Rows)
                              {
                                  xlsSheet.Cells[3, indexMonth] = drr[0];
                                  indexMonth += 3;

                                  xlsSheet.Range[xlsSheet.Cells[5, bahtPCS], xlsSheet.Cells[temp, bahtPCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                                  bahtPCS += 3;
                              }

                              //========================= Detail =========================================//

                              intStartRow = 5;
                              xlsSheet = xlsBook.Sheets[4];
                              xlsSheet.Name = "DetailByLenType";
                              xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                              rsSum = InvoiceDAL.getDetailbyLenType(InvoiceOBJ, true); //Come
                              if (dtMonthRange.Rows.Count > 1)
                              {


                                  //Column
                                  for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                                  {
                                      rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[15, 14]];
                                      rangeSource.EntireColumn.Copy();

                                      rangeDest = xlsSheet.Cells[3, 7];
                                      rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                                  }

                                  xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[10, 22]].EntireColumn.delete();

                                  Column = (15 + (dtMonthRange.Rows.Count * 8)) - ((15 + (dtMonthRange.Rows.Count * 8)) - 16);
                                  Column = Column + 15;
                                  xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 8) + 6] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);


                              }
                              else
                              {
                                  xlsSheet.Range[xlsSheet.Cells[3, 15], xlsSheet.Cells[14, 31]].EntireColumn.delete();
                                  Column = 12;
                              }


                              //Row
                              rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                              rangeSource.EntireRow.Copy();
                              rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                              rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                              xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();


                              xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                              intStartRow += rsSum.RecordCount;

                              rsSum = InvoiceDAL.getDetailbyLenType(InvoiceOBJ, false); //Nocome

                              if (rsSum.RecordCount > 0)
                              {
                                  //Row
                                  rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                                  rangeSource.EntireRow.Copy();
                                  rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                                  rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                  xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();

                                  xlsSheet.Range["A" + (intStartRow)].CopyFromRecordset(rsSum);
                              }


                              temp = rsSum.RecordCount + intStartRow;

                              /*
                              if (InvoiceOBJ.Factory == "GMO")
                              {

                                  rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, true, false,"DetailbyLentype");  //Trading - No return
                                  if (rsSum.RecordCount > 0)
                                  {
                                      xlsSheet.Range["G" + (temp)].CopyFromRecordset(rsSum);

                                  }
                                  temp = temp + 1;
                                  rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, true, true,"DetailbyLentype");  //Trading return
                                  if (rsSum.RecordCount > 0)
                                  {
                                      xlsSheet.Range["G" + (temp)].CopyFromRecordset(rsSum);

                                  }
                              }
                              */

                              salePricePCS = 10;
                              bahtPCS = 13;
                              salePriceSET = 11;
                              bahtSET = 14;
                              indexMonth = 7;
                              foreach (DataRow drr in dtMonthRange.Rows)
                              {
                                  xlsSheet.Cells[3, indexMonth] = drr[0];
                                  indexMonth += 8;

                                  xlsSheet.Range[xlsSheet.Cells[5, bahtPCS], xlsSheet.Cells[temp, bahtPCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-5],0)";
                                  bahtPCS += 8;

                                  xlsSheet.Range[xlsSheet.Cells[5, salePricePCS], xlsSheet.Cells[temp, salePricePCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                                  salePricePCS += 8;

                                  xlsSheet.Range[xlsSheet.Cells[5, salePriceSET], xlsSheet.Cells[temp, salePriceSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-4],0)";
                                  salePriceSET += 8;

                                  xlsSheet.Range[xlsSheet.Cells[5, bahtSET], xlsSheet.Cells[temp, bahtSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-7],0)";
                                  bahtSET += 8;

                              }

                          }
                      }// sale by lentype 



                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;
                  
                   
                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Summary sale by item


        public string getDetailByGroupCode(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < InvoiceOBJ.DateTo);

                rsSum = InvoiceDAL.getDetailbyByGroupCode(InvoiceOBJ, true); //External

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
                    int indexMonth = 7;
                    int Column = 0;
                    Excel.Workbook xlsBookTemplate;
                    if (InvoiceOBJ.Factory == "GMO")
                    {
                        xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleByGroupCode\DetailbyItemByGroupCodeGMO.xlsx");

                    }
                    else
                    {
                        xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleByGroupCode\DetailbyItemByGroupCode.xlsx");
                    }

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Detail By GroupCode";
                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                       

                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[15, 14]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 7];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }



                        xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[14, 22]].EntireColumn.delete();
                         Column = (15 + (dtMonthRange.Rows.Count * 8)) - ((15 + (dtMonthRange.Rows.Count * 8)) - 16);
                        Column = Column + 15;
                        xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 8) + 8] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);


                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 15], xlsSheet.Cells[14, 31]].EntireColumn.delete();
                        Column = 14;
                    }


                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();


                  
                    // DateTime thisDate1 = new DateTime(dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1].ToString());
                    //Console.WriteLine("Today is " + thisDate1.ToString("MMMM dd, yyyy") + ".");

                
                    xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                    intStartRow += rsSum.RecordCount;

                    rsSum = InvoiceDAL.getDetailbyByGroupCode(InvoiceOBJ, false); //NOCOME

                    if (rsSum.RecordCount > 0)
                    {
                        //Row
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();

                    }

                    // intStartRow += rsSum.RecordCount;
                    int temp = intStartRow + rsSum.RecordCount;
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);



                    if (InvoiceOBJ.Factory == "GMO"){

                       temp = temp + 1;
                       rsSum = InvoiceDAL.getSummaryByItemTrading(InvoiceOBJ,false); //Trading

                    if (rsSum.RecordCount > 0)
                    {
                         xlsSheet.Range["F" + (temp)].CopyFromRecordset(rsSum);

                      }

                  }




                     int salePricePCS = 10;
                     int  bahtPCS  = 13;
                     int  salePriceSET  = 11;
                     int  bahtSET = 14;
                     temp = temp + 1;
        
                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[3, indexMonth] = drr[0];
                        indexMonth += 8;

                        xlsSheet.Range[xlsSheet.Cells[5, bahtPCS], xlsSheet.Cells[temp, bahtPCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-5],0)";
                        bahtPCS += 8;

                        xlsSheet.Range[xlsSheet.Cells[5, salePricePCS], xlsSheet.Cells[temp, salePricePCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        salePricePCS += 8;

                        xlsSheet.Range[xlsSheet.Cells[5, salePriceSET], xlsSheet.Cells[temp, salePriceSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-4],0)";
                        salePriceSET += 8;

                        xlsSheet.Range[xlsSheet.Cells[5, bahtSET], xlsSheet.Cells[temp, bahtSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-7],0)";
                        bahtSET += 8;
                    
                    }



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


        public string getSummaryByLenType(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < InvoiceOBJ.DateTo);

                rsSum = InvoiceDAL.getSummarybyLenType(InvoiceOBJ, true); //Com

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
                    int indexMonth = 2;
                    int Column = 0;
                    int temp;
                    int bahtPCS;
                    Excel.Workbook xlsBookTemplate;
                
                     xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleByLenType\SummarySaleByLentype.xlsx");
               
                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    
                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "SummaryByLenType";
                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                   
                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column


                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 2], xlsSheet.Cells[15, 4]];
                            rangeSource.EntireColumn.Copy();

                            rangeDest = xlsSheet.Cells[3, 2];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[3, 2], xlsSheet.Cells[10, 7]].EntireColumn.delete();


                     
                        //xlsSheet.Range[xlsSheet.Cells[3, (5 + (dtMonthRange.Rows.Count * 3)) - 6], xlsSheet.Cells[14, (4 + (dtMonthRange.Rows.Count * 3)) - 3]].EntireColumn.delete();
                        Column = (5 + (dtMonthRange.Rows.Count * 3)) - ((5 + (dtMonthRange.Rows.Count * 3)) - 6);
                        Column = Column + 5;
                        xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 3) + 3] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);


                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 5], xlsSheet.Cells[15, 11]].EntireColumn.delete();
                        Column = 4;
                    }


                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();

                    xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                    intStartRow += rsSum.RecordCount;


                    rsSum = InvoiceDAL.getSummarybyLenType(InvoiceOBJ,false); //Nocome
                     if (rsSum.RecordCount > 0)
                     {
                         xlsSheet.Range["A" + (intStartRow)].CopyFromRecordset(rsSum);

                     }

                     temp = intStartRow + 3;


                     rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, false, false, "Summary"); //Return
                     if (rsSum.RecordCount > 0)
                     {
                         xlsSheet.Range["B" + (temp)].CopyFromRecordset(rsSum);

                     }


                       if (InvoiceOBJ.Factory == "GMO")
                      {
                          temp = temp + 2;
                          rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, true,false,"Summary");  //Trading No return
                          if (rsSum.RecordCount > 0)
                        {
                             xlsSheet.Range["B" + (temp)].CopyFromRecordset(rsSum);

                         }

                          temp = temp + 1;
                          rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, true, true,"Summary");  //Trading return
                          if (rsSum.RecordCount > 0)
                          {
                              xlsSheet.Range["B" + (temp)].CopyFromRecordset(rsSum);

                          }

                      }


                    bahtPCS = 4;
                    temp = temp + 1;

                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[3, indexMonth] = drr[0];
                        indexMonth += 3;

                        xlsSheet.Range[xlsSheet.Cells[5, bahtPCS], xlsSheet.Cells[temp, bahtPCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        bahtPCS += 3;
                    }
                    
        //========================= Detail =========================================//

                    intStartRow = 5;
                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Name = "DetailByLenType";
                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                    rsSum = InvoiceDAL.getDetailbyLenType(InvoiceOBJ, true); //Come
                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[12, 14]];
                            rangeSource.EntireColumn.Copy();

                            rangeDest = xlsSheet.Cells[3, 7];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[10, 22]].EntireColumn.delete();

                  
                        Column = (15 + (dtMonthRange.Rows.Count * 8)) - ((15 + (dtMonthRange.Rows.Count * 8)) - 16);
                        Column = Column + 15;
                        xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 8) + 8] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);


                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 15], xlsSheet.Cells[14, 31]].EntireColumn.delete();
                        Column = 12;
                    }


                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();


                    xlsSheet.Range["A" + 5].CopyFromRecordset(rsSum);
                    intStartRow += rsSum.RecordCount;
                 
                    rsSum = InvoiceDAL.getDetailbyLenType(InvoiceOBJ, false); //Nocome

                    if (rsSum.RecordCount > 0)
                    {
                        //Row
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]].EntireRow.delete();

                        xlsSheet.Range["A" + (intStartRow)].CopyFromRecordset(rsSum);
                    }


                    temp = rsSum.RecordCount + intStartRow;

                    /*

                    if (InvoiceOBJ.Factory == "GMO")
                    {

                        rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, true, false, "DetailbyLentype");  //Trading - No return
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["G" + (temp)].CopyFromRecordset(rsSum);

                        }
                        temp = temp + 1;
                        rsSum = InvoiceDAL.getSummarybyLenTypeReturnAndTrading(InvoiceOBJ, true, true,"DetailbyLentype");  //Trading return
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["G" + (temp)].CopyFromRecordset(rsSum);

                        }
                    }
                    */

                     int salePricePCS = 10;
                     bahtPCS = 13;
                     int salePriceSET = 11;
                     int bahtSET = 14;
                     //temp = temp + 1;
                     indexMonth = 7;
                     foreach (DataRow drr in dtMonthRange.Rows)
                     {
                         xlsSheet.Cells[3, indexMonth] = drr[0];
                         indexMonth += 8;

                         xlsSheet.Range[xlsSheet.Cells[5, bahtPCS], xlsSheet.Cells[temp, bahtPCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-5],0)";
                         bahtPCS += 8;

                         xlsSheet.Range[xlsSheet.Cells[5, salePricePCS], xlsSheet.Cells[temp, salePricePCS]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                         salePricePCS += 8;

                         xlsSheet.Range[xlsSheet.Cells[5, salePriceSET], xlsSheet.Cells[temp, salePriceSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-4],0)";
                         salePriceSET += 8;

                         xlsSheet.Range[xlsSheet.Cells[5, bahtSET], xlsSheet.Cells[temp, bahtSET]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-7],0)";
                         bahtSET += 8;

                     }

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Summary sale by Lentype


        public string getSaleByCurrency(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum,rsSumReturn= new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < InvoiceOBJ.DateTo);

                //rsSum = InvoiceDAL.getDetailbyByGroupCode(InvoiceOBJ, true); //External

            
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
                    int indexMonth = 3;
                    Excel.Workbook xlsBookTemplate;


                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleByCustomerAndCurrency\SalesByCurrency.xls");

                 
                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    switch (InvoiceOBJ.Factory)
                    {
                        case "GMO":
                            xlsSheet = xlsBook.Sheets[2];
                            break;
                        case "ALL":
                            xlsSheet = xlsBook.Sheets[1];
                             break;
                         default:
                             xlsSheet = xlsBook.Sheets[3];
                            break;
                    }


                    xlsSheet.Name = InvoiceOBJ.Factory + "-FACTORY";
                    xlsSheet.Cells[3, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));

                    int[] arrRow = { 7,12,17 };
                    //int[] arrHead = {4, 24, 44 };
                    int iRowNetSales = arrRow[0];
                    int iRowSale = arrRow[0];
                    int Sset,Ppcs;


                  
                    
                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                        xlsSheet.Cells[5, 21] = "TOTAL " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];


                       //xlsSheet.Cells[5, 21] = "TOTAL " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1]);

                        for (int i = 0; i < dtMonthRange.Rows.Count-1; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[5, 3], xlsSheet.Cells[105, 11]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[5, 12];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[5, 12], xlsSheet.Cells[105, 20]].EntireColumn.delete();
                    }
                    else
                    {
                        xlsSheet.Cells[5, 21] = "TOTAL - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];

                        xlsSheet.Range[xlsSheet.Cells[5, 12], xlsSheet.Cells[105, 20]].EntireColumn.delete();
                        //Column = 20;
                    }



                    if (InvoiceOBJ.Factory != "GMO" && InvoiceOBJ.Factory!="ALL")
                    {
                       // arrHead = new[] { 4, 21, 41, 61, 81 };
                        iRowNetSales = arrRow[2];
                        xlsSheet.Cells[4, 1] = InvoiceOBJ.Factory + "-FACTORY";
                        rsSum = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, true, false); //trading          
                        rsSumReturn = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, false, false); //trading

                        xlsSheet.Range["A" + arrRow[0]].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + arrRow[1]].CopyFromRecordset(rsSumReturn);

                        Sset = 10;
                        Ppcs = 11;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[5, indexMonth] = drr[0];

                            xlsSheet.Range[xlsSheet.Cells[7, Sset], xlsSheet.Cells[21, Sset]].FormulaR1C1 =@"=IFERROR(R[0]C[-1]/R[0]C[-7],""-"")";
                           // xlsSheet.Range[xlsSheet.Cells[26, Sset], xlsSheet.Cells[39, Sset]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-7],0)";
                           // xlsSheet.Range[xlsSheet.Cells[46, Sset], xlsSheet.Cells[59, Sset]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-7],0)";
                            Sset = Sset + 9;

                            xlsSheet.Range[xlsSheet.Cells[7, Ppcs], xlsSheet.Cells[21, Ppcs]].FormulaR1C1 = @"=IFERROR(R[0]C[-2]/R[0]C[-7],""-"")";
                           // xlsSheet.Range[xlsSheet.Cells[26, Ppcs], xlsSheet.Cells[39, Ppcs]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-7],0)";
                           // xlsSheet.Range[xlsSheet.Cells[46, Ppcs], xlsSheet.Cells[59, Ppcs]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-7],0)";
                            Ppcs = Ppcs + 9;

                            indexMonth += 9;

                        }
                        xlsSheet = xlsBook.Sheets[1];
                        xlsSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                        xlsSheet = xlsBook.Sheets[2];
                        xlsSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

                    }
                    else if(InvoiceOBJ.Factory=="GMO")
                    {
                        indexMonth = 3;

                        arrRow = new[] { 7, 12, 17 };
                       // arrHead = new[] { 4,24,44};
                        rsSum = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, true, false); //trading          
                        rsSumReturn = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, false, false); //trading
                        xlsSheet.Range["A" + arrRow[0]].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + arrRow[1]].CopyFromRecordset(rsSumReturn);


                        rsSum = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, true, true); //trading          
                        rsSumReturn = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, false, true); //trading
                        xlsSheet.Range["A" +28].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + 33].CopyFromRecordset(rsSumReturn);

                        Sset = 10;
                        Ppcs = 11;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[5, indexMonth] = drr[0];
                            xlsSheet.Cells[26, indexMonth] = drr[0];
                            xlsSheet.Cells[47, indexMonth] = drr[0];
                            indexMonth += 9;


                            xlsSheet.Range[xlsSheet.Cells[7, Sset], xlsSheet.Cells[21, Sset]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-7],""-"")";   //gmo
                            xlsSheet.Range[xlsSheet.Cells[28, Sset], xlsSheet.Cells[42, Sset]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-7],""-"")";  //trading
                            //xlsSheet.Range[xlsSheet.Cells[46, Sset], xlsSheet.Cells[59, Sset]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-7],0)";   
                            Sset = Sset + 9;

                            xlsSheet.Range[xlsSheet.Cells[7, Ppcs], xlsSheet.Cells[21, Ppcs]].FormulaR1C1 = @"=IFERROR(R[0]C[-2]/R[0]C[-7],""-"")";   //gmo
                             xlsSheet.Range[xlsSheet.Cells[28, Ppcs], xlsSheet.Cells[42, Ppcs]].FormulaR1C1 = @"=IFERROR(R[0]C[-2]/R[0]C[-7],""-"")";   //trading
                            // xlsSheet.Range[xlsSheet.Cells[46, Ppcs], xlsSheet.Cells[59, Ppcs]].FormulaR1C1 = "=IFERROR(R[0]C[-2]/R[0]C[-7],0)";
                            Ppcs = Ppcs + 9;


                        }

                        xlsSheet = xlsBook.Sheets[1];
                        xlsSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                        xlsSheet = xlsBook.Sheets[3];
                        xlsSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

                    }
                    else if (InvoiceOBJ.Factory == "ALL")
                    {
                        indexMonth = 3;

                        arrRow = new[] { 7, 12, 17 };

                        //GMO
                        InvoiceOBJ.Factory = "GMO";
                        rsSum = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, true, false); //trading          
                        rsSumReturn = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, false, false); //trading
                        xlsSheet.Range["A" + arrRow[0]].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + arrRow[1]].CopyFromRecordset(rsSumReturn);

                        //GMO Trading
                        rsSum = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, true, true); //trading          
                        rsSumReturn = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, false, true); //trading
                        xlsSheet.Range["A" + 28].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + 33].CopyFromRecordset(rsSumReturn);

                        //PO FAC
                        InvoiceOBJ.Factory = "PO";
                        rsSum = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, true, false); //PO          
                        rsSumReturn = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, false, false); //return
                        xlsSheet.Range["A" + 70].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + 75].CopyFromRecordset(rsSumReturn);

                        //RP FAC
                        InvoiceOBJ.Factory = "RP";
                        rsSum = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, true, false); //RP         
                        rsSumReturn = InvoiceDAL.getSaleByCustomerAndCurrency(InvoiceOBJ, false, false); //return
                        xlsSheet.Range["A" + 91].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + 96].CopyFromRecordset(rsSumReturn);


                        Sset = 10;
                        Ppcs = 11;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[5, indexMonth] = drr[0];
                            xlsSheet.Cells[26, indexMonth] = drr[0];
                            xlsSheet.Cells[47, indexMonth] = drr[0];
                            xlsSheet.Cells[68, indexMonth] = drr[0];
                            xlsSheet.Cells[89, indexMonth] = drr[0];
                            indexMonth += 9;

                            xlsSheet.Range[xlsSheet.Cells[7, Sset], xlsSheet.Cells[21, Sset]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-7],""-"")";   //gmo
                            xlsSheet.Range[xlsSheet.Cells[28, Sset], xlsSheet.Cells[42, Sset]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-7],""-"")";  //trading
                            xlsSheet.Range[xlsSheet.Cells[70, Sset], xlsSheet.Cells[84, Sset]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-7],""-"")";   //PO
                            xlsSheet.Range[xlsSheet.Cells[91, Sset], xlsSheet.Cells[105, Sset]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-7],""-"")";   //RP
                            Sset = Sset + 9;

                            xlsSheet.Range[xlsSheet.Cells[7, Ppcs], xlsSheet.Cells[21, Ppcs]].FormulaR1C1 = @"=IFERROR(R[0]C[-2]/R[0]C[-7],""-"")";   //gmo
                            xlsSheet.Range[xlsSheet.Cells[28, Ppcs], xlsSheet.Cells[42, Ppcs]].FormulaR1C1 = @"=IFERROR(R[0]C[-2]/R[0]C[-7],""-"")";   //trading
                            xlsSheet.Range[xlsSheet.Cells[70, Ppcs], xlsSheet.Cells[84, Ppcs]].FormulaR1C1 = @"=IFERROR(R[0]C[-2]/R[0]C[-7],""-"")"; //PO
                            xlsSheet.Range[xlsSheet.Cells[91, Ppcs], xlsSheet.Cells[105, Ppcs]].FormulaR1C1 = @"=IFERROR(R[0]C[-2]/R[0]C[-7],""-"")"; //RP

                            Ppcs = Ppcs + 9;

                        }

                        xlsSheet = xlsBook.Sheets[3];
                        xlsSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                        xlsSheet = xlsBook.Sheets[2];
                        xlsSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                    }





                    xlsSheet.Range["B:ZZ"].EntireColumn.AutoFit();
             

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Sale By Currency

        public string getSaleByCustomerAndCurrency(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum, rsSumReturn = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;
                DataRow dr;
                DataTable dt = new DataTable();
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < InvoiceOBJ.DateTo);

                //rsSum = InvoiceDAL.getDetailbyByGroupCode(InvoiceOBJ, true); //External


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
                int indexMonth = 5;
                Excel.Workbook xlsBookTemplate;
                bool trading = false;


                xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleByCustomerAndCurrency\SalesByCustomerAndCurrency.xls");


                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                switch (InvoiceOBJ.Factory)
                {
                    case "GMO":
                        xlsSheet = xlsBook.Sheets[2];
                        trading = true;
                        break;
                    case "ALL":
                        xlsBookTemplate.Close();
                        break;
                    default:
                        xlsSheet = xlsBook.Sheets[1];
                        break;
                }


                xlsSheet.Name = InvoiceOBJ.Factory + "-FACTORY";
                xlsSheet.Cells[4, 1] = InvoiceOBJ.Factory + "-FACTORY";
                xlsSheet.Cells[3, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));

                int[] arrRow = { 7, 11, 17,21 };
                int[] arrHead = {7, 12, 21 };
                int iRowNetSales = arrRow[0];
                int iRowSale = arrRow[0];
                int Sset, Ppcs;
                int intStartRow = 7;
                int StratInteRow = 36;
                System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();



                if (dtMonthRange.Rows.Count > 1)
                {
                    //Column
                    xlsSheet.Cells[5, 23] = "TOTAL " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];

                    //xlsSheet.Cells[5, 23] = "TOTAL " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1]);




                    for (int i = 0; i < dtMonthRange.Rows.Count - 1; i++)
                    {
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[5, 5], xlsSheet.Cells[105, 13]];
                        rangeSource.EntireColumn.Copy();
                        rangeDest = xlsSheet.Cells[5, 5];
                        rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                    }

                    xlsSheet.Range[xlsSheet.Cells[5, 5], xlsSheet.Cells[105, 13]].EntireColumn.delete();
                }
                else
                {
                    xlsSheet.Cells[5, 22] = "TOTAL - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];

                    xlsSheet.Range[xlsSheet.Cells[5, 14], xlsSheet.Cells[105, 22]].EntireColumn.delete();
                    //Column = 20;
                }


                   // Start External
                    rsSum = InvoiceDAL.getCustomer(InvoiceOBJ, false, true);  //external //no trading     
                    adapter.Fill(dt, rsSum);


//========================================================== Swipe Customer ===================================//
                    if (rsSum.RecordCount > 0)
                    {
                        if (InvoiceOBJ.Factory == "GMO")
                        {
                            DataTable dtTemp = new DataTable();
                            dtTemp = dt.Clone();

                            DataRow[] result = dt.Select("INVOICEACCOUNT = 'AREX016'");
                            dtTemp.Rows.Add(result[0].ItemArray);


                            for (int i = 0; i < dt.Rows.Count - 1; i++)
                            {
                                DataRow drr = dtTemp.NewRow();
                                drr.ItemArray = dt.Rows[i].ItemArray;
                                dtTemp.Rows.Add(drr);
                            }

                            dt.Clear();
                            dt = dtTemp.Copy();
                        }
                    }

//===================================================================================================================//


                    int j = 7;
                    int k = 1;

                    //Row
                    if (rsSum.RecordCount > 0)
                    {

                        for (int i = 0; i < rsSum.RecordCount - 1; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[(intStartRow + 6), (4 + dtMonthRange.Rows.Count * 16)]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (7), 1], xlsSheet.Cells[(intStartRow) + (7) + 4, (4 + dtMonthRange.Rows.Count * 16)]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            intStartRow = intStartRow + (7);
                        }

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 5) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();

                        
                        for (int i = 1; i <= rsSum.RecordCount; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[j, 1], xlsSheet.Cells[(j + 6), 1]].Merge();
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[j, 2], xlsSheet.Cells[(j + 6), 2]].Merge();


                             xlsSheet.Cells[1][j] = dt.Rows[i - 1][0];
                             xlsSheet.Cells[2][j] = dt.Rows[i - 1][1];
                          
                            

                            j = j + 7;
                            k = k + 1;
                        }

                        //adapter.Fill(dt, rsSum);
                        intStartRow = 7;
                        ///////////////////////////////////////////////////////////////////////////// Swipe
                      
                        /*
                        rsSum = InvoiceDAL.getSaleByCustomerAndCustCode(dtMonthRange, InvoiceOBJ, result[0][1].ToString(), true);

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, 3]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + 1, 3], xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, 3]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["C" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                        intStartRow = intStartRow + (5) + rsSum.RecordCount;
                        */
                        for (int y = 0; y < dt.Rows.Count; y++)
                        {
                                rsSum = InvoiceDAL.getSaleByCustomerAndCustCode(dtMonthRange, InvoiceOBJ, dt.Rows[y][1].ToString(), true);

                                rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, 3]];
                                rangeSource.EntireRow.Copy();
                                rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + 1, 3], xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, 3]];
                                rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                xlsSheet.Range["C" + intStartRow].CopyFromRecordset(rsSum);
                                xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                                intStartRow = intStartRow + (5) + rsSum.RecordCount;
                            
                        }

                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 5) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                    }// end rsSum
                    // END EXTERNAL

//=====================================================================================================================================================================================================//


                    //Start Internal
                    rsSum = InvoiceDAL.getCustomer(InvoiceOBJ, false, false);  //internal //no trading     
                    dt.Clear();
                    adapter.Fill(dt, rsSum);


                    if (rsSum.RecordCount > 0)
                    {
                        //========================================================== Swipe Customer ===================================//
                        if (InvoiceOBJ.Factory == "GMO")
                        {
                            DataTable dtTemp = new DataTable();
                            dtTemp = dt.Clone();

                            DataRow[] result = dt.Select("INVOICEACCOUNT = 'AREX016'");
                            dtTemp.Rows.Add(result[0].ItemArray);


                            for (int i = 0; i < dt.Rows.Count - 1; i++)
                            {
                                DataRow drr = dtTemp.NewRow();
                                drr.ItemArray = dt.Rows[i].ItemArray;
                                dtTemp.Rows.Add(drr);
                            }

                            dt.Clear();
                            dt = dtTemp.Copy();
                        }
                    }
                    //===================================================================================================================//



                    StratInteRow = intStartRow + 15;
                    intStartRow = intStartRow + 15;

                    //Row
                    if (rsSum.RecordCount > 0)
                    {
                        for (int i = 0; i < rsSum.RecordCount - 1; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[(intStartRow + 6), (4 + dtMonthRange.Rows.Count * 16)]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (7), 1], xlsSheet.Cells[(intStartRow) + (7) + 4, (4 + dtMonthRange.Rows.Count * 16)]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            intStartRow = intStartRow + (7);
                        }

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 5) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();


                        j = StratInteRow;
                        k = 1;
                        for (int i = 1; i <= rsSum.RecordCount; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[j, 1], xlsSheet.Cells[(j + 6), 1]].Merge();
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[j, 2], xlsSheet.Cells[(j + 6), 2]].Merge();
                            xlsSheet.Cells[1][j] = dt.Rows[i - 1][0];
                            xlsSheet.Cells[2][j] = dt.Rows[i - 1][1];

                            j = j + 7;
                            k = k + 1;
                        }

                        // ==================== Check GMO NP1 ====================//
                        Excel.Range findRang;
                        StringBuilder find = new StringBuilder();
                        findRang = xlsSheet.Range["A:A"].Find(What: "INTERNAL SALE (GMO (NP1))", LookIn: Excel.XlFindLookIn.xlFormulas,
                 LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            find.Append(dt.Rows[i][1].ToString()+",");
                        }

                            if (findRang != null)
                            {
                                xlsSheet.Range[xlsSheet.Cells[(findRang.Row), 1], xlsSheet.Cells[(findRang.Row + 6), (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                                DataRow[] rows;

                                rows = dt.Select("NUMBERSEQUENCEGROUP2='INTERNAL SALE (GMO (NP1))'");
                                foreach (DataRow row in rows)
                                {
                                    dt.Rows.Remove(row);

                                }

                            }




                        //adapter.Fill(dt, rsSum);
                        intStartRow = StratInteRow;
                        for (int y = 0; y < dt.Rows.Count; y++)
                        {
                            if (findRang!=null)
                            {
                                rsSum = InvoiceDAL.getSaleByCustomerAndCustCode2(InvoiceOBJ, find.ToString(), false);
                            }
                            else
                            {
                                rsSum = InvoiceDAL.getSaleByCustomerAndCustCode(dtMonthRange,InvoiceOBJ, dt.Rows[y][1].ToString(), false);
                            }
                          

                                rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, 3]];
                                rangeSource.EntireRow.Copy();
                                rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + 1, 3], xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, 3]];
                                rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                xlsSheet.Range["C" + intStartRow].CopyFromRecordset(rsSum);
                                xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                                intStartRow = intStartRow + (5) + rsSum.RecordCount;
                            
                        }

                        intStartRow = intStartRow + 29;
                    }
                    else
                    {

                        //xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 42) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                        //intStartRow = intStartRow - 30;

                    }// end rsSum

                
                    Sset = 12;
                    Ppcs = 13;
                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[5, indexMonth] = drr[0];

                        xlsSheet.Range[xlsSheet.Cells[7, Sset], xlsSheet.Cells[intStartRow -1 , Sset]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-7],""-"")";
                        Sset = Sset + 9;
                        xlsSheet.Range[xlsSheet.Cells[7, Ppcs], xlsSheet.Cells[intStartRow-1 , Ppcs]].FormulaR1C1 = @"=IFERROR(R[0]C[-2]/R[0]C[-7],""-"")";
                        Ppcs = Ppcs + 9;

                        indexMonth += 9;

                    }

 //=======================================================================================================================================================================================================//      

                    if (trading)
                    {
                        // Start External
                        intStartRow = intStartRow + 6;
                        StratInteRow = intStartRow;

                        dt.Clear();
                        rsSum = InvoiceDAL.getCustomer(InvoiceOBJ, true, true);  //external // trading     
                        adapter.Fill(dt, rsSum);


                        //========================================================== Swipe Customer ===================================//
                        if (rsSum.RecordCount > 0)
                        {
                            if (InvoiceOBJ.Factory == "GMO")
                            {
                                DataTable dtTemp = new DataTable();
                                dtTemp = dt.Clone();

                                DataRow[] result = dt.Select("INVOICEACCOUNT = 'AREX016'");
                                dtTemp.Rows.Add(result[0].ItemArray);


                                for (int i = 0; i < dt.Rows.Count - 1; i++)
                                {
                                    DataRow drr = dtTemp.NewRow();
                                    drr.ItemArray = dt.Rows[i].ItemArray;
                                    dtTemp.Rows.Add(drr);
                                }

                                dt.Clear();
                                dt = dtTemp.Copy();
                            }
                        }
                        //===================================================================================================================//


                        //Row
                        if (rsSum.RecordCount > 1)
                        {
                            for (int i = 0; i < rsSum.RecordCount - 1; i++)
                            {
                                rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[(intStartRow + 6), (4 + dtMonthRange.Rows.Count * 16)]];
                                rangeSource.EntireRow.Copy();
                                rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (7), 1], xlsSheet.Cells[(intStartRow) + (7) + 4, (4 + dtMonthRange.Rows.Count * 16)]];
                                rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                intStartRow = intStartRow + (7);
                            }

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 5) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                            //intStartRow = intStartRow - 30; //
                        }
                        /*
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 5) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                    }
                    */

                        j = StratInteRow;
                        k = 1;
                        for (int i = 1; i <= rsSum.RecordCount; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[j, 1], xlsSheet.Cells[(j + 6), 1]].Merge();
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[j, 2], xlsSheet.Cells[(j + 6), 2]].Merge();
                            xlsSheet.Cells[1][j] = dt.Rows[i - 1][0];
                            xlsSheet.Cells[2][j] = dt.Rows[i - 1][1];

                            j = j + 7;
                            k = k + 1;
                        }

                        //adapter.Fill(dt, rsSum);
                        intStartRow = StratInteRow;
                        for (int y = 0; y < dt.Rows.Count; y++)
                        {
                            rsSum = InvoiceDAL.getSaleByCustomerAndCustCodeTrading(dtMonthRange,InvoiceOBJ, true, dt.Rows[y][1].ToString()); //trading

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, 3]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + 1, 3], xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, 3]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["C" + intStartRow].CopyFromRecordset(rsSum);
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                            intStartRow = intStartRow + (5) + rsSum.RecordCount;


                        }
                        // END EXTERNAL

                        /*//Start Internal
                       rsSum = InvoiceDAL.getCustomer(InvoiceOBJ, true, false);  //internal // trading     
                       dt.Clear();
                       adapter.Fill(dt, rsSum);
                       StratInteRow = intStartRow + 15;
                       intStartRow = intStartRow + 15;

                       //Row
                       if (rsSum.RecordCount > 1)
                       {
                           for (int i = 0; i < rsSum.RecordCount - 1; i++)
                           {
                               rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[(intStartRow + 6), (4 + dtMonthRange.Rows.Count * 16)]];
                               rangeSource.EntireRow.Copy();
                               rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (7), 1], xlsSheet.Cells[(intStartRow) + (7) + 4, (4 + dtMonthRange.Rows.Count * 16)]];
                               rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                               intStartRow = intStartRow + (7);
                           }

                           xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 5) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                       }
                       else
                       {

                           xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 5) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();


                       }

                       j = StratInteRow;
                       k = 1;
                       for (int i = 1; i <= rsSum.RecordCount; i++)
                       {
                           rangeSource = xlsSheet.Range[xlsSheet.Cells[j, 1], xlsSheet.Cells[(j + 6), 1]].Merge();
                           rangeSource = xlsSheet.Range[xlsSheet.Cells[j, 2], xlsSheet.Cells[(j + 6), 2]].Merge();
                           xlsSheet.Cells[1][j] = dt.Rows[i - 1][0];
                           xlsSheet.Cells[2][j] = dt.Rows[i - 1][1];

                           j = j + 7;
                           k = k + 1;
                       }

                       //adapter.Fill(dt, rsSum);
                       intStartRow = StratInteRow;
                       for (int y = 0; y < dt.Rows.Count; y++)
                       {
                           rsSum = InvoiceDAL.getSaleByCustomerAndCustCode(InvoiceOBJ, true, dt.Rows[y][1].ToString(), false); //trading

                           rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, 3]];
                           rangeSource.EntireRow.Copy();
                           rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + 1, 3], xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, 3]];
                           rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                           xlsSheet.Range["C" + intStartRow].CopyFromRecordset(rsSum);
                           xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount) + 1, (4 + dtMonthRange.Rows.Count * 16)]].EntireRow.delete();
                           intStartRow = intStartRow + (5) + rsSum.RecordCount;


                       }
                       // END INTERNAL
                       */

                        Sset = 12;
                        Ppcs = 13;

                        foreach (DataRow drr in dtMonthRange.Rows)
                        {

                            xlsSheet.Range[xlsSheet.Cells[StratInteRow, Sset], xlsSheet.Cells[intStartRow + 14, Sset]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-7],""-"")";
                            Sset = Sset + 9;
                            xlsSheet.Range[xlsSheet.Cells[StratInteRow, Ppcs], xlsSheet.Cells[intStartRow + 14, Ppcs]].FormulaR1C1 = @"=IFERROR(R[0]C[-2]/R[0]C[-7],""-"")";
                            Ppcs = Ppcs + 9;

                        }
                        xlsSheet = xlsBook.Sheets[1];
                        xlsSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;


                    }
                    else
                    {
                        xlsSheet = xlsBook.Sheets[2];
                        xlsSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                   } //end Trading


                  

                xlsSheet.Range["C:AZ"].EntireColumn.AutoFit();


                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Sale By CurrencyByCustomer


        public string getInvoiceDetail(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;
    
              

                rsSum = InvoiceDAL.getInvoiceDetail(InvoiceOBJ); //External

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
                      xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\InvoiceDetail\InvoiceDetail.xlsx");

                   

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Invoice Detail";
                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));

                    xlsSheet.Cells[3, 9] = String.Format("{0:dd-MMM-yyyy} to {1:dd-MMM-yyyy}", InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo);
                   

                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 14]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 14]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 13]].EntireRow.delete();

                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 14]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";

                    xlsSheet.Range["A:N"].EntireColumn.AutoFit();

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


        public string getInvoiceByDate(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dtDateForm = InvoiceOBJ.DateFrom;
                DateTime dateRunning = dtDateForm;
                DateTime dtDateTo = InvoiceOBJ.DateTo;
                

                   string strSystemPath = System.IO.Directory.GetCurrentDirectory();


                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    

                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    int intStartRow = 5; //StartRow
                    int intEndRow;
                    int tmpStartRow;
                    Excel.Workbook xlsBookTemplate;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\InvoiceByDate\InvoiceByDate.xls");



                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();



                    if (!InvoiceOBJ.ShowWH)
                    {

                        xlsSheet = xlsBook.Sheets[1];
                        xlsSheet.Name = "InvoiceReportByDate";
                        xlsSheet.Cells[1, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                        xlsSheet.Cells[2, 1] = String.Format("{0:dd-MMM-yyyy} to {1:dd-MMM-yyyy}", InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo);

                        while (dateRunning <= dtDateTo)
                        {

                            rsSum = InvoiceDAL.getInvoiceByDate(InvoiceOBJ, dateRunning,""); //External

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                                //xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 14]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";
                                intEndRow = intStartRow + rsSum.RecordCount - 1;
                                tmpStartRow = drawTotalByCurrency(xlsSheet, intStartRow, intEndRow);
                                xlsSheet.Range["A" + intStartRow, "Q" + tmpStartRow].Borders.LineStyle = 1;
                                xlsSheet.Range["A" + tmpStartRow, "Q" + tmpStartRow].Interior.Color = 15261367;
                                intStartRow = tmpStartRow + 2;

                            } //end rsSum.RecordCount

                            rsSum.Close();
                            dateRunning = dateRunning.AddDays(1);
                        }

                        tmpStartRow = drawTotalByCurrency(xlsSheet, 3, intStartRow);
                        xlsSheet.Range["A" + (intStartRow + 1), "Q" + tmpStartRow].Borders.LineStyle = 1;
                        xlsSheet.Range["A" + tmpStartRow, "Q" + tmpStartRow].Interior.Color = 15261367;

                        xlsSheet.Range["B:Q"].EntireColumn.AutoFit();


                    }
                   
                       
                        //============================== F1 ===================================//
                        intStartRow = 5;
                        dateRunning = dtDateForm;
                        xlsSheet = xlsBook.Sheets[2];
                        xlsSheet.Name = "F1";
                        xlsSheet.Cells[1, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                        xlsSheet.Cells[2, 1] = String.Format("{0:dd-MMM-yyyy} to {1:dd-MMM-yyyy}", InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo);

                        while (dateRunning <= dtDateTo)
                        {

                            rsSum = InvoiceDAL.getInvoiceByDate(InvoiceOBJ, dateRunning,"F1"); //External

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                                //xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 14]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";
                                intEndRow = intStartRow + rsSum.RecordCount - 1;
                                tmpStartRow = drawTotalByCurrency(xlsSheet, intStartRow, intEndRow);
                                xlsSheet.Range["A" + intStartRow, "Q" + tmpStartRow].Borders.LineStyle = 1;
                                xlsSheet.Range["A" + tmpStartRow, "Q" + tmpStartRow].Interior.Color = 15261367;
                                intStartRow = tmpStartRow + 2;

                            } //end rsSum.RecordCount

                            rsSum.Close();
                            dateRunning = dateRunning.AddDays(1);
                        }

                        tmpStartRow = drawTotalByCurrency(xlsSheet, 3, intStartRow);
                        xlsSheet.Range["A" + (intStartRow + 1), "Q" + tmpStartRow].Borders.LineStyle = 1;
                        xlsSheet.Range["A" + tmpStartRow, "Q" + tmpStartRow].Interior.Color = 15261367;

                        xlsSheet.Range["B:Q"].EntireColumn.AutoFit();


                        //========================== F2 ================================//
                        intStartRow = 5;
                        xlsSheet = xlsBook.Sheets[3];
                        xlsSheet.Name = "F2";
                        dateRunning = dtDateForm;
                        xlsSheet.Cells[1, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                        xlsSheet.Cells[2, 1] = String.Format("{0:dd-MMM-yyyy} to {1:dd-MMM-yyyy}", InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo);

                        while (dateRunning <= dtDateTo)
                        {

                            rsSum = InvoiceDAL.getInvoiceByDate(InvoiceOBJ, dateRunning, "F2"); //External

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                                //xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 14]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";
                                intEndRow = intStartRow + rsSum.RecordCount - 1;
                                tmpStartRow = drawTotalByCurrency(xlsSheet, intStartRow, intEndRow);
                                xlsSheet.Range["A" + intStartRow, "Q" + tmpStartRow].Borders.LineStyle = 1;
                                xlsSheet.Range["A" + tmpStartRow, "Q" + tmpStartRow].Interior.Color = 15261367;
                                intStartRow = tmpStartRow + 2;

                            } //end rsSum.RecordCount

                            rsSum.Close();
                            dateRunning = dateRunning.AddDays(1);
                        }

                        tmpStartRow = drawTotalByCurrency(xlsSheet, 3, intStartRow);
                        xlsSheet.Range["A" + (intStartRow + 1), "Q" + tmpStartRow].Borders.LineStyle = 1;
                        xlsSheet.Range["A" + tmpStartRow, "Q" + tmpStartRow].Interior.Color = 15261367;

                        xlsSheet.Range["B:Q"].EntireColumn.AutoFit();



                        if (InvoiceOBJ.ShowWH)
                        {
                            xlsBook.Sheets[1].delete();
                        }
                     
                    



                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


             

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Invoice By Date

        static int drawTotalByCurrency(Excel.Worksheet xlsSheet,int intStartRow,int intEndRow){

            string[] arr = { "THB", "JPY", "USD", "CNY" };
            int intRow = intEndRow + 1;
            foreach (string strCurr in arr)
            {
                xlsSheet.Cells[intRow, 6] = "Total";
                xlsSheet.Cells[intRow, 7] = strCurr;
                xlsSheet.Range["H" + intRow + ":Q" + intRow].Formula = "=SUMIF($F$" + intStartRow + ":$F$" + intEndRow + ",$G$" + intRow + ",H$" + intStartRow + ":H$" + intEndRow + ")";

                intRow += 1;
            }

              xlsSheet.Cells[intRow, 6] = "Grand Total";
              xlsSheet.Range["H" + intRow + ":Q" + intRow].Formula = "=SUM(H" + (intEndRow + 1) + ":H" + (intEndRow + arr.Length) + ")";
        
              xlsSheet.Cells[intRow, 11] = "";
              xlsSheet.Cells[intRow, 16] = "";

            return intRow;
        }

        static DataTable Pivot(DataTable tbl)
        {


            var tblPivot = new DataTable();
            tblPivot.Columns.Add(tbl.Columns[0].ColumnName);
            int count = 0;
            for (int i = 0; i < tbl.Rows.Count; i++)
            {
                for (int y = 0; y < 3; y++)
                {
                    var r = tblPivot.NewRow();

                    r[count] = tbl.Rows[i][y];


                    tblPivot.Rows.Add(r);

                }
            }

            return tblPivot;
          
        }//end Pivot


        public string getInvoiceByCustomer(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;

                string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                Excel.Application xlsApp = new Excel.Application();
                System.Globalization.CultureInfo oldCI;
                oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");



                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                int intStartRow = 5;  //StartRow
                Excel.Workbook xlsBookTemplate;
                xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\InvoiceByCustomer\InvoiceReport.xls");



                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();



                if (!InvoiceOBJ.ShowWH)
                {

                    rsSum = InvoiceDAL.getInvoiceByCustomer(InvoiceOBJ, ""); //External

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet = xlsBook.Sheets[1];
                        xlsSheet.Name = "InvoiceReportByCustomer";
                        xlsSheet.Cells[1, 1] = "InvoiceReportByCustomer";

                        xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "Q" + (intStartRow + rsSum.RecordCount - 1)].Borders.LineStyle = 1;

                        xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=$D1=" + Convert.ToChar(34) + "GRAND TOTAL" + Convert.ToChar(34));
                        xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[1].Interior.Color = 14281213;


                        xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=AND($D1<>"""",$F1="""")");
                        xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[2].Interior.Color = 15986394;


                        xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=AND($D1<>"""",$E1="""",$F1<>"""")");
                        xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[3].Interior.Color = 14408946;


                        xlsSheet.Range["D:Q"].EntireColumn.AutoFit();



                    }//end rsSum.RecordCount



                }
               
                //===================================================================================//
                rsSum = InvoiceDAL.getInvoiceByCustomer(InvoiceOBJ, "F1"); //External

                if (rsSum.RecordCount > 0)
                {


                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Name = "F1";
                    xlsSheet.Cells[1, 1] = "InvoiceReportByCustomer";

                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "Q" + (intStartRow + rsSum.RecordCount - 1)].Borders.LineStyle = 1;

                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=$D1=" + Convert.ToChar(34) + "GRAND TOTAL" + Convert.ToChar(34));
                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[1].Interior.Color = 14281213;


                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=AND($D1<>"""",$F1="""")");
                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[2].Interior.Color = 15986394;


                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=AND($D1<>"""",$E1="""",$F1<>"""")");
                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[3].Interior.Color = 14408946;


                    xlsSheet.Range["D:Q"].EntireColumn.AutoFit();


                }

                //==============================================================================//
                rsSum = InvoiceDAL.getInvoiceByCustomer(InvoiceOBJ, "F2"); //External

                if (rsSum.RecordCount > 0)
                {


                    xlsSheet = xlsBook.Sheets[3];
                    xlsSheet.Name = "F2";
                    xlsSheet.Cells[1, 1] = "InvoiceReportByCustomer";

                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "Q" + (intStartRow + rsSum.RecordCount - 1)].Borders.LineStyle = 1;

                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=$D1=" + Convert.ToChar(34) + "GRAND TOTAL" + Convert.ToChar(34));
                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[1].Interior.Color = 14281213;


                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=AND($D1<>"""",$F1="""")");
                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[2].Interior.Color = 15986394;


                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=AND($D1<>"""",$E1="""",$F1<>"""")");
                    xlsSheet.Range["A1", "Q" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[3].Interior.Color = 14408946;


                    xlsSheet.Range["D:Q"].EntireColumn.AutoFit();




                }


                if (InvoiceOBJ.ShowWH)
                {
                    xlsBook.Sheets[1].delete();

                }

                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;



                return null;

            }





            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Invoice by Customer

        public string getInvoiceByItem(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;

                string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                Excel.Application xlsApp = new Excel.Application();
                System.Globalization.CultureInfo oldCI;
                oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");



                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                int intStartRow = 4;  //StartRow
                cExcel cExcel =  new cExcel();
                Excel.Workbook xlsBookTemplate;
                xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\InvoiceReportByItem\InvoiceReportByItem.xls");



                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();



                if (!InvoiceOBJ.ShowWH)
                {

                    rsSum = InvoiceDAL.getInvoiceByItem(InvoiceOBJ, true,""); //com

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet = xlsBook.Sheets[1];
                        xlsSheet.Name = "InvoiceReportByItem";
                        xlsSheet.Cells[1, 1] = "Invoice Report By Item";

                        xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "S" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                        intStartRow += rsSum.RecordCount;


                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""GRAND TOTAL*"",$B1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[1].Interior.Color = 10192433;



                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing,@"=SEARCH(""TOTAL"",$E1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[2].Interior.Color = 65535;


                        xlsSheet.Range[xlsSheet.Cells[intStartRow - 1, 1], xlsSheet.Cells[intStartRow - 1, 10]].Merge();
                        xlsSheet.Range[xlsSheet.Cells[intStartRow - 1, 1], xlsSheet.Cells[intStartRow - 1, 10]].EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        
                        //xlsSheet.Range["D:Q"].EntireColumn.AutoFit();

                    }//end rsSum Com

                    rsSum = InvoiceDAL.getInvoiceByItem(InvoiceOBJ, false, ""); //com

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "S" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                       // intStartRow += rsSum.RecordCount;


                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""GRAND TOTAL*"",$B1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[1].Interior.Color = 10192433;



                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""TOTAL"",$E1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[2].Interior.Color = 65535;

                        intStartRow += rsSum.RecordCount;

                        xlsSheet.Cells[intStartRow, 2] = "GRAND TOTAL COM+NOCOM";

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow - 1), 1], xlsSheet.Cells[(intStartRow - 1), 10]].Merge();
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow - 1), 1], xlsSheet.Cells[(intStartRow - 1), 10]].EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                        xlsSheet.Range[xlsSheet.Cells[intStartRow , 1], xlsSheet.Cells[intStartRow, 7]].EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow , 10]].Merge();

                        //xlsSheet.Range["D:Q"].EntireColumn.AutoFit();

                    }//end rsSum No com

                    xlsSheet.Range["K" + intStartRow + ":L" + intStartRow].Formula = @"=SUMIF($E$" + (4) + ":$E$" + (intStartRow - 1) + @",""Total"",K$" + (4) + ":L$" + (intStartRow - 1) + ")";
                    xlsSheet.Range["N" + intStartRow + ":N" + intStartRow].Formula = @"=SUMIF($E$" + (4) + ":$E$" + (intStartRow - 1) + @",""Total"",N$" + (4) + ":N$" + (intStartRow - 1) + ")";
                    xlsSheet.Range["P" + intStartRow + ":R" + intStartRow].Formula = @"=SUMIF($E$" + (4) + ":$E$" + (intStartRow - 1) + @",""Total"",P$" + (4) + ":R$" + (intStartRow - 1) + ")";

                    xlsSheet.Range["K" + intStartRow + ":S" + intStartRow].EntireColumn.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";
                    xlsSheet.Range["O:O"].NumberFormat = "0.00000";
                 


                }

                //===================================================================================//
                rsSum = InvoiceDAL.getInvoiceByItem(InvoiceOBJ,true, "F1"); //External
                intStartRow = 4;

                 if (rsSum.RecordCount > 0)
                    {

                        xlsSheet = xlsBook.Sheets[2];
                        xlsSheet.Name = "F1";
                        xlsSheet.Cells[1, 1] = "Invoice Report By Item";

                        xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "S" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                        intStartRow += rsSum.RecordCount;
                       


                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""GRAND TOTAL*"",$B1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[1].Interior.Color = 10192433;



                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""TOTAL"",$E1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[2].Interior.Color = 65535;


                        xlsSheet.Range[xlsSheet.Cells[intStartRow - 1, 1], xlsSheet.Cells[intStartRow - 1, 10]].Merge();
                        xlsSheet.Range[xlsSheet.Cells[intStartRow - 1, 1], xlsSheet.Cells[intStartRow - 1, 10]].EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        
                        //xlsSheet.Range["D:Q"].EntireColumn.AutoFit();

                    }//end rsSum Com

                    rsSum = InvoiceDAL.getInvoiceByItem(InvoiceOBJ, false, "F1"); //com

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "S" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                       // intStartRow += rsSum.RecordCount;


                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""GRAND TOTAL*"",$B1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[1].Interior.Color = 10192433;



                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""TOTAL"",$E1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[2].Interior.Color = 65535;

                        intStartRow +=  rsSum.RecordCount;

                        xlsSheet.Cells[intStartRow, 2] = "GRAND TOTAL COM+NOCOM";

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow - 1), 1], xlsSheet.Cells[(intStartRow - 1), 10]].Merge();
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow - 1), 1], xlsSheet.Cells[(intStartRow - 1), 10]].EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                        xlsSheet.Range[xlsSheet.Cells[intStartRow , 1], xlsSheet.Cells[intStartRow, 7]].EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow , 10]].Merge();

                        //xlsSheet.Range["D:Q"].EntireColumn.AutoFit();

                    }//end rsSum No com

                    xlsSheet.Range["K"+intStartRow+":L"+intStartRow].Formula = @"=SUMIF($E$" + (4 )+ ":$E$" + (intStartRow - 1) + @",""Total"",K$" + (4) + ":L$" + (intStartRow - 1) + ")";
                    xlsSheet.Range["N" + intStartRow + ":N" + intStartRow].Formula = @"=SUMIF($E$" + (4) + ":$E$" + (intStartRow - 1) + @",""Total"",N$" + (4) + ":N$" + (intStartRow - 1) + ")";
                    xlsSheet.Range["P" + intStartRow + ":R" + intStartRow].Formula = @"=SUMIF($E$" + (4) + ":$E$" + (intStartRow - 1) + @",""Total"",P$" + (4) + ":R$" + (intStartRow - 1) + ")";

                    xlsSheet.Range["K" + intStartRow + ":S" + intStartRow].EntireColumn.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";
                    xlsSheet.Range["O:O"].NumberFormat = "0.00000";
                 

                //============================================ End F1 ============================================//

           
                    rsSum = InvoiceDAL.getInvoiceByItem(InvoiceOBJ, true, "F2"); //External
                    intStartRow = 4;

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet = xlsBook.Sheets[3];
                        xlsSheet.Name = "F2";
                        xlsSheet.Cells[1, 1] = "Invoice Report By Item";

                        xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "S" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                        intStartRow += rsSum.RecordCount;


                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""GRAND TOTAL*"",$B1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[1].Interior.Color = 10192433;



                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""TOTAL"",$E1)");
                        xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[2].Interior.Color = 65535;


                        xlsSheet.Range[xlsSheet.Cells[(intStartRow - 1), 1], xlsSheet.Cells[(intStartRow - 1), 10]].Merge();
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow - 1), 1], xlsSheet.Cells[(intStartRow - 1), 10]].EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        //xlsSheet.Range["D:Q"].EntireColumn.AutoFit();

                    }//end rsSum Com


                    rsSum = InvoiceDAL.getInvoiceByItem(InvoiceOBJ, false, "F2"); //com

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "S" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                        //intStartRow += rsSum.RecordCount;

                        intStartRow += rsSum.RecordCount;

                        xlsSheet.Cells[intStartRow, 2] = "GRAND TOTAL COM+NOCOM";

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow - 1), 1], xlsSheet.Cells[(intStartRow - 1), 10]].Merge();
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow - 1), 1], xlsSheet.Cells[(intStartRow - 1), 10]].EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                        xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 7]].EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 10]].Merge();




                        //xlsSheet.Range["D:Q"].EntireColumn.AutoFit();

                    }//end rsSum No com


                    xlsSheet.Range["K" + intStartRow + ":L" + intStartRow].Formula = @"=SUMIF($E$" + (4) + ":$E$" + (intStartRow - 1) + @",""Total"",K$" + (4) + ":L$" + (intStartRow - 1) + ")";
                    xlsSheet.Range["N" + intStartRow + ":N" + intStartRow].Formula = @"=SUMIF($E$" + (4) + ":$E$" + (intStartRow - 1) + @",""Total"",N$" + (4) + ":N$" + (intStartRow - 1) + ")";
                    xlsSheet.Range["P" + intStartRow + ":R" + intStartRow].Formula = @"=SUMIF($E$" + (4) + ":$E$" + (intStartRow - 1) + @",""Total"",P$" + (4) + ":R$" + (intStartRow - 1) + ")";

                    xlsSheet.Range["K" + intStartRow + ":S" + intStartRow].EntireColumn.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";
                    xlsSheet.Range["O:O"].NumberFormat = "0.00000";

                 
                    //============================================ End F2 ============================================//


                if (InvoiceOBJ.ShowWH)
                {
                    xlsBook.Sheets[1].delete();

                }

                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;



                return null;

            }





            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Invoice by Customer


        public string getInvoiceByInvoice(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;
                cExcel cExcel = new cExcel();
                string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                Excel.Application xlsApp = new Excel.Application();
                System.Globalization.CultureInfo oldCI;
                oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");



                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                int intStartRow = 4;  //StartRow
                Excel.Workbook xlsBookTemplate;
                xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\InvoiceReportByInvoice\InvoiceReportByInvoice.xls");



                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();



                if (!InvoiceOBJ.ShowWH)
                {

                    rsSum = InvoiceDAL.getInvoiceByInvoice(InvoiceOBJ, true,""); //External

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet = xlsBook.Sheets[1];
                        xlsSheet.Name = "InvoiceReportByInvoice";
                        xlsSheet.Cells[1, 1] = "Invoice Report By Invoice";

                        xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "Q" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                        xlsSheet.Cells[(rsSum.RecordCount+intStartRow), 2] = "TOTAL";
                        intStartRow += rsSum.RecordCount;


                        xlsSheet.Range["H" + intStartRow + ":H" + intStartRow].Formula = "=SUM(H" + 4 + ":H" + (intStartRow - 1) + ")";
                        xlsSheet.Range["I" + intStartRow + ":I" + intStartRow].Formula = "=SUM(I" + 4 + ":I" + (intStartRow - 1) + ")";
                        xlsSheet.Range["J" + intStartRow + ":J" + intStartRow].Formula = "=SUM(J" + 4 + ":J" + (intStartRow - 1) + ")";
                        xlsSheet.Range["L" + intStartRow + ":L" + intStartRow].Formula = "=SUM(L" + 4 + ":L" + (intStartRow - 1) + ")";

                        xlsSheet.Range["N" + intStartRow + ":N" + intStartRow].Formula = "=SUM(N" + 4 + ":N" + (intStartRow - 1) + ")";
                        xlsSheet.Range["O" + intStartRow + ":O" + intStartRow].Formula = "=SUM(O" + 4 + ":O" + (intStartRow - 1) + ")";
                        xlsSheet.Range["P" + intStartRow + ":P" + intStartRow].Formula = "=SUM(P" + 4 + ":P" + (intStartRow - 1) + ")";



                    }//end rsSum.RecordCount


                    rsSum = InvoiceDAL.getInvoiceByInvoice(InvoiceOBJ, false, ""); //External

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet.Range["A" + (intStartRow + 1)].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + (intStartRow + 1), "Q" + (intStartRow + rsSum.RecordCount+2)].Borders.LineStyle = 1;
                        xlsSheet.Cells[(rsSum.RecordCount + intStartRow+1), 2] = "TOTAL";
                        intStartRow += rsSum.RecordCount;


                        xlsSheet.Range["H" + (intStartRow + 1) + ":H" + (intStartRow + 1)].Formula = "=SUM(H" + (intStartRow) + ":H" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Range["I" + (intStartRow + 1) + ":I" + (intStartRow + 1)].Formula = "=SUM(I" + (intStartRow) + ":I" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Range["J" + (intStartRow + 1) + ":J" + (intStartRow + 1)].Formula = "=SUM(J" + (intStartRow) + ":J" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Range["L" + (intStartRow + 1) + ":L" + (intStartRow + 1)].Formula = "=SUM(L" + (intStartRow) + ":L" + (intStartRow + rsSum.RecordCount - 1) + ")";

                        xlsSheet.Range["N" + (intStartRow + 1) + ":N" + (intStartRow + 1)].Formula = "=SUM(N" + (intStartRow) + ":N" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Range["O" + (intStartRow + 1) + ":O" + (intStartRow + 1)].Formula = "=SUM(O" + (intStartRow) + ":O" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Range["P" + (intStartRow + 1) + ":P" + (intStartRow + 1)].Formula = "=SUM(P" + (intStartRow) + ":P" + (intStartRow + rsSum.RecordCount - 1) + ")";


                  

                        

                    }//end rsSum.RecordCount

                    xlsSheet.Cells[(intStartRow + 2), 2] = "GRAND TOTAL";

                    xlsSheet.Range["H" + (intStartRow + 2) + ":J" + (intStartRow + 2)].Formula = @"=SUMIF($B$" + 4 + ":$B$" + (intStartRow + 1) + @",""TOTAL"",H$" + 4 + ":J$" + (intStartRow + 1) + ")";
                    xlsSheet.Range["L" + (intStartRow + 2) + ":L" + (intStartRow + 2)].Formula = @"=SUMIF($B$" + 4 + ":$B$" + (intStartRow + 1) + @",""TOTAL"",L$" + 4 + ":L$" + (intStartRow + 1) + ")";
                    xlsSheet.Range["N" + (intStartRow + 2) + ":P" + (intStartRow + 2)].Formula = @"=SUMIF($B$" + 4 + ":$B$" + (intStartRow + 1) + @",""TOTAL"",N$" + 4 + ":P$" + (intStartRow + 1) + ")";

                    xlsSheet.Range["H" + intStartRow + ":Q" + (intStartRow )].EntireColumn.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";

                    xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""GRAND TOTAL*"",$B1)");
                    xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[1].Interior.Color = 10192433;



                    xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""TOTAL*"",$B1)");
                    xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[2].Interior.Color = 65535;
                    xlsSheet.Range["K:K"].NumberFormat = "0.00000";

                    xlsSheet.Range["B:Q"].EntireColumn.AutoFit();
                }


                //=================================F1==================================================//

                rsSum = InvoiceDAL.getInvoiceByInvoice(InvoiceOBJ, true, "F1"); //External
                intStartRow = 4;
                if (rsSum.RecordCount > 0)
                {

                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Name = "F1";
                    xlsSheet.Cells[1, 1] = "Invoice Report By Invoice";

                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "Q" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Cells[(rsSum.RecordCount + intStartRow), 2] = "TOTAL";
                    intStartRow += rsSum.RecordCount;

                    xlsSheet.Range["H" + intStartRow + ":H" + intStartRow].Formula = "=SUM(H" + 4 + ":H" + (intStartRow - 1) + ")";
                    xlsSheet.Range["I" + intStartRow + ":I" + intStartRow].Formula = "=SUM(I" + 4 + ":I" + (intStartRow - 1) + ")";
                    xlsSheet.Range["J" + intStartRow + ":J" + intStartRow].Formula = "=SUM(J" + 4 + ":J" + (intStartRow - 1) + ")";
                    xlsSheet.Range["L" + intStartRow + ":L" + intStartRow].Formula = "=SUM(L" + 4 + ":L" + (intStartRow - 1) + ")";

                    xlsSheet.Range["N" + intStartRow + ":N" + intStartRow].Formula = "=SUM(N" + 4 + ":N" + (intStartRow - 1) + ")";
                    xlsSheet.Range["O" + intStartRow + ":O" + intStartRow].Formula = "=SUM(O" + 4 + ":O" + (intStartRow - 1) + ")";
                    xlsSheet.Range["P" + intStartRow + ":P" + intStartRow].Formula = "=SUM(P" + 4 + ":P" + (intStartRow - 1) + ")";


                }//end rsSum.RecordCount


                rsSum = InvoiceDAL.getInvoiceByInvoice(InvoiceOBJ, false, "F1"); //External

                if (rsSum.RecordCount > 0)
                {

                    xlsSheet.Range["A" + (intStartRow+1)].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + (intStartRow + 1), "Q" + (intStartRow + rsSum.RecordCount+2)].Borders.LineStyle = 1;
                    xlsSheet.Cells[(rsSum.RecordCount + intStartRow+1), 2] = "TOTAL";
                    intStartRow += rsSum.RecordCount;


                        xlsSheet.Range["H" + (intStartRow + 1) + ":H" + (intStartRow + 1)].Formula = "=SUM(H" + (intStartRow) + ":H" + (intStartRow +rsSum.RecordCount-1) + ")";
                        xlsSheet.Range["I" + (intStartRow + 1) + ":I" + (intStartRow + 1)].Formula = "=SUM(I" + (intStartRow) + ":I" + (intStartRow + rsSum.RecordCount-1) + ")";
                        xlsSheet.Range["J" + (intStartRow + 1) + ":J" + (intStartRow + 1)].Formula = "=SUM(J" + (intStartRow) + ":J" + (intStartRow + rsSum.RecordCount-1) + ")";
                        xlsSheet.Range["L" + (intStartRow + 1) + ":L" + (intStartRow + 1)].Formula = "=SUM(L" + (intStartRow) + ":L" + (intStartRow + rsSum.RecordCount-1) + ")";

                        xlsSheet.Range["N" + (intStartRow + 1) + ":N" + (intStartRow + 1)].Formula = "=SUM(N" + (intStartRow) + ":N" + (intStartRow + rsSum.RecordCount-1) + ")";
                        xlsSheet.Range["O" + (intStartRow + 1) + ":O" + (intStartRow + 1)].Formula = "=SUM(O" + (intStartRow) + ":O" + (intStartRow + rsSum.RecordCount-1) + ")";
                        xlsSheet.Range["P" + (intStartRow + 1) + ":P" + (intStartRow + 1)].Formula = "=SUM(P" + (intStartRow) + ":P" + (intStartRow + rsSum.RecordCount-1) + ")";


                  



                }//end rsSum.RecordCount

                xlsSheet.Cells[(intStartRow + 2), 2] = "GRAND TOTAL";
                xlsSheet.Range["H" + (intStartRow +2)+":J" + (intStartRow + 2)].Formula = @"=SUMIF($B$" + 4 + ":$B$" + (intStartRow + 1) + @",""TOTAL"",H$" + 4 + ":J$" + (intStartRow + 1) + ")";
                xlsSheet.Range["L" + (intStartRow +2)+":L" + (intStartRow + 2)].Formula = @"=SUMIF($B$" + 4 + ":$B$" + (intStartRow + 1) + @",""TOTAL"",L$" + 4 + ":L$" + (intStartRow + 1) + ")";
                xlsSheet.Range["N" + (intStartRow +2)+":P" + (intStartRow + 2)].Formula = @"=SUMIF($B$" + 4 + ":$B$" + (intStartRow + 1) + @",""TOTAL"",N$" + 4 + ":P$" + (intStartRow + 1) + ")";

                xlsSheet.Range["H" + intStartRow + ":Q" + (intStartRow)].EntireColumn.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";

                xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""GRAND TOTAL*"",$B1)");
                xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[1].Interior.Color = 10192433;



                xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""TOTAL*"",$B1)");
                xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[2].Interior.Color = 65535;
                xlsSheet.Range["K:K"].NumberFormat = "0.00000";

                xlsSheet.Range["B:Q"].EntireColumn.AutoFit();




                //================================================ F2 ===========================================//
                rsSum = InvoiceDAL.getInvoiceByInvoice(InvoiceOBJ, true, "F2"); //External
                intStartRow = 4;
                if (rsSum.RecordCount > 0)
                {

                    xlsSheet = xlsBook.Sheets[3];
                    xlsSheet.Name = "F2";
                    xlsSheet.Cells[1, 1] = "Invoice Report By Invoice";

                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} {3}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateTo, InvoiceOBJ.CustomerGroup.Replace("','", ", "));
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "Q" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Cells[(rsSum.RecordCount + intStartRow), 2] = "TOTAL";
                    intStartRow += rsSum.RecordCount;


                    xlsSheet.Range["H" + intStartRow + ":H" + intStartRow].Formula = "=SUM(H" + 4 + ":H" + (intStartRow - 1) + ")";
                    xlsSheet.Range["I" + intStartRow + ":I" + intStartRow].Formula = "=SUM(I" + 4 + ":I" + (intStartRow - 1) + ")";
                    xlsSheet.Range["J" + intStartRow + ":J" + intStartRow].Formula = "=SUM(J" + 4 + ":J" + (intStartRow - 1) + ")";
                    xlsSheet.Range["L" + intStartRow + ":L" + intStartRow].Formula = "=SUM(L" + 4 + ":L" + (intStartRow - 1) + ")";

                    xlsSheet.Range["N" + intStartRow + ":N" + intStartRow].Formula = "=SUM(N" + 4 + ":N" + (intStartRow - 1) + ")";
                    xlsSheet.Range["O" + intStartRow + ":O" + intStartRow].Formula = "=SUM(O" + 4 + ":O" + (intStartRow - 1) + ")";
                    xlsSheet.Range["P" + intStartRow + ":P" + intStartRow].Formula = "=SUM(P" + 4 + ":P" + (intStartRow - 1) + ")";



                }//end rsSum.RecordCount


                rsSum = InvoiceDAL.getInvoiceByInvoice(InvoiceOBJ, false, "F2"); //External

                if (rsSum.RecordCount > 0)
                {

                    xlsSheet.Range["A" + (intStartRow + 1)].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + (intStartRow + 1), "Q" + (intStartRow + rsSum.RecordCount+2)].Borders.LineStyle = 1;
                    xlsSheet.Cells[(rsSum.RecordCount + intStartRow+1), 2] = "TOTAL";
                    intStartRow += rsSum.RecordCount;

                    xlsSheet.Range["H" + (intStartRow + 1) + ":H" + (intStartRow + 1)].Formula = "=SUM(H" + (intStartRow) + ":H" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range["I" + (intStartRow + 1) + ":I" + (intStartRow + 1)].Formula = "=SUM(I" + (intStartRow) + ":I" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range["J" + (intStartRow + 1) + ":J" + (intStartRow + 1)].Formula = "=SUM(J" + (intStartRow) + ":J" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range["L" + (intStartRow + 1) + ":L" + (intStartRow + 1)].Formula = "=SUM(L" + (intStartRow) + ":L" + (intStartRow + rsSum.RecordCount - 1) + ")";

                    xlsSheet.Range["N" + (intStartRow + 1) + ":N" + (intStartRow + 1)].Formula = "=SUM(N" + (intStartRow) + ":N" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range["O" + (intStartRow + 1) + ":O" + (intStartRow + 1)].Formula = "=SUM(O" + (intStartRow) + ":O" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range["P" + (intStartRow + 1) + ":P" + (intStartRow + 1)].Formula = "=SUM(P" + (intStartRow) + ":P" + (intStartRow + rsSum.RecordCount - 1) + ")";


                  



                }//end rsSum.RecordCount

                xlsSheet.Cells[(intStartRow + 2), 2] = "GRAND TOTAL";
                xlsSheet.Range["H" + (intStartRow + 2) + ":J" + (intStartRow + 2)].Formula = @"=SUMIF($B$" + 4 + ":$B$" + (intStartRow + 1) + @",""TOTAL"",H$" + 4 + ":J$" + (intStartRow + 1) + ")";
                xlsSheet.Range["L" + (intStartRow + 2) + ":L" + (intStartRow + 2)].Formula = @"=SUMIF($B$" + 4 + ":$B$" + (intStartRow + 1) + @",""TOTAL"",L$" + 4 + ":L$" + (intStartRow + 1) + ")";
                xlsSheet.Range["N" + (intStartRow + 2) + ":P" + (intStartRow + 2)].Formula = @"=SUMIF($B$" + 4 + ":$B$" + (intStartRow + 1) + @",""TOTAL"",N$" + 4 + ":P$" + (intStartRow + 1) + ")";

                xlsSheet.Range["H" + intStartRow + ":Q" + (intStartRow)].EntireColumn.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";

                xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""GRAND TOTAL*"",$B1)");
                xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[1].Interior.Color = 10192433;



                xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, @"=SEARCH(""TOTAL*"",$B1)");
                xlsSheet.Range["$A:$" + cExcel.Num2Col(rsSum.Fields.Count)].FormatConditions[2].Interior.Color = 65535;
                xlsSheet.Range["K:K"].NumberFormat = "0.00000";

                xlsSheet.Range["B:Q"].EntireColumn.AutoFit();




                if (InvoiceOBJ.ShowWH)
                {
                    xlsBook.Sheets[1].delete();

                }

                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;



                return null;

            }





            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Invoice by Invoice

        public string getSaleByCustomer(InvoiceOBJ InvoiceOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = InvoiceOBJ.DateFrom;



                rsSum = InvoiceDAL.getSalesByCustomer(InvoiceOBJ); //External

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
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SalesByCustomer\SalesByCustomer.xls");



                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsSheet = xlsBook.Sheets[1];

                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", InvoiceOBJ.Factory, InvoiceOBJ.DateFrom, InvoiceOBJ.DateFrom);

                   
                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 15]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 15]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                   
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount) - 1, 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 15]].EntireRow.delete();

                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                   // xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 14]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";

                    xlsSheet.Range["A:O"].EntireColumn.AutoFit();

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
   
    }//end class
}
