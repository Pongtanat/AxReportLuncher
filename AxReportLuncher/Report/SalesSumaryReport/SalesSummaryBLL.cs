using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;


namespace NewVersion.Report
{
    class SalesSummaryBLL
    {
        SalesSummaryDAL SalesSummaryDAL = new SalesSummaryDAL();
        SalesSummaryOBJ SalesSummaryOBJ = new SalesSummaryOBJ();
  
            public String getSalesSummary(SalesSummaryOBJ SalesSummaryOBJ)
            {

                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(bool));
                DateTime dateRunning = SalesSummaryOBJ.DateFrom;
                DataRow dr ;
                DataTable dtSumRP = new DataTable();
                DataTable dtSumRSMRS = new DataTable();

                do {
                    dr= dtMonthRange.NewRow();
                    dr["dt"]=dateRunning;
                    dr["YearMonth"]=String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"]=false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning=dateRunning.AddMonths(1);
                } while (dateRunning < SalesSummaryOBJ.DateTo);


                if (dtMonthRange.Rows.Count > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();


                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;

                    Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleSummary\SalesSummary.xls");
                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    ADODB.Recordset rsSum, rsSumRP,RSMRS,SumRSMRS = new ADODB.Recordset();
                    ADODB.Recordset rsSumReturn= new ADODB.Recordset();
                    ADODB.Recordset rsSumRPReturn = new ADODB.Recordset();

                   
                   
                   


                    Excel.Range rangeSource, rangeDest;
                    int[] arrRow = { 16, 12, 6 };
                    int iRowNetSales = arrRow[0];
                    int iRowSale = arrRow[2];
                    int iRowSaleReturn = arrRow[1];
                    Excel.Range findRang, findRangReturn;

                    //===calltes
                    

                    //*************************************
                    if (SalesSummaryOBJ.Factory == "RP")
                    {

                        
                        xlsSheet = xlsBook.Sheets[4];
                        rsSum = SalesSummaryDAL.getSaleSummaryByCustomer2(dtMonthRange, SalesSummaryOBJ, "Total", true);
                        rsSumReturn = SalesSummaryDAL.getSaleSummaryByCustomer2(dtMonthRange, SalesSummaryOBJ, "Total", false);
                        RSMRS = SalesSummaryDAL.getSaleSummaryByCustomerRS_MRS(dtMonthRange, SalesSummaryOBJ, "Total", true, true);
                      
                        rsSumRP = SalesSummaryDAL.getSaleSummaryByCustomerRP(dtMonthRange, SalesSummaryOBJ, "Total", true);
                        rsSumRPReturn = SalesSummaryDAL.getSaleSummaryByCustomerRP(dtMonthRange, SalesSummaryOBJ, "Total", false);
                        SumRSMRS = SalesSummaryDAL.getSaleSummaryByCustomerRS_MRS(dtMonthRange, SalesSummaryOBJ, "Total", false, true);


                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;
                        xlsSheet.Name = "Company - Total (RP)";


                        rangeSource = xlsSheet.Range[xlsSheet.Cells[1, 4], xlsSheet.Cells[1, 4]];
                        rangeSource.EntireColumn.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[1, 4 + 1], xlsSheet.Cells[1, dtMonthRange.Rows.Count + 4]];
                        rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        xlsSheet.Range[xlsSheet.Cells[1, 4 + dtMonthRange.Rows.Count], xlsSheet.Cells[1, 4 + dtMonthRange.Rows.Count + 1]].EntireColumn.delete();

                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[4, dtMonthRange.Rows.IndexOf(drr) + 4] = drr[0];

                        }


                        arrRow = new[] { 16, 12, 6 };
                        iRowNetSales = arrRow[0];
                        iRowSale = arrRow[2];
                        iRowSaleReturn = arrRow[1];
       
                        foreach (int iRow in arrRow)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[iRow, 1], xlsSheet.Cells[iRow, 1]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[iRow + 1, 1], xlsSheet.Cells[iRow + rsSum.RecordCount, 1]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range[xlsSheet.Cells[iRow + rsSum.RecordCount, 1], xlsSheet.Cells[iRow + rsSum.RecordCount, 1]].EntireRow.Delete();
                            iRowNetSales += rsSum.RecordCount;
                        }


                        iRowNetSales = iRowNetSales - (rsSum.RecordCount + 2);
                        iRowSaleReturn = iRowNetSales - (rsSum.RecordCount + 3);


                        //iRowNetSales = rsSum.RecordCount + (iRowNetSales + 2);
                        xlsSheet.Range[xlsSheet.Cells[iRowNetSales, 4], xlsSheet.Cells[(iRowNetSales + rsSum.RecordCount), (dtMonthRange.Rows.Count + 4)]].Formula = "=D" + iRowSale + "-D" + (iRowSale+ (rsSum.RecordCount + 5));
                        xlsSheet.Range["A" + iRowSale].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + (iRowSaleReturn)].CopyFromRecordset(rsSumReturn);
                        xlsSheet.Range["A" + (iRowSale + rsSum.RecordCount)].CopyFromRecordset(rsSumRP);
                        xlsSheet.Range["A" + (iRowSaleReturn+rsSum.RecordCount)].CopyFromRecordset(rsSumRPReturn);

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSaleReturn, 1], xlsSheet.Cells[(iRowSaleReturn + rsSum.RecordCount), 3]];
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowNetSales, 1], xlsSheet.Cells[iRowNetSales + rsSum.RecordCount, 3]];
                        rangeDest.Value = rangeSource.Value;

                        xlsSheet.Range["B" + (iRowSale + rsSum.RecordCount + 1)].CopyFromRecordset(RSMRS);

                        findRang = xlsSheet.Range["C:C"].Find(What: "A811", LookIn: Excel.XlFindLookIn.xlFormulas,
                      LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);


                        System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();
                        adapter.Fill(dtSumRP, rsSumRP);
                        adapter.Fill(dtSumRSMRS, SumRSMRS);


                        for (int i = 3; i < dtSumRP.Columns.Count; i++)
                        {
                            if (dtSumRP.Rows[0][i].ToString() != "" && dtSumRSMRS.Rows[0][i].ToString() != "")
                            {
                                xlsSheet.Cells[findRang.Row, (i + 1)] = Convert.ToDouble(dtSumRP.Rows[0][i].ToString()) - Convert.ToDouble(dtSumRSMRS.Rows[0][i].ToString());

                            }
                            else if (dtSumRP.Rows[0][i].ToString() != "")
                            {

                                xlsSheet.Cells[findRang.Row, (i + 1)] = Convert.ToDouble(dtSumRP.Rows[0][i].ToString());

                            }
                        }


                    
                  
                        int rowA811 = (rsSum.RecordCount * 4) + 5;
                        findRang = xlsSheet.Range["C" + rowA811 + ":C" + rowA811 + 4].Find(What: "A811", LookIn: Excel.XlFindLookIn.xlFormulas,
                       LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                        int saleReturn = (rsSum.RecordCount * 3);
                        findRangReturn = xlsSheet.Range["C" + saleReturn + ":C" + rowA811 + 6].Find(What: "A811", LookIn: Excel.XlFindLookIn.xlFormulas,
                       LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                        if (findRang != null)
                        {
                            if (findRang.Row > 0)
                            {
                                xlsSheet.Range[xlsSheet.Cells[findRang.Row, 4], xlsSheet.Cells[findRang.Row, (dtMonthRange.Rows.Count + 4)]].Formula = "=IFERROR(SUM((D" + (rsSum.RecordCount + 6) + ":D" + (rsSum.RecordCount + (6 + 2)) + "))-D" + findRangReturn.Row + ",0)";

                            }
                        }

                        xlsSheet.Range["A:P"].Columns.EntireColumn.AutoFit();

                    }
                      
                    else if (SalesSummaryOBJ.Factory != "RP")
                    {


                        //end Sheets 4
                        //==============================================================================================================//
                        xlsSheet = xlsBook.Sheets[4];
                        xlsSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

                        xlsSheet = xlsBook.Sheets[2];
                        rsSum = SalesSummaryDAL.getSaleSummaryByCustomer(dtMonthRange, SalesSummaryOBJ, "Total", true);
                        rsSumReturn = SalesSummaryDAL.getSaleSummaryByCustomer(dtMonthRange, SalesSummaryOBJ, "Total", false);

                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;
                        xlsSheet.Name = "Company - Total";
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[1, 4], xlsSheet.Cells[1, 4]];
                        rangeSource.EntireColumn.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[1, 4 + 1], xlsSheet.Cells[1, dtMonthRange.Rows.Count + 4]];
                        rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        xlsSheet.Range[xlsSheet.Cells[1, 4 + dtMonthRange.Rows.Count], xlsSheet.Cells[1, 4 + dtMonthRange.Rows.Count + 1]].EntireColumn.Delete();

                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[4, dtMonthRange.Rows.IndexOf(drr) + 4] = drr[0];
                        }

                        arrRow = new int[] { 14, 10, 6 };
                        iRowNetSales = arrRow[0];
                        iRowSale = arrRow[2];
                        iRowSaleReturn = arrRow[1];


                        foreach (int iRow in arrRow)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[iRow, 1], xlsSheet.Cells[iRow, 1]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[iRow + 1, 1], xlsSheet.Cells[iRow + rsSum.RecordCount, 1]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range[xlsSheet.Cells[iRow + rsSum.RecordCount, 1], xlsSheet.Cells[iRow + rsSum.RecordCount + 1, 1]].EntireRow.Delete();
                        }

                        iRowNetSales = iRowNetSales + ((rsSum.RecordCount - 2) * 2);
                        xlsSheet.Range[xlsSheet.Cells[iRowNetSales, 4], xlsSheet.Cells[iRowNetSales + rsSum.RecordCount, dtMonthRange.Rows.Count + 4]].Formula = "=D" + iRowSale + "-D" + (iRowSale + rsSum.RecordCount + 2);
                        xlsSheet.Range["A" + iRowSale].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + (rsSum.RecordCount + iRowSale + 2)].CopyFromRecordset(rsSumReturn);
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale + rsSum.RecordCount, 3]];
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowNetSales, 1], xlsSheet.Cells[iRowNetSales + rsSum.RecordCount, 3]];
                        rangeDest.Value = rangeSource.Value;



                        //end Sheets 2 //==============================================================================================================//

                        xlsSheet = xlsBook.Sheets[3];
                        rsSum = SalesSummaryDAL.getSaleSummaryByCustomer(dtMonthRange, SalesSummaryOBJ, "Normal", true);
                        rsSumReturn = SalesSummaryDAL.getSaleSummaryByCustomer(dtMonthRange, SalesSummaryOBJ, "Normal", false);
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;
                        xlsSheet.Name = "Company - " + SalesSummaryOBJ.Factory;

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[1, 4], xlsSheet.Cells[1, 4]];
                        rangeSource.EntireColumn.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[1, 4 + 1], xlsSheet.Cells[1, dtMonthRange.Rows.Count + 4]];
                        rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        xlsSheet.Range[xlsSheet.Cells[1, 4 + dtMonthRange.Rows.Count], xlsSheet.Cells[1, 4 + dtMonthRange.Rows.Count + 1]].EntireColumn.Delete();
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[4, dtMonthRange.Rows.IndexOf(drr) + 4] = drr[0];
                        }

                        arrRow = new[] { 14, 10, 6 };
                        iRowNetSales = arrRow[0];
                        iRowSale = arrRow[2];

                        foreach (int iRow in arrRow)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[iRow, 1], xlsSheet.Cells[iRow, 1]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[iRow + 1, 1], xlsSheet.Cells[iRow + rsSum.RecordCount, 1]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range[xlsSheet.Cells[iRow + rsSum.RecordCount, 1], xlsSheet.Cells[iRow + rsSum.RecordCount + 1, 1]].EntireRow.Delete();
                        }


                        iRowNetSales = iRowNetSales + ((rsSum.RecordCount - 2) * 2);
                        xlsSheet.Range[xlsSheet.Cells[iRowNetSales, 4], xlsSheet.Cells[iRowNetSales + rsSum.RecordCount, dtMonthRange.Rows.Count + 4]].Formula = "=D" + iRowSale + "-D" + (iRowSale + (rsSum.RecordCount + 2));
                        xlsSheet.Range["A" + iRowSale].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + (iRowSale + rsSum.RecordCount + 2)].CopyFromRecordset(rsSumReturn);
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale + rsSum.RecordCount, 3]];
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowNetSales, 1], xlsSheet.Cells[iRowNetSales + rsSum.RecordCount, 3]];
                        rangeDest.Value = rangeSource.Value;
                        xlsSheet.Range["A:P"].Columns.EntireColumn.AutoFit();
                    }
                        //end Sheets 3
                        //==============================================================================================================//

                    if (SalesSummaryOBJ.Factory == "GMO")
                    {
                        xlsSheet = xlsBook.Sheets[5];
                        rsSum = SalesSummaryDAL.getSaleSummaryByCustomer(dtMonthRange, SalesSummaryOBJ, "Trading", true);
                        rsSumReturn = SalesSummaryDAL.getSaleSummaryByCustomer(dtMonthRange, SalesSummaryOBJ, "Trading", false);
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;
                        xlsSheet.Name = "Company - Trading";

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[1, 4], xlsSheet.Cells[1, 4]];
                        rangeSource.EntireColumn.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[1, 4 + 1], xlsSheet.Cells[1, dtMonthRange.Rows.Count + 4]];
                        rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        xlsSheet.Range[xlsSheet.Cells[1, 4 + dtMonthRange.Rows.Count], xlsSheet.Cells[1, 4 + dtMonthRange.Rows.Count + 1]].EntireColumn.Delete();
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[4, dtMonthRange.Rows.IndexOf(drr) + 4] = drr[0];
                        }

                        arrRow = new[] { 14, 10, 6 };
                        iRowNetSales = arrRow[0];
                        iRowSale = arrRow[2];

                        foreach (int iRow in arrRow)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[iRow, 1], xlsSheet.Cells[iRow, 1]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[iRow + 1, 1], xlsSheet.Cells[iRow + rsSum.RecordCount, 1]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range[xlsSheet.Cells[iRow + rsSum.RecordCount, 1], xlsSheet.Cells[iRow + rsSum.RecordCount + 1, 1]].EntireRow.Delete();
                        }

                        iRowNetSales = iRowNetSales + ((rsSum.RecordCount - 2) * 2);
                        xlsSheet.Range[xlsSheet.Cells[iRowNetSales, 4], xlsSheet.Cells[iRowNetSales + rsSum.RecordCount, dtMonthRange.Rows.Count + 4]].Formula = "=D" + iRowSale + "-D" + (iRowSale + (rsSum.RecordCount + 2));
                        xlsSheet.Range["A" + iRowSale].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + (iRowSale + rsSum.RecordCount + 2)].CopyFromRecordset(rsSumReturn);
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale + rsSum.RecordCount, 3]];
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowNetSales, 1], xlsSheet.Cells[iRowNetSales + rsSum.RecordCount, 3]];
                        rangeDest.Value = rangeSource.Value;
                        xlsSheet.Range["A:P"].Columns.EntireColumn.AutoFit();

                    }
                      //end Sheets 5
                      //==============================================================================================================//
  
                 
                 
                    xlsSheet = xlsBook.Sheets[8];
                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        rsSum = SalesSummaryDAL.getSalesResultBySalesGroup(Convert.ToDateTime(drr["dt"]), SalesSummaryOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            drr["SalesData"] = true;
                            xlsSheet = xlsBook.Sheets[xlsBook.Sheets.Count];
                            xlsSheet.Copy(Before:xlsBook.Sheets[8]);

                            xlsSheet = xlsBook.Sheets[8];
                            xlsSheet.Name = String.Format("Result Sale - {0:MMM yyyy}", drr["dt"]);
                            xlsSheet.Cells.Font.Name = "Arial";
                            xlsSheet.Cells.Font.Size = 8;

                            ADODB.Recordset rsResultCust = SalesSummaryDAL.getSalesResultByCustomer(Convert.ToDateTime(drr["dt"]), SalesSummaryOBJ, true, "ByMonth");

                            iRowSale = 24;
                            if (rsResultCust.RecordCount > 3)
                            {
                                rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSale + 1, 1], xlsSheet.Cells[iRowSale + 1, 17]];
                                rangeSource.EntireRow.Copy();
                                rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowSale + 2, 1], xlsSheet.Cells[iRowSale + rsResultCust.RecordCount - 2,17]];
                                rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            }

                            xlsSheet.Range["A" + iRowSale].CopyFromRecordset(rsResultCust);
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 9], xlsSheet.Cells[(iRowSale + rsResultCust.RecordCount), 9]].Formula = "=IF(F" + iRowSale + "=0,0,H" + iRowSale + "/F" + iRowSale + ")";
                            
                            iRowSale = 6;
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSale + 1, 1], xlsSheet.Cells[iRowSale + 1, 1]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowSale + 2, 1], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 2, 1]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                            xlsSheet.Range["A" + iRowSale].CopyFromRecordset(rsSum);
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 9], xlsSheet.Cells[(iRowSale + rsSum.RecordCount), 9]].Formula = "=IF(F" + iRowSale + "=0,0,H" + iRowSale + "/F" + iRowSale + ")";

                            int k = iRowSale;
                            for (int j = iRowSale; j <= iRowSale + rsSum.RecordCount - 1; j++ )
                            {
                                if (xlsSheet.Cells[j,2].Text=="TOTAL")
                                {
                                    xlsSheet.Range[xlsSheet.Cells[k, 1], xlsSheet.Cells[(j - 1), 1]].Merge();
                                    xlsSheet.Range[xlsSheet.Cells[k,1],xlsSheet.Cells[k,1]].HorizontalAlignment=Excel.Constants.xlLeft;
                                    xlsSheet.Range[xlsSheet.Cells[k,1],xlsSheet.Cells[k,1]].VerticalAlignment=Excel.Constants.xlCenter;
                                    xlsSheet.Range[xlsSheet.Cells[j,10],xlsSheet.Cells[j,12]].Formula="=SUM(J" + k+":j" + (j-1)+")";
                                
                                    k=j+1;

                                }else if(xlsSheet.Cells[j,2].Text==""){

                                    k+=1;
                                }

                            }

                            int iRowSalesNoCom = iRowSale + rsSum.RecordCount + 1;
                            int iRowSalesTrading = iRowSale + rsSum.RecordCount + 9;

                            ADODB.Recordset rsNoCOM = SalesSummaryDAL.getSalesResultNoCOM(Convert.ToDateTime(drr["dt"]), SalesSummaryOBJ, true);
                            ADODB.Recordset rsNoCOMReturn = SalesSummaryDAL.getSalesResultNoCOM(Convert.ToDateTime(drr["dt"]), SalesSummaryOBJ, false);
                            if (rsNoCOM.RecordCount > 0)
                            {
                                xlsSheet.Range["F" + iRowSalesNoCom].CopyFromRecordset(rsNoCOM);
                                xlsSheet.Range["J" + iRowSalesNoCom].CopyFromRecordset(rsNoCOMReturn);
                            }

                            ADODB.Recordset rsTrading = SalesSummaryDAL.getSalesResultTrading(Convert.ToDateTime(drr["dt"]), SalesSummaryOBJ,1, true);
                            ADODB.Recordset rsTradingReturn = SalesSummaryDAL.getSalesResultTrading(Convert.ToDateTime(drr["dt"]), SalesSummaryOBJ,1,false);
                          
                           // ADODB.Recordset rsTrading = SalesSummaryDAL.getSalesResultTrading(Convert.ToDateTime(dr["dt"]), SalesSummaryOBJ,1);
                            if (rsTrading.RecordCount > 0)
                            {
                                xlsSheet.Range["F" + iRowSalesTrading].CopyFromRecordset(rsTrading);
                                xlsSheet.Range["J" + iRowSalesTrading].CopyFromRecordset(rsTradingReturn);

                            }


                            xlsSheet.Cells.Replace(What: "ZZZZZZZZZ", Replacement:"OTHER", LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);
                            xlsSheet.Cells.Replace(What: "ZZZZZZ", Replacement: "OT", LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);


                            findRang = xlsSheet.Range["B:B"].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                              LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                            Excel.Range find;

                            if (SalesSummaryOBJ.Factory == "GMO")
                            {
                                find = xlsSheet.Range["A:A"].Find(What: "SALE - TRADING", LookIn: Excel.XlFindLookIn.xlFormulas,
                                LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                                xlsSheet.Range[xlsSheet.Cells[find.Row - 1, 10], xlsSheet.Cells[find.Row - 1, 12]].Formula = "=SUMIF($A$" + 6 + ":$A$" + (findRang.Row - 2) + @",""TOTAL INTERNAL SALE*"",J$" + 6 + ":J$" + (findRang.Row - 2) + ")";
                                xlsSheet.Range[xlsSheet.Cells[findRang.Row - 1, 10], xlsSheet.Cells[findRang.Row - 1, 12]].Formula = "=SUMIF($A$" + 6 + ":$A$" + (findRang.Row - 2) + @",""TOTAL EXTERNAL SALE*""," + "J$" + 6 + ":J$" + (findRang.Row - 2) + ")";
                            }
                            else
                            {

                                find = xlsSheet.Range["A:A"].Find(What: "INTERNAL SALE", LookIn: Excel.XlFindLookIn.xlFormulas,
                                                                LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                                xlsSheet.Range[xlsSheet.Cells[find.Row - 1, 10], xlsSheet.Cells[find.Row - 1, 12]].Formula = "=SUMIF($A$" + 6 + ":$A$" + (find.Row - 2) + @",""TOTAL EXTERNAL SALE*"",J$" + 6 + ":J$" + (find.Row - 2) + ")";

                                xlsSheet.Range[xlsSheet.Cells[findRang.Row - 1, 10], xlsSheet.Cells[findRang.Row - 1, 12]].Formula = "=SUMIF($A$" + 6 +":$A$" +(findRang.Row - 2)+ @",""TOTAL INTERNAL SALE*"",J$" + 6 + ":J$" + (findRang.Row - 2) + ")";
   
                            }

                            findRang = xlsSheet.Range["B:B"].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                                LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                            xlsSheet.Range[xlsSheet.Cells[findRang.Row, 10], xlsSheet.Cells[findRang.Row, 12]].Formula = "=SUM(J" + (findRang.Row + 7) + ":J" + (findRang.Row + 8) + ")";


                            for (int j = iRowSale; j <= (iRowSale + rsSum.RecordCount);j++ )
                            {
                                if (xlsSheet.Cells[j, 2].Text == "TOTAL")
                                {
                                    for (int sets = 9; sets <= 17; sets++)
                                    {
                                        xlsSheet.Range[xlsSheet.Cells[j, sets], xlsSheet.Cells[j, sets]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-3],0)";
                                        sets += 3;

                                    }

                                }
                                else if (xlsSheet.Cells[j, 2].Text == "GRAND TOTAL")
                                {
                                    for (int sets = 9; sets <= 17; sets++)
                                    {
                                        xlsSheet.Range[xlsSheet.Cells[(j-1), sets], xlsSheet.Cells[(j-1), sets]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-3],0)";
                                        xlsSheet.Range[xlsSheet.Cells[j, sets], xlsSheet.Cells[j, sets]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-3],0)";
                                        sets += 3;

                                    }

                                }

                            }//end for


                            xlsSheet.Range["A:Q"].EntireColumn.AutoFit();


                        }//end if


                    }//end for

                  

                    //=============================================================================== END SHEET 8 ===========================================================

                    xlsBook.Sheets[xlsBook.Worksheets.Count].delete();
                    xlsSheet = xlsBook.Sheets[6];
                    xlsSheet.Activate();

                    rsSum = SalesSummaryDAL.getSalesResultBySalesGroup(SalesSummaryOBJ.DateFrom, SalesSummaryOBJ, true);
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;
                    xlsSheet.Name = "Results Sale Total";

                    if (rsSum.RecordCount > 0)
                    {
                        ADODB.Recordset rsResultCust = SalesSummaryDAL.getSalesResultByCustomer(SalesSummaryOBJ.DateFrom, SalesSummaryOBJ, true, "Total");
                        iRowSale = 24;
                        rangeSource=xlsSheet.Range[xlsSheet.Cells[iRowSale+1,1],xlsSheet.Cells[iRowSale+1,1]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowSale + 2, 1], xlsSheet.Cells[iRowSale + rsResultCust.RecordCount - 2, 1]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + iRowSale].CopyFromRecordset(rsResultCust);
                    

                    iRowSale = 6;
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSale + 1, 1], xlsSheet.Cells[iRowSale + 1, 1]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowSale + 2, 1], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 2, 1]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale, 4]].CopyFromRecordset(rsSum);

                    xlsSheet.Cells.Replace(What: "ZZZZZZZZZ", Replacement: "OTHER", LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);
                    xlsSheet.Cells.Replace(What: "ZZZZZZ", Replacement: "OT", LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

                    int[] arrMonCol = new [] {58, 62, 66, 6, 10, 14, 22, 26, 30, 42, 46, 50};

                    Excel.Worksheet xlsSheetByMonth = new Excel.Worksheet();
                    Excel.Range fRange;


                    foreach (DataRow item in dtMonthRange.Rows)
                    {
                        DateTime dt = new DateTime();
                       // DateTime.TryParse(drr("dt"), dt);
                        dt = Convert.ToDateTime(item["dt"]);

                        if (item["SalesData"].Equals(true))
                        {
                            xlsSheetByMonth = xlsBook.Sheets[xlsBook.Sheets.Count - dtMonthRange.Rows.IndexOf(item)];

                            fRange = xlsSheet.Range["B:B"].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                            LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                            //4

                            xlsSheet.Range[xlsSheet.Cells[iRowSale, arrMonCol[dt.Month - 1]], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, arrMonCol[dt.Month - 1] + 2]].Formula =

                            String.Format(@"=IF('Results Sale Total'!$C{0}="""","""",SUMIFS('{2}'!N${0}:N${1}" +
                              ",'{2}'!$C${0}:$C${1},'Results Sale Total'!$C{0}" + ",'{2}'!$B${0}:$B${1},'Results Sale Total'!$B{0}))", iRowSale, fRange.Row, xlsSheetByMonth.Name);



                            Excel.Range rangSalesNoCOM = xlsSheet.Cells.Find(What: "SALE NO COMMERCIAL");

                            Excel.Range rangSalesNoCOMbyMonth = xlsSheetByMonth.Cells.Find(What: "SALE NO COMMERCIAL");


                            if (rangSalesNoCOM != null || rangSalesNoCOMbyMonth !=  null)
                            {
                                xlsSheet.Range[xlsSheet.Cells[rangSalesNoCOM.Row, arrMonCol[dt.Month - 1]], xlsSheet.Cells[rangSalesNoCOM.Row, arrMonCol[dt.Month - 1] + 2]].Formula = String.Format(@"='{0}'!N{1}", xlsSheetByMonth.Name, rangSalesNoCOMbyMonth.Row);
                                int rowSaleTrading = rangSalesNoCOM.Row + 8;
                                int rowSaleTradingbyMonth = rangSalesNoCOMbyMonth.Row + 8;
                                xlsSheet.Range[xlsSheet.Cells[rowSaleTrading, arrMonCol[dt.Month - 1]], xlsSheet.Cells[rowSaleTrading + 1, arrMonCol[dt.Month - 1] + 2]].Formula = String.Format(@"='{0}'!N{1}", xlsSheetByMonth.Name, rowSaleTradingbyMonth);

                                int rowSumCustomer = rangSalesNoCOM.Row + 14;
                                int rowSumCustomerEnd = rowSumCustomer + rsResultCust.RecordCount - 1;
                                int rowSaleCustomerbyMonth = rangSalesNoCOMbyMonth.Row + 14;



                                xlsSheet.Range[xlsSheet.Cells[rowSumCustomer, arrMonCol[dt.Month - 1]], xlsSheet.Cells[rowSumCustomerEnd, arrMonCol[dt.Month - 1] + 2]].Formula = String.Format(@"=SUMIF('{0}'!$B${1}:$B${2},'Results Sale Total'!$B{3},'{0}'!N${1}:N${2})", xlsSheetByMonth.Name,
                                                 rowSaleCustomerbyMonth, xlsSheetByMonth.Range["A" + xlsSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row, rowSumCustomer, rowSumCustomerEnd);

                            }
                        }
                    } //foreach


                    //-Formular for each Total-----------------
                    int j = iRowSale;
                    int k = iRowSale;

                    for (int i = iRowSale; i <= iRowSale + rsSum.RecordCount - 1;i++ )
                    {

                        if (xlsSheet.Cells[i, 2].Text== "TOTAL")
                        {
                            xlsSheet.Range[xlsSheet.Cells[i, 6], xlsSheet.Cells[i, 68]].Formula =  String.Format(@"=SUMIF($B$" + iRowSale + ":$B$" + (iRowSale + rsSum.RecordCount - 1) + ",$B$" + (i - 1) + ",F$" + iRowSale + ":F$" + (iRowSale + rsSum.RecordCount - 1) + ")");
                           
                            xlsSheet.Range[xlsSheet.Cells[k, 1], xlsSheet.Cells[(i - 1), 1]].Merge();
                            xlsSheet.Range[xlsSheet.Cells[k, 1], xlsSheet.Cells[k,1]].HorizontalAlignment = Excel.Constants.xlLeft;
                            xlsSheet.Range[xlsSheet.Cells[k, 1], xlsSheet.Cells[k,1]].VerticalAlignment = Excel.Constants.xlCenter;
                            k = i + 1;
                        }
                        else if (xlsSheet.Cells[i, 2].Text== "")
                        {
                            k += 1;
                        }//end if

                         if(xlsSheet.Cells[i, 2].Text == "")
                        {
                            xlsSheet.Range[xlsSheet.Cells[i, 6], xlsSheet.Cells[i,68]].Formula = @"=IFERROR(SUMIF($B$" + j + ":$B$" + ( i - 1) + ",$B$" + (i - 1) + ",F$" + j + ":F$" + (i - 1) +"),0)";
                            j = i + 1;
                         }

                    }//end for


                    xlsSheet.Range[xlsSheet.Cells[iRowSale+rsSum.RecordCount-1,6], xlsSheet.Cells[iRowSale+rsSum.RecordCount-1, 68]].Formula = String.Format(@"=IFERROR(SUMIF($B$" + iRowSale + ":$B$" + (iRowSale + rsSum.RecordCount - 2) + @","""",F$" + iRowSale + ":F$" + (iRowSale + rsSum.RecordCount - 2) + "),0)");

                    xlsSheet.Range[xlsSheet.Cells[iRowSale, 18], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 18 + 2]].Formula = String.Format(@"=IFERROR(F" + iRowSale + "+J" + iRowSale + "+N" + iRowSale + ",0)"); //Q1
                    xlsSheet.Range[xlsSheet.Cells[iRowSale, 34], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 34 + 2]].Formula = String.Format(@"=IFERROR(V" + iRowSale + "+Z" + iRowSale + "+AD" + iRowSale + ",0)"); //Q2
                    xlsSheet.Range[xlsSheet.Cells[iRowSale, 38], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 38 + 2]].Formula = String.Format(@"=IFERROR(R" + iRowSale + "+AH" + iRowSale +",0)"); //H1
                    xlsSheet.Range[xlsSheet.Cells[iRowSale, 54], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 54 + 2]].Formula = String.Format(@"=IFERROR(AP" + iRowSale + "+AT" + iRowSale + "+AX" + iRowSale + ",0)"); //Q3
                    xlsSheet.Range[xlsSheet.Cells[iRowSale, 70], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 70 + 2]].Formula = String.Format(@"=IFERROR(BF" + iRowSale + "+BJ" + iRowSale + "+BN" + iRowSale + ",0)"); //Q4
                    xlsSheet.Range[xlsSheet.Cells[iRowSale, 74], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 74 + 2]].Formula = String.Format(@"=IFERROR(BB" + iRowSale + "+BR" + iRowSale +",0)"); //H2
                    xlsSheet.Range[xlsSheet.Cells[iRowSale, 78], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 78 + 2]].Formula = String.Format(@"=IFERROR(AL" + iRowSale + "+BV" + iRowSale +",0)"); //Y



                    for (int i = iRowSale; i <= iRowSale + rsSum.RecordCount - 1; i++)
                    {

                        if (xlsSheet.Cells[i, 2].Text== "TOTAL")
                        {
                            for (int sets = 9; sets <= 68; sets++)
                            {
                                xlsSheet.Range[xlsSheet.Cells[i, sets], xlsSheet.Cells[i , sets]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-3],0)";
                                sets += 3;
                            }


                        }
                        else if (xlsSheet.Cells[i, 2].Text == "GRAND TOTAL")
                        {
                            for (int sets = 9; sets <= 68; sets++)
                            {
                                xlsSheet.Range[xlsSheet.Cells[i-1, sets], xlsSheet.Cells[i-1, sets]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-3],0)";
                                xlsSheet.Range[xlsSheet.Cells[i , sets], xlsSheet.Cells[i , sets]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-3],0)";
                                sets += 3;
                            }
                        }//end if

                        if (xlsSheet.Cells[i, 1].Text == "TOTAL EXTERNAL SALE")
                        {
                            for (int sets = 9; sets <= 68; sets++)
                            {
                                xlsSheet.Range[xlsSheet.Cells[i, sets], xlsSheet.Cells[i, sets]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-3],0)";
                                sets += 3;

                            }
                        }//end if
                    }//end for
                    
                    }//end if rsSum

                    xlsSheet.Range["A:CC"].EntireColumn.AutoFit();
                    xlsSheet.Range["B:B"].EntireColumn.Hidden = true;

      

          //========================= Sale By Application ====
                    if (SalesSummaryOBJ.Factory != "RP")
                    {
                        xlsSheet = xlsBook.Sheets[7];

                        rsSum = rsSum = SalesSummaryDAL.getSalesSummaryByApplication(dtMonthRange, SalesSummaryOBJ);
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;
                        xlsSheet.Name = "SALE BY APPLICATION";

                        if (rsSum.RecordCount > 0)
                        {
                            iRowSale = 17;
                            int[] arrMonCol = new[] { 43, 46, 49, 4, 7, 10, 16, 19, 22, 31, 34, 37 };
                            Excel.Worksheet xlsSheetByMonth = new Excel.Worksheet();
                            Excel.Range fRange;


                            foreach (DataRow item in dtMonthRange.Rows)
                            {
                                DateTime dt = new DateTime();
                                dt = Convert.ToDateTime(item["dt"]);

                                if (item["SalesData"].Equals(true))
                                {
                                    xlsSheetByMonth = xlsBook.Sheets[xlsBook.Sheets.Count - dtMonthRange.Rows.IndexOf(item)];

                                    Excel.Range rangeSaleNoCome = xlsSheet.Cells.Find(What: "SALE NO COMMERCIAL");

                                    Excel.Range rangeSaleNoComeByMonth = xlsSheetByMonth.Cells.Find(What: "SALE NO COMMERCIAL");

                                    if (rangeSaleNoCome != null || rangeSaleNoComeByMonth != null)
                                    {
                                        xlsSheet.Range[xlsSheet.Cells[rangeSaleNoCome.Row, arrMonCol[dt.Month - 1]], xlsSheet.Cells[rangeSaleNoCome.Row, arrMonCol[dt.Month - 1] + 2]].Formula = String.Format(@"='{0}'!N{1}", xlsSheetByMonth.Name, rangeSaleNoComeByMonth.Row);
                                    }


                                    Excel.Range rangeSaleReturn = xlsSheet.Range["A:A"].Find(What: "SALE RETURN", LookIn: Excel.XlFindLookIn.xlFormulas,
                              LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                                    Excel.Range rangeSaleReturnbyMonth = xlsSheetByMonth.Cells.Find(What: "TOTAL COMMERCIAL SALE");


                                    if (rangeSaleReturn != null || rangeSaleReturnbyMonth != null)
                                    {
                                        xlsSheet.Range[xlsSheet.Cells[rangeSaleReturn.Row, arrMonCol[dt.Month - 1]], xlsSheet.Cells[rangeSaleReturn.Row, arrMonCol[dt.Month - 1] + 2]].Formula = String.Format(@"='{0}'!J{1}", xlsSheetByMonth.Name, rangeSaleReturnbyMonth.Row);
                                    }
                                }
                            }//end for


                            iRowSale = 15;
                            ADODB.Recordset rsApplicationList = SalesSummaryDAL.getSalesSummaryByApplicationList(dtMonthRange, SalesSummaryOBJ);

                            if (rsApplicationList.RecordCount > 0)
                            {

                                rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSale + 1, 1], xlsSheet.Cells[iRowSale + 1, 1]];
                                rangeSource.EntireRow.Copy();
                                rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowSale + 2, 1], xlsSheet.Cells[iRowSale + rsApplicationList.RecordCount - 2, 1]];
                                rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                xlsSheet.Range["C" + iRowSale].CopyFromRecordset(rsApplicationList);
                                xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale + rsApplicationList.RecordCount + 1, 1]].Merge();
                                xlsSheet.Cells[iRowSale, 1] = "TOTAL SALE";

                                xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale, 1]].HorizontalAlignment = Excel.Constants.xlCenter;
                                xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale, 1]].VerticalAlignment = Excel.Constants.xlCenter;

                            }


                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 13], xlsSheet.Cells[(iRowSale + rsApplicationList.RecordCount + 2), (13 + 2)]].Formula = String.Format("=D" + (iRowSale) + "+G" + (iRowSale) + "+J" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 25], xlsSheet.Cells[(iRowSale + rsApplicationList.RecordCount + 2), (25 + 2)]].Formula = String.Format("=P" + (iRowSale) + "+S" + (iRowSale) + "+V" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 28], xlsSheet.Cells[(iRowSale + rsApplicationList.RecordCount + 2), (28 + 2)]].Formula = String.Format("=M" + (iRowSale) + "+Y" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 40], xlsSheet.Cells[(iRowSale + rsApplicationList.RecordCount + 2), (40 + 2)]].Formula = String.Format("=AE" + (iRowSale) + "+AH" + (iRowSale) + "+AK" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 52], xlsSheet.Cells[(iRowSale + rsApplicationList.RecordCount + 2), (52 + 2)]].Formula = String.Format("=AQ" + (iRowSale) + "+AT" + (iRowSale) + "+AW" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 55], xlsSheet.Cells[(iRowSale + rsApplicationList.RecordCount + 2), (55 + 2)]].Formula = String.Format("=AN" + (iRowSale) + "+AZ" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 58], xlsSheet.Cells[(iRowSale + rsApplicationList.RecordCount + 2), (58 + 2)]].Formula = String.Format("=AB" + (iRowSale) + "+BC" + (iRowSale));



                            iRowSale = 11;
                            ADODB.Recordset rsTradingByApp = SalesSummaryDAL.getSalesTradingByApplication(dtMonthRange, SalesSummaryOBJ);

                            if (rsTradingByApp.RecordCount > 0)
                            {
                                if (rsTradingByApp.RecordCount > 3)
                                {
                                    rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSale + 1, 1], xlsSheet.Cells[iRowSale + 1, 1]];
                                    rangeSource.EntireRow.Copy();
                                    rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowSale + 2, 1], xlsSheet.Cells[iRowSale + rsTradingByApp.RecordCount - 2, 1]];
                                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                }


                                xlsSheet.Range["C" + iRowSale].CopyFromRecordset(rsTradingByApp);
                                xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale + rsTradingByApp.RecordCount - 1, 1]].Merge();
                                xlsSheet.Cells[iRowSale, 1] = "SALE TRADING" + Environment.NewLine + "(RE-INVOICE HOOP)";

                                xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale, 1]].HorizontalAlignment = Excel.Constants.xlCenter;
                                xlsSheet.Range[xlsSheet.Cells[iRowSale, 1], xlsSheet.Cells[iRowSale, 1]].VerticalAlignment = Excel.Constants.xlCenter;

                            }


                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 13], xlsSheet.Cells[iRowSale + rsTradingByApp.RecordCount, 13 + 2]].Formula = String.Format("=D" + (iRowSale) + "+G" + (iRowSale) + "+J" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 25], xlsSheet.Cells[iRowSale + rsTradingByApp.RecordCount, 25 + 2]].Formula = String.Format("=P" + (iRowSale) + "+S" + (iRowSale) + "+V" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 28], xlsSheet.Cells[iRowSale + rsTradingByApp.RecordCount, 28 + 2]].Formula = String.Format("=M" + (iRowSale) + "+Y" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 40], xlsSheet.Cells[iRowSale + rsTradingByApp.RecordCount, 40 + 2]].Formula = String.Format("=AE" + (iRowSale) + "+AH" + (iRowSale) + "+AK" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 52], xlsSheet.Cells[iRowSale + rsTradingByApp.RecordCount, 52 + 2]].Formula = String.Format("=AQ" + (iRowSale) + "+AT" + (iRowSale) + "+AW" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 55], xlsSheet.Cells[iRowSale + rsTradingByApp.RecordCount, 55 + 2]].Formula = String.Format("=AN" + (iRowSale) + "+AZ" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 58], xlsSheet.Cells[iRowSale + rsTradingByApp.RecordCount, 58 + 2]].Formula = String.Format("=AB" + (iRowSale) + "+BC" + (iRowSale));



                            iRowSale = 6;

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[iRowSale + 1, 1], xlsSheet.Cells[iRowSale + 1, 1]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[iRowSale + 2, 1], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 2, 1]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                            xlsSheet.Range["A" + iRowSale].CopyFromRecordset(rsSum);
                            xlsSheet.Cells.Replace(What: "ZZZZZZZZZ", Replacement: "OTHER", LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);
                            xlsSheet.Cells.Replace(What: "ZZZZZZ", Replacement: "NOCOM", LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 13], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 13 + 2]].Formula = String.Format("=D" + (iRowSale) + "+G" + (iRowSale) + "+J" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 25], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 25 + 2]].Formula = String.Format("=P" + (iRowSale) + "+S" + (iRowSale) + "+V" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 28], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 28 + 2]].Formula = String.Format("=M" + (iRowSale) + "+Y" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 40], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 40 + 2]].Formula = String.Format("=AE" + (iRowSale) + "+AH" + (iRowSale) + "+AK" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 52], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 52 + 2]].Formula = String.Format("=AQ" + (iRowSale) + "+AT" + (iRowSale) + "+AW" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 55], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 55 + 2]].Formula = String.Format("=AN" + (iRowSale) + "+AZ" + (iRowSale));
                            xlsSheet.Range[xlsSheet.Cells[iRowSale, 58], xlsSheet.Cells[iRowSale + rsSum.RecordCount - 1, 58 + 2]].Formula = String.Format("=AB" + (iRowSale) + "+BC" + (iRowSale));



                            int k = iRowSale;

                            for (int i = iRowSale; i <= iRowSale + rsSum.RecordCount - 1; i++)
                            {

                                if (xlsSheet.Cells[i, 2].Text == "TOTAL")
                                {
                                    xlsSheet.Range[xlsSheet.Cells[k, 1], xlsSheet.Cells[i - 1, 1]].Merge();
                                    xlsSheet.Range[xlsSheet.Cells[k, 1], xlsSheet.Cells[k, 1]].HorizontalAlignment = Excel.Constants.xlLeft;
                                    xlsSheet.Range[xlsSheet.Cells[k, 1], xlsSheet.Cells[k, 1]].VerticalAlignment = Excel.Constants.xlCenter;
                                    k = i + 1;
                                }
                                else if (xlsSheet.Cells[i, 2].Text == "")
                                {
                                    k += 1;
                                }//end if

                            }//end for


                            xlsSheet.Range["A:BH"].EntireColumn.AutoFit();
                            xlsSheet.Range["B:B"].EntireColumn.Hidden = true;

                        }//end sale by application

                    }/// END sale by application

       
                xlsSheet = xlsBook.Sheets[1];
               
                 if(SalesSummaryOBJ.Factory =="GMO"){
                     xlsSheet.Cells[14, 1] = "GMO - LENS + GMO - TRADING";

                 }
                 else if (SalesSummaryOBJ.Factory == "RP")
                 {

                     xlsSheet.Cells[14, 1] = "RP";
                 }
                 else
                 {
                     xlsSheet.Cells[14, 1] = "PO";

                 }


                 xlsSheet.Cells[16, 1] = dtMonthRange.Rows[dtMonthRange.Rows.Count-1][1];
                

                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;





                } // end if rsSum

            return "";
            } 
    }
}
