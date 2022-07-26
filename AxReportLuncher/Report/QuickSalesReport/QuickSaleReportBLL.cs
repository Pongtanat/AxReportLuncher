using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;


namespace NewVersion.Report.QuickSales_Report
{
     class QuickSaleReportBLL
    {
        
        QuickSalesReportDAL QuickSalesReportDAL = new QuickSalesReportDAL();

        public string getNumberSequenceGroup(string strFac, int intShipmentLocation)
        {
            string strNumberSequenceGroup = "";
            DataTable dt = QuickSalesReportDAL.getNumberSequenceGroup(strFac, intShipmentLocation);

            if (dt.Rows.Count > 0)
            {
                strNumberSequenceGroup = dt.Rows[0][0].ToString();

            }
            return strNumberSequenceGroup;
        }


        public string getQuickSalesReport(QuickSaleReportOBJ QuickSaleReportOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();

                var strNumberSequence = new DataTable();
                bool trading = false;


                DateTime TempDate = QuickSaleReportOBJ.DateTo; ////
                DateTime dateTo = QuickSaleReportOBJ.DateTo;
                DateTime dateFrom = QuickSaleReportOBJ.DateFrom;
                DateTime LockDateTo = new DateTime(QuickSaleReportOBJ.DateFrom.AddMonths(0).Year, QuickSaleReportOBJ.DateFrom.AddMonths(0).Month, 1);

                var firstDayBeforeMonth = new DateTime(QuickSaleReportOBJ.DateFrom.AddMonths(-1).Year, QuickSaleReportOBJ.DateFrom.AddMonths(-1).Month, 1);
                var lastDayOfBeforeMonth = firstDayBeforeMonth.AddMonths(1).AddDays(-1);


                //===set month for get header
                QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;


                string strSystemPath = System.IO.Directory.GetCurrentDirectory();
                Excel.Application xlsApp = new Excel.Application();
                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                Excel.Range rangeSource, rangeDest;
                Excel.Workbook xlsBookTemplate;



                Excel.Range findRang;

                xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\QuickSales\QuickSale.xls");


                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsSheet = xlsBook.Sheets[1];

                int intStartRow = 10;

                DataTable dt = new DataTable();
                System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();
                Excel.Range xlRangeLine;

                int Col = 11;
                int DtCol = 0;
                strNumberSequence.Columns.Add("NUMBERSEQUENCE");
                string strNocome = "SALE PCS NO COM";

                if (QuickSaleReportOBJ.Factory == "RP")
                {


                    //=======================RP=======================================//
                    xlsSheet = xlsBook.Sheets[6];
                    QuickSaleReportOBJ.Factory = "RP";
                    rsSum = QuickSalesReportDAL.getQuickSalesHeader(QuickSaleReportOBJ, trading);
                    dt.Clear();
                    adapter.Fill(dt, rsSum);
                    intStartRow = 10;
                    DtCol = 0;
                    Col = 11;


                     strNocome = "SALE PCS NO COM";
                    if (rsSum.RecordCount > 0)
                    {


                        for (int i = 1; i <= rsSum.RecordCount - 1; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow + i, 1], xlsSheet.Cells[intStartRow + 3 + (i), 20]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (i + 4), 1], xlsSheet.Cells[(intStartRow) + (i + 4) + 3, 20]];

                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            intStartRow = intStartRow + (3);
                        }
                        xlsSheet.Range[xlsSheet.Cells[(((rsSum.RecordCount + 1) * 4) + 10) - 3, 1], xlsSheet.Cells[((rsSum.RecordCount + 1) * 4) + 10, 20]].EntireRow.Delete();



                        for (int i = 1; i <= rsSum.RecordCount; i++)
                        {
                            xlsSheet.Cells[Col, 4] = dt.Rows[DtCol][0];
                            Col += 4;
                            DtCol += 1;
                        }


                        xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;

                        xlRangeLine = xlsSheet.UsedRange;


                        //Result 1-15=======================================================================================
                        QuickSaleReportOBJ.DateFrom = LockDateTo;
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 11] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                        // }


                        //Result 1-20==================================================================
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                        if (TempDate.Date > LockDateTo.AddDays(14))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;

                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 14] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }


                                }
                            }
                        }



                        //Result 1-25
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                        if (TempDate.Date > LockDateTo.AddDays(19))
                        {

                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;
                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 17] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }


                                }
                            }
                        }


                        //End Of Month
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                        if (TempDate.Date > LockDateTo.AddDays(24))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;

                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 20] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }

                                }
                            }
                        }

                        //Before month
                        dt.Clear();
                        Col = 11;

                        QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                        QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 7] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }

                        }



                        //Nocome
                        findRang = xlsSheet.Range[xlsSheet.Cells[11, 2], xlsSheet.Cells[100, 2]].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                                     LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                        findRang = xlsSheet.Range[xlsSheet.Cells[11, 3], xlsSheet.Cells[findRang.Row, 3]].Find(What: strNocome, LookIn: Excel.XlFindLookIn.xlFormulas,
                                   LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                        //Nocom=============================================================================================


                        strNumberSequence.Clear();

                        QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);

                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["K" + findRang.Row].CopyFromRecordset(rsSum);

                        }



                        //Result 1-20

                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                        if (TempDate.Date > LockDateTo.AddDays(14))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["N" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }


                        //Result 1-25
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                        if (TempDate.Date > LockDateTo.AddDays(19))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["Q" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }




                        //End Of Month

                        QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                        if (TempDate.Date > LockDateTo.AddDays(24))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["T" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }


                        //Before month
                        QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                        QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["G" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                        //End Nocome
                    }//end RP

                }
                else if (QuickSaleReportOBJ.Factory == "PO")
                {

                    //============================PO==============================================================//

                    Col = 11;
                    DtCol = 0;
                    intStartRow = 10;
                    dt.Clear();
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                    QuickSaleReportOBJ.DateFrom = QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                    xlsSheet = xlsBook.Sheets[2];

                    QuickSaleReportOBJ.Factory = "PO";
                    rsSum = QuickSalesReportDAL.getQuickSalesHeader(QuickSaleReportOBJ, trading);
                    adapter.Fill(dt, rsSum);

                    strNocome = "SALE PCS NO COM";
                    if (rsSum.RecordCount > 0)
                    {


                        for (int i = 1; i <= rsSum.RecordCount - 1; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow + i, 1], xlsSheet.Cells[intStartRow + 3 + (i), 20]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (i + 4), 1], xlsSheet.Cells[(intStartRow) + (i + 4) + 3, 20]];

                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            intStartRow = intStartRow + (3);
                        }
                        xlsSheet.Range[xlsSheet.Cells[(((rsSum.RecordCount + 1) * 4) + 10) - 3, 1], xlsSheet.Cells[((rsSum.RecordCount + 1) * 4) + 10, 20]].EntireRow.Delete();



                        for (int i = 1; i <= rsSum.RecordCount; i++)
                        {
                            xlsSheet.Cells[Col, 4] = dt.Rows[DtCol][0];
                            Col += 4;
                            DtCol += 1;
                        }


                        xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;

                        xlRangeLine = xlsSheet.UsedRange;


                        //Result 1-15

                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateFrom = LockDateTo;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 11] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }





                        //Result 1-20

                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                        if (TempDate.Date > LockDateTo.AddDays(14))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;

                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 14] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }


                                }
                            }
                        }



                        //Result 1-25        
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                        if (TempDate.Date > LockDateTo.AddDays(19))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;
                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 17] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }


                                }
                            }
                        }




                        //End Of Month
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                        if (TempDate.Date > LockDateTo.AddDays(24))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;

                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 20] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }

                                }
                            }
                        }




                        //Before Month
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                        QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;


                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 7] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }

                        }



                        //Nocome
                        findRang = xlsSheet.Range[xlsSheet.Cells[11, 2], xlsSheet.Cells[100, 2]].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                                     LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                        findRang = xlsSheet.Range[xlsSheet.Cells[11, 3], xlsSheet.Cells[findRang.Row, 3]].Find(What: strNocome, LookIn: Excel.XlFindLookIn.xlFormulas,
                                   LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                        //Nocom=============================================================================================




                        //Result 1-15
                        QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);

                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["K" + findRang.Row].CopyFromRecordset(rsSum);

                        }



                        //Result 1-20
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                        if (TempDate.Date > LockDateTo.AddDays(14))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["N" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }




                        //Result 1-25
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                        if (TempDate.Date > LockDateTo.AddDays(19))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["Q" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }


                        //End Of Month
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                        if (TempDate.Date > LockDateTo.AddDays(24))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["T" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }

                        //Before Month
                        QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                        QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["G" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                        //End Nocome

                    }//end PO
                }
                else if (QuickSaleReportOBJ.Factory == "GMO")
                {


                    //=================================GMO===================================//


                    Col = 11;
                    DtCol = 0;
                    intStartRow = 10;
                    dt.Clear();
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                    QuickSaleReportOBJ.DateFrom = QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;


                    QuickSaleReportOBJ.Factory = "GMO";
                    rsSum = QuickSalesReportDAL.getQuickSalesHeader(QuickSaleReportOBJ, trading);
                    adapter.Fill(dt, rsSum);

                    trading = true;
                    strNocome = "SALE PCS NO COM GMO LENS";

                    if (rsSum.RecordCount > 0)
                    {
                        //Rows
                        for (int sheet = 5; sheet > 2; sheet--)
                        {
                            xlsSheet = xlsBook.Sheets[sheet];

                            for (int i = 1; i <= rsSum.RecordCount - 1; i++)
                            {
                                rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow + i, 1], xlsSheet.Cells[intStartRow + 3 + (i), 20]];
                                rangeSource.EntireRow.Copy();
                                rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (i + 4), 1], xlsSheet.Cells[(intStartRow) + (i + 4) + 3, 20]];

                                rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                intStartRow = intStartRow + (3);
                            }

                            xlsSheet.Range[xlsSheet.Cells[(((rsSum.RecordCount + 1) * 4) + 10) - 3, 1], xlsSheet.Cells[((rsSum.RecordCount + 1) * 4) + 10, 20]].EntireRow.Delete();



                            for (int i = 1; i <= rsSum.RecordCount; i++)
                            {
                                xlsSheet.Cells[Col, 4] = dt.Rows[DtCol][0];
                                Col += 4;
                                DtCol += 1;
                            }

                            intStartRow = 10;
                            Col = 11;
                            DtCol = 0;
                            xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                        }//end for sheets


                        xlsSheet = xlsBook.Sheets[4];
                        xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;

                        xlRangeLine = xlsSheet.UsedRange;


                        //Result 1-15

                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateFrom = LockDateTo;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 11] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }



                        //Result 1-20
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                        if (TempDate.Date > LockDateTo.AddDays(14))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;

                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 14] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }


                                }
                            }
                        }



                        //Result 1-25
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                        if (TempDate.Date > LockDateTo.AddDays(19))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;
                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 17] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }


                                }
                            }
                        }



                        //End Of Month
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                        if (TempDate.Date > LockDateTo.AddDays(24))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;

                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 20] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }

                                }
                            }
                        }


                        //Before Month
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                        QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;


                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 7] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }

                        }



                        //Nocome
                        findRang = xlsSheet.Range[xlsSheet.Cells[11, 2], xlsSheet.Cells[100, 2]].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                                     LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                        findRang = xlsSheet.Range[xlsSheet.Cells[11, 3], xlsSheet.Cells[findRang.Row, 3]].Find(What: strNocome, LookIn: Excel.XlFindLookIn.xlFormulas,
                                   LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                        //Nocom=============================================================================================


                        strNumberSequence.Clear();


                        QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);


                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["K" + findRang.Row].CopyFromRecordset(rsSum);

                        }


                        //Result 1-20
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                        if (TempDate.Date > LockDateTo.AddDays(14))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["N" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }


                        //LockDateTo.Date > TempDate.Date && LockDateTo.Date <= QuickSaleReportOBJ.DateTo

                        //Result 1-25
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                        if (TempDate.Date > LockDateTo.AddDays(19))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["Q" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }


                        //End Of Month
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                        if (TempDate.Date > LockDateTo.AddDays(24))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["T" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }


                        //Before Month
                        QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                        QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["G" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                        //End Nocome

                    }//end GMO




                    //Trading ========================================================================================
                    if (trading)
                    {

                        xlsSheet = xlsBook.Sheets[5];
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;
                        xlRangeLine = xlsSheet.UsedRange;

                        intStartRow = 10;


                        QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 11] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }



                        //Result 1-20
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                        if (TempDate.Date > LockDateTo.AddDays(14))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;

                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 14] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }


                                }
                            }
                        }




                        //Result 1-25
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                        if (TempDate.Date > LockDateTo.AddDays(19))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;

                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 17] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }


                                }
                            }
                        }


                        //End Of Month
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                        if (TempDate.Date > LockDateTo.AddDays(24))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                            if (rsSum.RecordCount > 0)
                            {
                                adapter.Fill(dt, rsSum);
                                dt = Pivot(dt);
                                for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value != null)
                                    {
                                        Col += 4;

                                        for (int y = 0; y < dt.Rows.Count; y++)
                                        {
                                            if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                            {
                                                for (int colX = 1; colX <= 2; colX++)
                                                {
                                                    xlsSheet.Cells[(Col + colX) - 4, 20] = dt.Rows[y + colX][0];

                                                }

                                            }

                                        }
                                    }


                                }
                            }
                        }




                        //Before Month
                        dt.Clear();
                        Col = 11;
                        QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                        QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 7] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }

                        // Nocome Tranding

                        findRang = xlsSheet.Range[xlsSheet.Cells[11, 2], xlsSheet.Cells[100, 2]].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                                     LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                        findRang = xlsSheet.Range[xlsSheet.Cells[11, 3], xlsSheet.Cells[findRang.Row, 3]].Find(What: strNocome, LookIn: Excel.XlFindLookIn.xlFormulas,
                                   LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                        //Nocom=============================================================================================


                        strNumberSequence.Clear();


                        QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);

                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["K" + findRang.Row].CopyFromRecordset(rsSum);

                        }



                        //Result 1-20

                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                        if (TempDate.Date > LockDateTo.AddDays(14))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["N" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }


                        //Result 1-25
                        //LockDateTo.Date > TempDate.Date && LockDateTo.Date <= QuickSaleReportOBJ.DateTo

                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                        if (TempDate.Date > LockDateTo.AddDays(19))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["Q" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }



                        //End Of Month

                        //LockDateTo.Date > TempDate.Date && LockDateTo.Date <= QuickSaleReportOBJ.DateTo

                        QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                        if (TempDate.Date > LockDateTo.AddDays(24))
                        {
                            rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["T" + findRang.Row].CopyFromRecordset(rsSum);

                            }
                        }


                        //Before Month
                        QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                        QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["G" + findRang.Row].CopyFromRecordset(rsSum);

                        }

                    } //Trading
                }


                /*

                xlsSheet = xlsBook.Sheets[1];
                QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                QuickSaleReportOBJ.DateTo = LockDateTo;
                xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                //External Sale Adishima(PO and GMO)

                xlsSheet.Range[xlsSheet.Cells[12, 6], xlsSheet.Cells[12, 8]].Formula =
                 String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}+'PO LENS'!F{2}", 12, 12, 12);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[13, 6], xlsSheet.Cells[13, 8]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}+'PO LENS'!F{2}", 13, 13, 13);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[12, 11], xlsSheet.Cells[12, 11]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}", 12, 12);//Result month 

                xlsSheet.Range[xlsSheet.Cells[13, 11], xlsSheet.Cells[13, 11]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}+'PO LENS'!K{2}", 13, 13, 13);//Result month 1-15

                xlsSheet.Range[xlsSheet.Cells[12, 14], xlsSheet.Cells[12, 14]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}+'PO LENS'!N{2}", 12, 12, 12);//Result month

                xlsSheet.Range[xlsSheet.Cells[13, 14], xlsSheet.Cells[13, 14]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}+'PO LENS'!N{2}", 13, 13, 13);//Result month 1-20

                xlsSheet.Range[xlsSheet.Cells[12, 17], xlsSheet.Cells[12, 17]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}+'PO LENS'!Q{2}", 12, 12, 12);//Result month

                xlsSheet.Range[xlsSheet.Cells[13, 17], xlsSheet.Cells[13, 17]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}+'PO LENS'!Q{2}", 13, 13, 13);//Result month 1-25

                xlsSheet.Range[xlsSheet.Cells[12, 20], xlsSheet.Cells[12, 20]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}+'PO LENS'!T{2}", 12, 12, 12);//Result month

                xlsSheet.Range[xlsSheet.Cells[13, 20], xlsSheet.Cells[13, 20]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}+'PO LENS'!T{2}", 13, 13, 13);//Result month 1-30


                //External Sale (HOOP)
                xlsSheet.Range[xlsSheet.Cells[16, 6], xlsSheet.Cells[16, 8]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!F{0}", 16);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[17, 6], xlsSheet.Cells[17, 8]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!F{0}", 17);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[16, 11], xlsSheet.Cells[16, 11]].Formula =
                           String.Format(@"='GMO TOTAL FAC'!K{0}", 16);//Result month 

                xlsSheet.Range[xlsSheet.Cells[17, 11], xlsSheet.Cells[17, 11]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!K{0}", 17);//Result month 1-15

                xlsSheet.Range[xlsSheet.Cells[16, 14], xlsSheet.Cells[16, 14]].Formula =
                            String.Format(@"='GMO TOTAL FAC'!N{0}", 16);//Result month

                xlsSheet.Range[xlsSheet.Cells[17, 14], xlsSheet.Cells[17, 14]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!N{0}", 17);//Result month 1-20

                xlsSheet.Range[xlsSheet.Cells[16, 17], xlsSheet.Cells[16, 17]].Formula =
                         String.Format(@"='GMO TOTAL FAC'!Q{0}", 16);//Result month

                xlsSheet.Range[xlsSheet.Cells[17, 17], xlsSheet.Cells[17, 17]].Formula =
                     String.Format(@"='GMO TOTAL FAC'!Q{0}", 17);//Result month 1-25

                xlsSheet.Range[xlsSheet.Cells[16, 20], xlsSheet.Cells[16, 20]].Formula =
                     String.Format(@"='GMO TOTAL FAC'!T{0}", 16);//Result month

                xlsSheet.Range[xlsSheet.Cells[17, 20], xlsSheet.Cells[17, 20]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!T{0}", 17);//Result month 1-30


                //External Sale (HOPA)
                xlsSheet.Range[xlsSheet.Cells[20, 6], xlsSheet.Cells[20, 8]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}+'PO LENS'!F{2}", 20, 16, 16);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[21, 6], xlsSheet.Cells[21, 8]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}+'PO LENS'!F{2}", 21, 17, 17);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[20, 11], xlsSheet.Cells[20, 11]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}+'PO LENS'!K{2}", 20, 16, 16);//Result month 

                xlsSheet.Range[xlsSheet.Cells[21, 11], xlsSheet.Cells[21, 11]].Formula =
                     String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}+'PO LENS'!K{2}", 21, 17, 17);//Result month 1-15

                xlsSheet.Range[xlsSheet.Cells[20, 14], xlsSheet.Cells[20, 14]].Formula =
                      String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}+'PO LENS'!N{2}", 20, 16, 16);//Result month

                xlsSheet.Range[xlsSheet.Cells[21, 14], xlsSheet.Cells[21, 14]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}+'PO LENS'!N{2}", 21, 17, 17);//Result month 1-20

                xlsSheet.Range[xlsSheet.Cells[20, 17], xlsSheet.Cells[20, 17]].Formula =
                       String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}+'PO LENS'!Q{2}", 20, 16, 16);//Result month

                xlsSheet.Range[xlsSheet.Cells[21, 17], xlsSheet.Cells[21, 17]].Formula =
                     String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}+'PO LENS'!Q{2}", 21, 17, 17);//Result month 1-25

                xlsSheet.Range[xlsSheet.Cells[20, 20], xlsSheet.Cells[20, 20]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}+'PO LENS'!T{2}", 20, 16, 16);//Result month

                xlsSheet.Range[xlsSheet.Cells[21, 20], xlsSheet.Cells[21, 20]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}+'PO LENS'!T{2}", 21, 17, 17);//Result month 1-30





                //External Sale (HOWT)
                xlsSheet.Range[xlsSheet.Cells[24, 6], xlsSheet.Cells[24, 8]].Formula =
                  String.Format(@"='RP FAC'!F{0}", 20);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[25, 6], xlsSheet.Cells[25, 8]].Formula =
                     String.Format(@"='RP FAC'!F{0}", 21);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[24, 11], xlsSheet.Cells[24, 11]].Formula =
                   String.Format(@"='RP FAC'!K{0}", 20);//Result month 

                xlsSheet.Range[xlsSheet.Cells[25, 11], xlsSheet.Cells[25, 11]].Formula =
                    String.Format(@"='RP FAC'!K{0}", 21);//Result month 1-15

                xlsSheet.Range[xlsSheet.Cells[24, 14], xlsSheet.Cells[24, 14]].Formula =
                      String.Format(@"='RP FAC'!N{0}", 20);//Result month

                xlsSheet.Range[xlsSheet.Cells[25, 14], xlsSheet.Cells[25, 14]].Formula =
                    String.Format(@"='RP FAC'!N{0}", 21);//Result month 1-20

                xlsSheet.Range[xlsSheet.Cells[24, 17], xlsSheet.Cells[24, 17]].Formula =
                     String.Format(@"='RP FAC'!Q{0}", 20);//Result month

                xlsSheet.Range[xlsSheet.Cells[25, 17], xlsSheet.Cells[25, 17]].Formula =
                    String.Format(@"='RP FAC'!Q{0}", 21);//Result month 1-25

                xlsSheet.Range[xlsSheet.Cells[24, 20], xlsSheet.Cells[24, 20]].Formula =
                   String.Format(@"='RP FAC'!T{0}", 20);//Result month

                xlsSheet.Range[xlsSheet.Cells[25, 20], xlsSheet.Cells[25, 20]].Formula =
                   String.Format(@"='RP FAC'!T{0}", 21);//Result month 1-30



                //External Sale (NIKON)
                xlsSheet.Range[xlsSheet.Cells[28, 6], xlsSheet.Cells[28, 8]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}", 24, 24);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[29, 6], xlsSheet.Cells[29, 8]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}", 25, 25);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[28, 11], xlsSheet.Cells[28, 11]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}", 24, 24);//Result month 

                xlsSheet.Range[xlsSheet.Cells[29, 11], xlsSheet.Cells[29, 11]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}", 25, 25);//Result month 1-15

                xlsSheet.Range[xlsSheet.Cells[28, 14], xlsSheet.Cells[28, 14]].Formula =
                     String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}", 24, 24);//Result month

                xlsSheet.Range[xlsSheet.Cells[29, 14], xlsSheet.Cells[29, 14]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}", 25, 25);//Result month 1-20

                xlsSheet.Range[xlsSheet.Cells[28, 17], xlsSheet.Cells[28, 17]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}", 24, 24);//Result month

                xlsSheet.Range[xlsSheet.Cells[29, 17], xlsSheet.Cells[29, 17]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}", 25, 25);//Result month 1-25

                xlsSheet.Range[xlsSheet.Cells[28, 20], xlsSheet.Cells[28, 20]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}", 24, 24);//Result month

                xlsSheet.Range[xlsSheet.Cells[29, 20], xlsSheet.Cells[29, 20]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}", 25, 25);//Result month 1-30



                //External Sale (RICOH MANUFACTURING)
                xlsSheet.Range[xlsSheet.Cells[32, 6], xlsSheet.Cells[32, 8]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!F{0}", 28);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[33, 6], xlsSheet.Cells[33, 8]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!F{0}", 29);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[32, 11], xlsSheet.Cells[32, 11]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!K{0}", 28);//Result month 

                xlsSheet.Range[xlsSheet.Cells[33, 11], xlsSheet.Cells[33, 11]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!K{0}", 29);//Result month 1-15

                xlsSheet.Range[xlsSheet.Cells[32, 14], xlsSheet.Cells[32, 14]].Formula =
                     String.Format(@"='GMO TOTAL FAC'!N{0}", 28);//Result month

                xlsSheet.Range[xlsSheet.Cells[33, 14], xlsSheet.Cells[33, 14]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!N{0}", 29);//Result month 1-20

                xlsSheet.Range[xlsSheet.Cells[32, 17], xlsSheet.Cells[32, 17]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!Q{0}", 28);//Result month

                xlsSheet.Range[xlsSheet.Cells[33, 17], xlsSheet.Cells[33, 17]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!Q{0}", 29);//Result month 1-25

                xlsSheet.Range[xlsSheet.Cells[32, 20], xlsSheet.Cells[32, 20]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!T{0}", 28);//Result month

                xlsSheet.Range[xlsSheet.Cells[33, 20], xlsSheet.Cells[33, 20]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!T{0}", 29);//Result month 1-30



                //External Sale (SONY THAI)
                xlsSheet.Range[xlsSheet.Cells[36, 6], xlsSheet.Cells[36, 8]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!F{0}+'PO LENS'!F{1}", 32, 20);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[37, 6], xlsSheet.Cells[37, 8]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!F{0}+'PO LENS'!F{1}", 33, 21);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[36, 11], xlsSheet.Cells[36, 11]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!K{0}+'PO LENS'!K{1}", 32, 20);//Result month 

                xlsSheet.Range[xlsSheet.Cells[37, 11], xlsSheet.Cells[37, 11]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!K{0}+'PO LENS'!K{1}", 33, 21);//Result month 1-15

                xlsSheet.Range[xlsSheet.Cells[36, 14], xlsSheet.Cells[36, 14]].Formula =
                     String.Format(@"='GMO TOTAL FAC'!N{0}+'PO LENS'!N{1}", 32, 20);//Result month

                xlsSheet.Range[xlsSheet.Cells[37, 14], xlsSheet.Cells[37, 14]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!N{0}+'PO LENS'!N{1}", 33, 21);//Result month 1-20

                xlsSheet.Range[xlsSheet.Cells[36, 17], xlsSheet.Cells[36, 17]].Formula =
                    String.Format(@"='GMO TOTAL FAC'!Q{0}+'PO LENS'!Q{1}", 32, 20);//Result month

                xlsSheet.Range[xlsSheet.Cells[37, 17], xlsSheet.Cells[37, 17]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!Q{0}+'PO LENS'!Q{1}", 33, 21);//Result month 1-25

                xlsSheet.Range[xlsSheet.Cells[36, 20], xlsSheet.Cells[36, 20]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!T{0}+'PO LENS'!T{1}", 32, 20);//Result month

                xlsSheet.Range[xlsSheet.Cells[37, 20], xlsSheet.Cells[37, 20]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!T{0}+'PO LENS'!T{1}", 33, 21);//Result month 1-30




                //External Sale (PCS NO COME)
                xlsSheet.Range[xlsSheet.Cells[39, 6], xlsSheet.Cells[39, 8]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!F{0}+'PO LENS'!F{1}+'RP FAC'!F{2}", 35, 39, 31);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[39, 11], xlsSheet.Cells[39, 11]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!K{0}+'PO LENS'!K{1}+'RP FAC'!K{2}", 35, 39, 31);//Result month 1-15

                xlsSheet.Range[xlsSheet.Cells[39, 14], xlsSheet.Cells[39, 14]].Formula =
                String.Format(@"='GMO TOTAL FAC'!N{0}+'PO LENS'!N{1}+'RP FAC'!N{2}", 35, 39, 31);//Result month 1-20

                xlsSheet.Range[xlsSheet.Cells[39, 17], xlsSheet.Cells[39, 17]].Formula =
               String.Format(@"='GMO TOTAL FAC'!Q{0}+'PO LENS'!Q{1}+'RP FAC'!Q{2}", 35, 39, 31);//Result month 1-25

                xlsSheet.Range[xlsSheet.Cells[39, 20], xlsSheet.Cells[39, 20]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!T{0}+'PO LENS'!T{1}+'RP FAC'!T{2}", 35, 39, 31);//Result month 1-30


                //External Sale (PARTIAL LENS)
                xlsSheet.Range[xlsSheet.Cells[40, 6], xlsSheet.Cells[40, 8]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!F{0}", 36);//3Q PLAN //Last Month //estimaate

                xlsSheet.Range[xlsSheet.Cells[40, 11], xlsSheet.Cells[40, 11]].Formula =
                   String.Format(@"='GMO TOTAL FAC'!K{0}", 36);//Result month 1-15

                xlsSheet.Range[xlsSheet.Cells[40, 14], xlsSheet.Cells[40, 14]].Formula =
                String.Format(@"='GMO TOTAL FAC'!N{0}", 36);//Result month 1-20

                xlsSheet.Range[xlsSheet.Cells[40, 17], xlsSheet.Cells[40, 17]].Formula =
               String.Format(@"='GMO TOTAL FAC'!Q{0}", 36);//Result month 1-25

                xlsSheet.Range[xlsSheet.Cells[40, 20], xlsSheet.Cells[40, 20]].Formula =
                  String.Format(@"='GMO TOTAL FAC'!T{0}", 36);//Result month 1-30


                */




                xlsApp.SheetsInNewWorkbook = 3;
                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;

                rsSum = null;
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }// end getQuickSalesReport


        public string getQuickSalesReportAll(QuickSaleReportOBJ QuickSaleReportOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
              
                var strNumberSequence = new DataTable();
                bool trading = false;


                DateTime TempDate = QuickSaleReportOBJ.DateTo; ////
                DateTime dateTo = QuickSaleReportOBJ.DateTo;
                DateTime dateFrom = QuickSaleReportOBJ.DateFrom;
                DateTime LockDateTo = new DateTime(QuickSaleReportOBJ.DateFrom.AddMonths(0).Year, QuickSaleReportOBJ.DateFrom.AddMonths(0).Month, 1);

               var firstDayBeforeMonth = new DateTime(QuickSaleReportOBJ.DateFrom.AddMonths(-1).Year, QuickSaleReportOBJ.DateFrom.AddMonths(-1).Month, 1);
               var lastDayOfBeforeMonth = firstDayBeforeMonth.AddMonths(1).AddDays(-1);
      

                //===set month for get header
               QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
              

                string strSystemPath = System.IO.Directory.GetCurrentDirectory();
                Excel.Application xlsApp = new Excel.Application();
                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                Excel.Range rangeSource, rangeDest;
                Excel.Workbook xlsBookTemplate;
             

          
                Excel.Range findRang;

                xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\QuickSales\QuickSale.xls");


                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsSheet = xlsBook.Sheets[1];

                int intStartRow = 10;

                DataTable dt = new DataTable();
                System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();
                Excel.Range xlRangeLine;

                int Col = 11;
                int DtCol = 0;
                strNumberSequence.Columns.Add("NUMBERSEQUENCE");
               

                //=======================RP=======================================//
                xlsSheet = xlsBook.Sheets[6];
                QuickSaleReportOBJ.Factory = "RP";
                rsSum = QuickSalesReportDAL.getQuickSalesHeader(QuickSaleReportOBJ, trading);
                 dt.Clear();
                adapter.Fill(dt, rsSum);
                intStartRow = 10;
                DtCol = 0;
                Col = 11;

       
                string strNocome = "SALE PCS NO COM";
                if (rsSum.RecordCount > 0)
                {


                    for (int i = 1; i <= rsSum.RecordCount - 1; i++)
                    {
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow + i, 1], xlsSheet.Cells[intStartRow + 3 + (i), 20]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (i + 4), 1], xlsSheet.Cells[(intStartRow) + (i + 4) + 3, 20]];

                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        intStartRow = intStartRow + (3);
                    }
                    xlsSheet.Range[xlsSheet.Cells[(((rsSum.RecordCount + 1) * 4) + 10) - 3, 1], xlsSheet.Cells[((rsSum.RecordCount + 1) * 4) + 10, 20]].EntireRow.Delete();



                    for (int i = 1; i <= rsSum.RecordCount; i++)
                    {
                        xlsSheet.Cells[Col, 4] = dt.Rows[DtCol][0];
                        Col += 4;
                        DtCol += 1;
                    }


                    xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                   xlRangeLine = xlsSheet.UsedRange;


                    //Result 1-15=======================================================================================
                    QuickSaleReportOBJ.DateFrom = LockDateTo;             
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 11] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                   // }


                    //Result 1-20==================================================================
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                    if (TempDate.Date > LockDateTo.AddDays(14))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 14] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                    }
                   


                    //Result 1-25
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                    if (TempDate.Date > LockDateTo.AddDays(19))
                    {

                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;
                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 17] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                    }
                 

                    //End Of Month
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                    if (TempDate.Date > LockDateTo.AddDays(24))
                    {
                       rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 20] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }

                            }
                        }
                    }
                 
                    //Before month
                    dt.Clear();
                    Col = 11;

                    QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                    QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                    rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                    if (rsSum.RecordCount > 0)
                    {
                        adapter.Fill(dt, rsSum);
                        dt = Pivot(dt);
                        for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                        {
                            if (xlRangeLine.Cells[i, 4].Value != null)
                            {
                                Col += 4;

                                for (int y = 0; y < dt.Rows.Count; y++)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                    {
                                        for (int colX = 1; colX <= 2; colX++)
                                        {
                                            xlsSheet.Cells[(Col + colX) - 4, 7] = dt.Rows[y + colX][0];

                                        }

                                    }

                                }
                            }


                        }

                    }



                    //Nocome
                    findRang = xlsSheet.Range[xlsSheet.Cells[11, 2], xlsSheet.Cells[100, 2]].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                                 LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                    findRang = xlsSheet.Range[xlsSheet.Cells[11, 3], xlsSheet.Cells[findRang.Row, 3]].Find(What: strNocome, LookIn: Excel.XlFindLookIn.xlFormulas,
                               LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                    //Nocom=============================================================================================


                    strNumberSequence.Clear();

                    QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
   
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["K" + findRang.Row].CopyFromRecordset(rsSum);

                        }

                

                    //Result 1-20
                  
                     QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                     if (TempDate.Date > LockDateTo.AddDays(14))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["N" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }


                    //Result 1-25
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                    if (TempDate.Date > LockDateTo.AddDays(19))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["Q" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }




                    //End Of Month

                    QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                    if (TempDate.Date > LockDateTo.AddDays(24))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["T" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }


                    //Before month
                    QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                    QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                    rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                    if (rsSum.RecordCount > 0)
                    {
                        xlsSheet.Range["G" + findRang.Row].CopyFromRecordset(rsSum);

                    }
                    //End Nocome
                }//end RP



                 //============================PO==============================================================//

                Col = 11;
                DtCol = 0;
                intStartRow = 10;
                dt.Clear();
                QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                QuickSaleReportOBJ.DateFrom = QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                xlsSheet = xlsBook.Sheets[2];

                QuickSaleReportOBJ.Factory = "PO";
                rsSum = QuickSalesReportDAL.getQuickSalesHeader(QuickSaleReportOBJ, trading);
                adapter.Fill(dt, rsSum);
       
                strNocome = "SALE PCS NO COM";
                if (rsSum.RecordCount > 0)
                {


                    for (int i = 1; i <= rsSum.RecordCount - 1; i++)
                    {
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow + i, 1], xlsSheet.Cells[intStartRow + 3 + (i), 20]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (i + 4), 1], xlsSheet.Cells[(intStartRow) + (i + 4) + 3, 20]];

                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        intStartRow = intStartRow + (3);
                    }
                    xlsSheet.Range[xlsSheet.Cells[(((rsSum.RecordCount + 1) * 4) + 10) - 3, 1], xlsSheet.Cells[((rsSum.RecordCount + 1) * 4) + 10, 20]].EntireRow.Delete();



                    for (int i = 1; i <= rsSum.RecordCount; i++)
                    {
                        xlsSheet.Cells[Col, 4] = dt.Rows[DtCol][0];
                        Col += 4;
                        DtCol += 1;
                    }


                    xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlRangeLine = xlsSheet.UsedRange;


                    //Result 1-15
                    
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateFrom = LockDateTo;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14); 
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 11] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
              




                    //Result 1-20
                   
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                    if (TempDate.Date > LockDateTo.AddDays(14))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 14] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                    }
                  


                    //Result 1-25        
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo =LockDateTo.AddDays(24);
                    if (TempDate.Date > LockDateTo.AddDays(19))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;
                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 17] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                    }
                 



                    //End Of Month
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo =LockDateTo.AddMonths(1).AddDays(-1);
                    if (TempDate.Date > LockDateTo.AddDays(24))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 20] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }

                            }
                        }
                    }
                   



                    //Before Month
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                    QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;


                    rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                    if (rsSum.RecordCount > 0)
                    {
                        adapter.Fill(dt, rsSum);
                        dt = Pivot(dt);
                        for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                        {
                            if (xlRangeLine.Cells[i, 4].Value != null)
                            {
                                Col += 4;

                                for (int y = 0; y < dt.Rows.Count; y++)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                    {
                                        for (int colX = 1; colX <= 2; colX++)
                                        {
                                            xlsSheet.Cells[(Col + colX) - 4, 7] = dt.Rows[y + colX][0];

                                        }

                                    }

                                }
                            }


                        }

                    }



                    //Nocome
                    findRang = xlsSheet.Range[xlsSheet.Cells[11, 2], xlsSheet.Cells[100, 2]].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                                 LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                    findRang = xlsSheet.Range[xlsSheet.Cells[11, 3], xlsSheet.Cells[findRang.Row, 3]].Find(What: strNocome, LookIn: Excel.XlFindLookIn.xlFormulas,
                               LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                    //Nocom=============================================================================================


                    

                    //Result 1-15
                    QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
   
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["K" + findRang.Row].CopyFromRecordset(rsSum);

                        }



                    //Result 1-20
                        QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                        if (TempDate.Date > LockDateTo.AddDays(14))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["N" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }



              
                    //Result 1-25
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                    if (TempDate.Date > LockDateTo.AddDays(19))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["Q" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }


                    //End Of Month
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                    if (TempDate.Date > LockDateTo.AddDays(24))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["T" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }

                    //Before Month
                    QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                    QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                    rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                    if (rsSum.RecordCount > 0)
                    {
                        xlsSheet.Range["G" + findRang.Row].CopyFromRecordset(rsSum);

                    }
                    //End Nocome

                }//end PO



                //=================================GMO===================================//
             

                Col = 11;
                DtCol = 0;
                intStartRow = 10;
                dt.Clear();
                QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                QuickSaleReportOBJ.DateFrom = QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
           

                QuickSaleReportOBJ.Factory = "GMO";
                rsSum = QuickSalesReportDAL.getQuickSalesHeader(QuickSaleReportOBJ, trading);
                adapter.Fill(dt, rsSum);

                trading = true;
                strNocome = "SALE PCS NO COM GMO LENS";

                if (rsSum.RecordCount > 0)
                {
                    //Rows
                    for (int sheet = 5; sheet >2; sheet--)
                    {
                        xlsSheet = xlsBook.Sheets[sheet];

                        for (int i = 1; i <= rsSum.RecordCount - 1; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow + i, 1], xlsSheet.Cells[intStartRow + 3 + (i), 20]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow) + (i + 4), 1], xlsSheet.Cells[(intStartRow) + (i + 4) + 3, 20]];

                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            intStartRow = intStartRow + (3);
                        }

                        xlsSheet.Range[xlsSheet.Cells[(((rsSum.RecordCount + 1) * 4) + 10) - 3, 1], xlsSheet.Cells[((rsSum.RecordCount + 1) * 4) + 10, 20]].EntireRow.Delete();



                        for (int i = 1; i <= rsSum.RecordCount; i++)
                        {
                            xlsSheet.Cells[Col, 4] = dt.Rows[DtCol][0];
                            Col += 4;
                            DtCol += 1;
                        }

                        intStartRow = 10;
                        Col = 11;
                        DtCol = 0;
                        xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                    }//end for sheets


                    xlsSheet = xlsBook.Sheets[4];
                    xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlRangeLine = xlsSheet.UsedRange;


                    //Result 1-15
                  
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateFrom = LockDateTo;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
                    rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 11] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                  


                    //Result 1-20
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo =LockDateTo.AddDays(19);
                    if (TempDate.Date > LockDateTo.AddDays(14))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 14] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                    }



                   //Result 1-25
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                    if (TempDate.Date > LockDateTo.AddDays(19))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;
                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 17] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                    }



                    //End Of Month
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                    if (TempDate.Date > LockDateTo.AddDays(24))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 20] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }

                            }
                        }
                    }


                    //Before Month
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                    QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;


                    rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, false);
                    if (rsSum.RecordCount > 0)
                    {
                        adapter.Fill(dt, rsSum);
                        dt = Pivot(dt);
                        for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                        {
                            if (xlRangeLine.Cells[i, 4].Value != null)
                            {
                                Col += 4;

                                for (int y = 0; y < dt.Rows.Count; y++)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                    {
                                        for (int colX = 1; colX <= 2; colX++)
                                        {
                                            xlsSheet.Cells[(Col + colX) - 4, 7] = dt.Rows[y + colX][0];

                                        }

                                    }

                                }
                            }


                        }

                    }



                    //Nocome
                    findRang = xlsSheet.Range[xlsSheet.Cells[11, 2], xlsSheet.Cells[100, 2]].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                                 LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                    findRang = xlsSheet.Range[xlsSheet.Cells[11, 3], xlsSheet.Cells[findRang.Row, 3]].Find(What: strNocome, LookIn: Excel.XlFindLookIn.xlFormulas,
                               LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                    //Nocom=============================================================================================


                    strNumberSequence.Clear();


                    QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
   
                 
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["K" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                 

                    //Result 1-20
                    QuickSaleReportOBJ.DateTo =LockDateTo.AddDays(19);
                    if (TempDate.Date > LockDateTo.AddDays(14))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["N" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }


                    //LockDateTo.Date > TempDate.Date && LockDateTo.Date <= QuickSaleReportOBJ.DateTo
      
                    //Result 1-25
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                    if (TempDate.Date > LockDateTo.AddDays(19))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["Q" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }


                    //End Of Month
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                    if (TempDate.Date > LockDateTo.AddDays(24))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["T" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }


                    //Before Month
                    QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                    QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                    rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, false);
                    if (rsSum.RecordCount > 0)
                    {
                        xlsSheet.Range["G" + findRang.Row].CopyFromRecordset(rsSum);

                    }
                    //End Nocome

                }//end GMO




                //Trading ========================================================================================
                if (trading)
                {

                    xlsSheet = xlsBook.Sheets[5];
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;
                    xlRangeLine = xlsSheet.UsedRange;

                    intStartRow = 10;


                    QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 11] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                


                    //Result 1-20
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                    if (TempDate.Date > LockDateTo.AddDays(14))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 14] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                    }



     
                    //Result 1-25
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                    if (TempDate.Date > LockDateTo.AddDays(19))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 17] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                    }


                    //End Of Month
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                    if (TempDate.Date > LockDateTo.AddDays(24))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            adapter.Fill(dt, rsSum);
                            dt = Pivot(dt);
                            for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                            {
                                if (xlRangeLine.Cells[i, 4].Value != null)
                                {
                                    Col += 4;

                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            for (int colX = 1; colX <= 2; colX++)
                                            {
                                                xlsSheet.Cells[(Col + colX) - 4, 20] = dt.Rows[y + colX][0];

                                            }

                                        }

                                    }
                                }


                            }
                        }
                    }




                    //Before Month
                    dt.Clear();
                    Col = 11;
                    QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                    QuickSaleReportOBJ.DateTo =lastDayOfBeforeMonth;

                    rsSum = QuickSalesReportDAL.getQuickSalesReport(QuickSaleReportOBJ, strNumberSequence, trading);
                    if (rsSum.RecordCount > 0)
                    {
                        adapter.Fill(dt, rsSum);
                        dt = Pivot(dt);
                        for (int i = 11; i <= xlRangeLine.Rows.Count; i += 4)
                        {
                            if (xlRangeLine.Cells[i, 4].Value != null)
                            {
                                Col += 4;

                                for (int y = 0; y < dt.Rows.Count; y++)
                                {
                                    if (xlRangeLine.Cells[i, 4].Value2.ToString() == dt.Rows[y][0].ToString())
                                    {
                                        for (int colX = 1; colX <= 2; colX++)
                                        {
                                            xlsSheet.Cells[(Col + colX) - 4, 7] = dt.Rows[y + colX][0];

                                        }

                                    }

                                }
                            }


                        }
                    }

                    // Nocome Tranding

                    findRang = xlsSheet.Range[xlsSheet.Cells[11, 2], xlsSheet.Cells[100, 2]].Find(What: "GRAND TOTAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                                 LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                    findRang = xlsSheet.Range[xlsSheet.Cells[11, 3], xlsSheet.Cells[findRang.Row, 3]].Find(What: strNocome, LookIn: Excel.XlFindLookIn.xlFormulas,
                               LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);
                    //Nocom=============================================================================================


                    strNumberSequence.Clear();


                    QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(14);
  
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["K" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                


                    //Result 1-20
                   
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(19);
                    if (TempDate.Date > LockDateTo.AddDays(14))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["N" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }


                    //Result 1-25
                    //LockDateTo.Date > TempDate.Date && LockDateTo.Date <= QuickSaleReportOBJ.DateTo
     
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddDays(24);
                    if (TempDate.Date > LockDateTo.AddDays(19))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["Q" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }



                    //End Of Month

                    //LockDateTo.Date > TempDate.Date && LockDateTo.Date <= QuickSaleReportOBJ.DateTo
     
                    QuickSaleReportOBJ.DateTo = LockDateTo.AddMonths(1).AddDays(-1);
                    if (TempDate.Date > LockDateTo.AddDays(24))
                    {
                        rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["T" + findRang.Row].CopyFromRecordset(rsSum);

                        }
                    }


                    //Before Month
                    QuickSaleReportOBJ.DateFrom = firstDayBeforeMonth;
                    QuickSaleReportOBJ.DateTo = lastDayOfBeforeMonth;

                    rsSum = QuickSalesReportDAL.getQuickSalesReportNOCOME(QuickSaleReportOBJ, trading);
                    if (rsSum.RecordCount > 0)
                    {
                        xlsSheet.Range["G" + findRang.Row].CopyFromRecordset(rsSum);

                    }
                
                } //Trading


                      xlsSheet = xlsBook.Sheets[1];
                      QuickSaleReportOBJ.DateFrom = new DateTime(LockDateTo.AddMonths(0).Year, LockDateTo.AddMonths(0).Month, 1);
                      QuickSaleReportOBJ.DateTo = LockDateTo;
                      xlsSheet.Cells[8, 2] = QuickSaleReportOBJ.DateTo;

                //External Sale Adishima(PO and GMO)
 
                          xlsSheet.Range[xlsSheet.Cells[12, 6], xlsSheet.Cells[12, 8]].Formula =
                           String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}+'PO LENS'!F{2}", 12, 12,12);//3Q PLAN //Last Month //estimaate

                          xlsSheet.Range[xlsSheet.Cells[ 13, 6], xlsSheet.Cells[ 13, 8]].Formula =
                             String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}+'PO LENS'!F{2}", 13, 13,13);//3Q PLAN //Last Month //estimaate
                        
                            xlsSheet.Range[xlsSheet.Cells[12, 11], xlsSheet.Cells[12, 11]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}", 12, 12);//Result month 

                            xlsSheet.Range[xlsSheet.Cells[13, 11], xlsSheet.Cells[13, 11]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}+'PO LENS'!K{2}", 13, 13,13);//Result month 1-15

                            xlsSheet.Range[xlsSheet.Cells[12, 14], xlsSheet.Cells[12, 14]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}+'PO LENS'!N{2}", 12, 12,12);//Result month

                            xlsSheet.Range[xlsSheet.Cells[13,14], xlsSheet.Cells[13, 14]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}+'PO LENS'!N{2}", 13, 13,13);//Result month 1-20

                            xlsSheet.Range[xlsSheet.Cells[12, 17], xlsSheet.Cells[12, 17]].Formula =
                                          String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}+'PO LENS'!Q{2}", 12, 12,12);//Result month

                            xlsSheet.Range[xlsSheet.Cells[13, 17], xlsSheet.Cells[13, 17]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}+'PO LENS'!Q{2}", 13, 13,13);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[12, 20], xlsSheet.Cells[12, 20]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}+'PO LENS'!T{2}", 12, 12,12);//Result month

                            xlsSheet.Range[xlsSheet.Cells[13, 20], xlsSheet.Cells[13, 20]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}+'PO LENS'!T{2}", 13, 13,13);//Result month 1-30


/*
                //External Sale (Akishima)
                            xlsSheet.Range[xlsSheet.Cells[16, 6], xlsSheet.Cells[16, 8]].Formula =
                              String.Format(@"='RP FAC'!F{0}", 12);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[17, 6], xlsSheet.Cells[17, 8]].Formula =
                              String.Format(@"='RP FAC'!F{0}", 13);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[16, 11], xlsSheet.Cells[16, 11]].Formula =
                                     String.Format(@"='RP FAC'!K{0}", 12);//Result month 

                            xlsSheet.Range[xlsSheet.Cells[17, 11], xlsSheet.Cells[17, 11]].Formula =
                               String.Format(@"='RP FAC'!K{0}", 13);//Result month 1-15

                            xlsSheet.Range[xlsSheet.Cells[16, 14], xlsSheet.Cells[16, 14]].Formula =
                                          String.Format(@"='RP FAC'!N{0}", 12);//Result month

                            xlsSheet.Range[xlsSheet.Cells[17, 14], xlsSheet.Cells[17, 14]].Formula =
                               String.Format(@"='RP FAC'!N{0}", 13);//Result month 1-20

                            xlsSheet.Range[xlsSheet.Cells[16, 17], xlsSheet.Cells[16, 17]].Formula =
                                      String.Format(@"='RP FAC'!Q{0}", 12);//Result month

                            xlsSheet.Range[xlsSheet.Cells[17, 17], xlsSheet.Cells[17, 17]].Formula =
                                String.Format(@"='RP FAC'!Q{0}", 13);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[16, 20], xlsSheet.Cells[16, 20]].Formula =
                                String.Format(@"='RP FAC'!T{0}", 12);//Result month

                            xlsSheet.Range[xlsSheet.Cells[17, 20], xlsSheet.Cells[17, 20]].Formula =
                              String.Format(@"='RP FAC'!T{0}", 13);//Result month 1-30

*/


                    //External Sale (HOOP)
                            xlsSheet.Range[xlsSheet.Cells[16, 6], xlsSheet.Cells[16, 8]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!F{0}", 16);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[17, 6], xlsSheet.Cells[17, 8]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!F{0}", 17);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[16, 11], xlsSheet.Cells[16, 11]].Formula =
                                       String.Format(@"='GMO TOTAL FAC'!K{0}", 16);//Result month 

                            xlsSheet.Range[xlsSheet.Cells[17, 11], xlsSheet.Cells[17, 11]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!K{0}", 17);//Result month 1-15

                            xlsSheet.Range[xlsSheet.Cells[16, 14], xlsSheet.Cells[16, 14]].Formula =
                                        String.Format(@"='GMO TOTAL FAC'!N{0}", 16);//Result month

                            xlsSheet.Range[xlsSheet.Cells[17, 14], xlsSheet.Cells[17, 14]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!N{0}", 17);//Result month 1-20

                            xlsSheet.Range[xlsSheet.Cells[16, 17], xlsSheet.Cells[16, 17]].Formula =
                                     String.Format(@"='GMO TOTAL FAC'!Q{0}", 16);//Result month

                            xlsSheet.Range[xlsSheet.Cells[17, 17], xlsSheet.Cells[17, 17]].Formula =
                                 String.Format(@"='GMO TOTAL FAC'!Q{0}", 17);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[16, 20], xlsSheet.Cells[16, 20]].Formula =
                                 String.Format(@"='GMO TOTAL FAC'!T{0}", 16);//Result month

                            xlsSheet.Range[xlsSheet.Cells[17, 20], xlsSheet.Cells[17, 20]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!T{0}", 17);//Result month 1-30


                            //External Sale (HOPA)
                            xlsSheet.Range[xlsSheet.Cells[20, 6], xlsSheet.Cells[20, 8]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}+'PO LENS'!F{2}", 20,16,16);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[21, 6], xlsSheet.Cells[21, 8]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}+'PO LENS'!F{2}",21,17, 17);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[20, 11], xlsSheet.Cells[20, 11]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}+'PO LENS'!K{2}", 20,16,16);//Result month 

                            xlsSheet.Range[xlsSheet.Cells[21, 11], xlsSheet.Cells[21, 11]].Formula =
                                 String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}+'PO LENS'!K{2}", 21,17,17);//Result month 1-15

                            xlsSheet.Range[xlsSheet.Cells[20, 14], xlsSheet.Cells[20, 14]].Formula =
                                  String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}+'PO LENS'!N{2}",20,16, 16);//Result month

                            xlsSheet.Range[xlsSheet.Cells[21, 14], xlsSheet.Cells[21, 14]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}+'PO LENS'!N{2}", 21,17,17);//Result month 1-20

                            xlsSheet.Range[xlsSheet.Cells[20, 17], xlsSheet.Cells[20, 17]].Formula =
                                   String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}+'PO LENS'!Q{2}",20,16, 16);//Result month

                            xlsSheet.Range[xlsSheet.Cells[21, 17], xlsSheet.Cells[21, 17]].Formula =
                                 String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}+'PO LENS'!Q{2}",21,17, 17);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[20, 20], xlsSheet.Cells[20, 20]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}+'PO LENS'!T{2}",20,16, 16);//Result month

                            xlsSheet.Range[xlsSheet.Cells[21, 20], xlsSheet.Cells[21, 20]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}+'PO LENS'!T{2}",21,17, 17);//Result month 1-30





                            //External Sale (HOWT)
                            xlsSheet.Range[xlsSheet.Cells[24, 6], xlsSheet.Cells[24, 8]].Formula =
                              String.Format(@"='RP FAC'!F{0}", 20);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[25, 6], xlsSheet.Cells[25, 8]].Formula =
                                 String.Format(@"='RP FAC'!F{0}", 21);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[24, 11], xlsSheet.Cells[24, 11]].Formula =
                               String.Format(@"='RP FAC'!K{0}", 20);//Result month 

                            xlsSheet.Range[xlsSheet.Cells[25, 11], xlsSheet.Cells[25, 11]].Formula =
                                String.Format(@"='RP FAC'!K{0}", 21);//Result month 1-15

                            xlsSheet.Range[xlsSheet.Cells[24, 14], xlsSheet.Cells[24, 14]].Formula =
                                  String.Format(@"='RP FAC'!N{0}", 20);//Result month

                            xlsSheet.Range[xlsSheet.Cells[25, 14], xlsSheet.Cells[25, 14]].Formula =
                                String.Format(@"='RP FAC'!N{0}", 21);//Result month 1-20

                            xlsSheet.Range[xlsSheet.Cells[24, 17], xlsSheet.Cells[24, 17]].Formula =
                                 String.Format(@"='RP FAC'!Q{0}", 20);//Result month

                            xlsSheet.Range[xlsSheet.Cells[25, 17], xlsSheet.Cells[25, 17]].Formula =
                                String.Format(@"='RP FAC'!Q{0}", 21);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[24, 20], xlsSheet.Cells[24, 20]].Formula =
                               String.Format(@"='RP FAC'!T{0}", 20);//Result month

                            xlsSheet.Range[xlsSheet.Cells[25, 20], xlsSheet.Cells[25, 20]].Formula =
                               String.Format(@"='RP FAC'!T{0}", 21);//Result month 1-30



                            //External Sale (NIKON)
                            xlsSheet.Range[xlsSheet.Cells[28, 6], xlsSheet.Cells[28, 8]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}", 24,24);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[29, 6], xlsSheet.Cells[29, 8]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!F{0}+'RP FAC'!F{1}", 25,25);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[28, 11], xlsSheet.Cells[28, 11]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}", 24,24);//Result month 

                            xlsSheet.Range[xlsSheet.Cells[29, 11], xlsSheet.Cells[29, 11]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!K{0}+'RP FAC'!K{1}", 25,25);//Result month 1-15

                            xlsSheet.Range[xlsSheet.Cells[28, 14], xlsSheet.Cells[28, 14]].Formula =
                                 String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}", 24,24);//Result month

                            xlsSheet.Range[xlsSheet.Cells[29 ,14], xlsSheet.Cells[29, 14]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!N{0}+'RP FAC'!N{1}", 25,25);//Result month 1-20

                            xlsSheet.Range[xlsSheet.Cells[28, 17], xlsSheet.Cells[28, 17]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}", 24,24);//Result month

                            xlsSheet.Range[xlsSheet.Cells[29, 17], xlsSheet.Cells[29, 17]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!Q{0}+'RP FAC'!Q{1}", 25,25);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[28, 20], xlsSheet.Cells[28, 20]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}", 24,24);//Result month

                            xlsSheet.Range[xlsSheet.Cells[29, 20], xlsSheet.Cells[29, 20]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!T{0}+'RP FAC'!T{1}", 25,25);//Result month 1-30



                            //External Sale (RICOH MANUFACTURING)
                            xlsSheet.Range[xlsSheet.Cells[32, 6], xlsSheet.Cells[32, 8]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!F{0}", 28);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[33, 6], xlsSheet.Cells[33, 8]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!F{0}", 29);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[32, 11], xlsSheet.Cells[32, 11]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!K{0}", 28);//Result month 

                            xlsSheet.Range[xlsSheet.Cells[33, 11], xlsSheet.Cells[33, 11]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!K{0}", 29);//Result month 1-15

                            xlsSheet.Range[xlsSheet.Cells[32, 14], xlsSheet.Cells[32, 14]].Formula =
                                 String.Format(@"='GMO TOTAL FAC'!N{0}", 28);//Result month

                            xlsSheet.Range[xlsSheet.Cells[33, 14], xlsSheet.Cells[33, 14]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!N{0}", 29);//Result month 1-20

                            xlsSheet.Range[xlsSheet.Cells[32, 17], xlsSheet.Cells[32, 17]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!Q{0}", 28);//Result month

                            xlsSheet.Range[xlsSheet.Cells[33, 17], xlsSheet.Cells[33, 17]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!Q{0}", 29);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[32, 20], xlsSheet.Cells[32, 20]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!T{0}", 28);//Result month

                            xlsSheet.Range[xlsSheet.Cells[33, 20], xlsSheet.Cells[33, 20]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!T{0}", 29);//Result month 1-30



                            //External Sale (SONY THAI)
                            xlsSheet.Range[xlsSheet.Cells[36, 6], xlsSheet.Cells[36, 8]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!F{0}+'PO LENS'!F{1}", 32,20);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[37, 6], xlsSheet.Cells[37, 8]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!F{0}+'PO LENS'!F{1}", 33, 21);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[36, 11], xlsSheet.Cells[36, 11]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!K{0}+'PO LENS'!K{1}", 32, 20);//Result month 

                            xlsSheet.Range[xlsSheet.Cells[37, 11], xlsSheet.Cells[37, 11]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!K{0}+'PO LENS'!K{1}", 33, 21);//Result month 1-15

                            xlsSheet.Range[xlsSheet.Cells[36, 14], xlsSheet.Cells[36, 14]].Formula =
                                 String.Format(@"='GMO TOTAL FAC'!N{0}+'PO LENS'!N{1}", 32, 20);//Result month

                            xlsSheet.Range[xlsSheet.Cells[37, 14], xlsSheet.Cells[37, 14]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!N{0}+'PO LENS'!N{1}", 33, 21);//Result month 1-20

                            xlsSheet.Range[xlsSheet.Cells[36, 17], xlsSheet.Cells[36, 17]].Formula =
                                String.Format(@"='GMO TOTAL FAC'!Q{0}+'PO LENS'!Q{1}", 32, 20);//Result month

                            xlsSheet.Range[xlsSheet.Cells[37, 17], xlsSheet.Cells[37, 17]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!Q{0}+'PO LENS'!Q{1}", 33, 21);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[36, 20], xlsSheet.Cells[36, 20]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!T{0}+'PO LENS'!T{1}", 32, 20);//Result month

                            xlsSheet.Range[xlsSheet.Cells[37, 20], xlsSheet.Cells[37, 20]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!T{0}+'PO LENS'!T{1}", 33, 21);//Result month 1-30




                            //External Sale (PCS NO COME)
                            xlsSheet.Range[xlsSheet.Cells[39, 6], xlsSheet.Cells[39, 8]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!F{0}+'PO LENS'!F{1}+'RP FAC'!F{2}", 35, 39,31);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[39, 11], xlsSheet.Cells[39, 11]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!K{0}+'PO LENS'!K{1}+'RP FAC'!K{2}", 35, 39, 31);//Result month 1-15

                              xlsSheet.Range[xlsSheet.Cells[39, 14], xlsSheet.Cells[39, 14]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!N{0}+'PO LENS'!N{1}+'RP FAC'!N{2}", 35, 39, 31);//Result month 1-20

                              xlsSheet.Range[xlsSheet.Cells[39, 17], xlsSheet.Cells[39, 17]].Formula =
                             String.Format(@"='GMO TOTAL FAC'!Q{0}+'PO LENS'!Q{1}+'RP FAC'!Q{2}", 35, 39, 31);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[39, 20], xlsSheet.Cells[39, 20]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!T{0}+'PO LENS'!T{1}+'RP FAC'!T{2}", 35, 39, 31);//Result month 1-30


                            //External Sale (PARTIAL LENS)
                            xlsSheet.Range[xlsSheet.Cells[40, 6], xlsSheet.Cells[40, 8]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!F{0}", 36);//3Q PLAN //Last Month //estimaate

                            xlsSheet.Range[xlsSheet.Cells[40, 11], xlsSheet.Cells[40, 11]].Formula =
                               String.Format(@"='GMO TOTAL FAC'!K{0}", 36);//Result month 1-15

                            xlsSheet.Range[xlsSheet.Cells[40, 14], xlsSheet.Cells[40, 14]].Formula =
                            String.Format(@"='GMO TOTAL FAC'!N{0}", 36);//Result month 1-20

                            xlsSheet.Range[xlsSheet.Cells[40, 17], xlsSheet.Cells[40, 17]].Formula =
                           String.Format(@"='GMO TOTAL FAC'!Q{0}", 36);//Result month 1-25

                            xlsSheet.Range[xlsSheet.Cells[40, 20], xlsSheet.Cells[40, 20]].Formula =
                              String.Format(@"='GMO TOTAL FAC'!T{0}", 36);//Result month 1-30
                           




                   
             
               xlsApp.SheetsInNewWorkbook = 3;
               xlsApp.DisplayAlerts = true;
               xlsApp.Visible = true;
                
                rsSum = null;
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }// end getQuickSalesReportAll

        static DataTable  Pivot(DataTable tbl)
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
            /*
            var tblPivot = new DataTable();
            tblPivot.Columns.Add(tbl.Columns[0].ColumnName);
            for (int i = 1; i < tbl.Rows.Count; i++)
            {
                tblPivot.Columns.Add(Convert.ToString(i));
            }
            for (int col = 0; col < tbl.Columns.Count; col++)
            {
                var r = tblPivot.NewRow();
                r[0] = tbl.Columns[col].ToString();
                for (int j = 1; j < tbl.Rows.Count; j++)
                    r[j] = tbl.Rows[j][col];

                tblPivot.Rows.Add(r);
            }
            return tblPivot;*/
        }//end Pivot

        public string getQuickSalesReportSupport(QuickSaleReportOBJ QuickSaleReportOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();


              
                bool trading = false;
                int getDate;

                QuickSaleReportOBJ.DateTo = QuickSaleReportOBJ.DateFrom.AddMonths(1).AddDays(-1);
                getDate = QuickSaleReportOBJ.DateTo.Day;
                //Header

                rsSum = QuickSalesReportDAL.getQuickSalesSupport(QuickSaleReportOBJ, true, false, trading); //External no return


                if (rsSum.RecordCount > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();
                    Excel.Application xlsApp = new Excel.Application();
                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Range rangeSource, rangeDest;
                    Excel.Workbook xlsBookTemplate;

               
                    if (QuickSaleReportOBJ.strFactory == "GMO")
                    {
                       xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\QuickSales\SupportGMO.xls");
                  
                    
                    }
                    else if (QuickSaleReportOBJ.strFactory == "RP")
                    {
                        xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\QuickSales\SupportRP.xls");
                      
                    }
                    else
                    {
                        xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\QuickSales\SupportPO.xls");
                       
                    }

               
                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];

                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsSheet = xlsBook.Sheets[1];

                   
                     int intStartRow = 6;
                     int row = 0;

                    DataTable dtExternal = new DataTable();
                    DataTable dtInternal = new DataTable();
                    DataTable dtTrading = new DataTable();
                    DataTable dtNocome = new DataTable();


                    System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();
                    adapter.Fill(dtExternal, rsSum);
                    //dt = Pivot(dt);

                    xlsSheet.Cells[6, 1] = QuickSaleReportOBJ.DateFrom;

                    if (QuickSaleReportOBJ.Factory == "GMO")
                    {
                        //Rows
                        for (int sheet = 3; sheet >0; sheet--)
                        {
                            xlsSheet = xlsBook.Sheets[sheet];

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow + 1, 1], xlsSheet.Cells[intStartRow + 1, 34]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[intStartRow + 2, 1], xlsSheet.Cells[getDate + intStartRow, 34]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                            xlsSheet.Range[xlsSheet.Cells[getDate + 6, 1], xlsSheet.Cells[getDate + 7, 34]].EntireRow.Delete();


                           // xlsSheet.Range[xlsSheet.Cells[(((rsSum.RecordCount + 1) * 4) + 10) - 3, 1], xlsSheet.Cells[((rsSum.RecordCount + 1) * 4) + 10, 20]].EntireRow.Delete();

                            intStartRow = 6;
                            xlsSheet.Cells[6, 1] = QuickSaleReportOBJ.DateFrom;

                        }//end for sheets

                        xlsSheet = xlsBook.Sheets[2];

                    }
                    else
                    {


                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow + 1, 1], xlsSheet.Cells[intStartRow + 1, 34]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[intStartRow + 2, 1], xlsSheet.Cells[getDate + intStartRow, 34]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                        xlsSheet.Range[xlsSheet.Cells[getDate + 6, 1], xlsSheet.Cells[getDate + 7, 34]].EntireRow.Delete();

                    }//end if


                        Excel.Range xlRangeLine = xlsSheet.UsedRange;

                        //External
                        for (int loop = 0; loop < 2; loop++)
                        {
                            for (int i = 0; i < getDate; i++)
                            {
                                for (int y = 0; y < dtExternal.Rows.Count; y++)
                                {

                                    DateTime T1 = Convert.ToDateTime(xlRangeLine.Cells[i + intStartRow, 1].Value.ToString());
                                    DateTime T2 = Convert.ToDateTime(dtExternal.Rows[y][0].ToString());
                                    if (String.Format("{0:yyyy/MM/dd}", T1.Date) == String.Format("{0:yyyy/MM/dd}", T2.Date))
                                    {
                                        {
                                            if (loop == 0)
                                            {
                                                xlsSheet.Cells[i + intStartRow, 8] = dtExternal.Rows[y][1];
                                                xlsSheet.Cells[i + intStartRow, 9] = dtExternal.Rows[y][2];
                                                xlsSheet.Cells[i + intStartRow, 10] = dtExternal.Rows[y][3];

                                                xlsSheet.Cells[i + intStartRow, 11] = dtExternal.Rows[y][4];
                                                xlsSheet.Cells[i + intStartRow, 12] = dtExternal.Rows[y][5];
                                                xlsSheet.Cells[i + intStartRow, 13] = dtExternal.Rows[y][6];

                                                xlsSheet.Cells[i + intStartRow, 14] = dtExternal.Rows[y][7];
                                                xlsSheet.Cells[i + intStartRow, 15] = dtExternal.Rows[y][8];
                                                xlsSheet.Cells[i + intStartRow, 16] = dtExternal.Rows[y][9];

                                                xlsSheet.Cells[i + intStartRow, 17] = dtExternal.Rows[y][10]; //sony
                                                xlsSheet.Cells[i + intStartRow, 18] = dtExternal.Rows[y][11];

                                                xlsSheet.Cells[i + intStartRow, 19] = dtExternal.Rows[y][12];//ricoh
                                                xlsSheet.Cells[i + intStartRow, 20] = dtExternal.Rows[y][13];

                                                xlsSheet.Cells[i + intStartRow, 21] = dtExternal.Rows[y][14];//nicon
                                                xlsSheet.Cells[i + intStartRow, 22] = dtExternal.Rows[y][15];

                                                xlsSheet.Cells[i + intStartRow, 23] = dtExternal.Rows[y][16]; //nidec
                                                xlsSheet.Cells[i + intStartRow, 24] = dtExternal.Rows[y][17];
                                            }
                                            else
                                            {
                                                xlsSheet.Cells[i + intStartRow, 25] = dtExternal.Rows[y][1];
                                                xlsSheet.Cells[i + intStartRow, 26] = dtExternal.Rows[y][2];
                                                xlsSheet.Cells[i + intStartRow, 27] = dtExternal.Rows[y][3];

                                                xlsSheet.Cells[i + intStartRow, 28] = dtExternal.Rows[y][4];
                                                xlsSheet.Cells[i + intStartRow, 29] = dtExternal.Rows[y][5];
                                                xlsSheet.Cells[i + intStartRow, 30] = dtExternal.Rows[y][6];

                                                xlsSheet.Cells[i + intStartRow, 31] = dtExternal.Rows[y][7];
                                                xlsSheet.Cells[i + intStartRow, 32] = dtExternal.Rows[y][8];
                                                xlsSheet.Cells[i + intStartRow, 33] = dtExternal.Rows[y][9];

                                                xlsSheet.Cells[i + intStartRow, 34] = dtExternal.Rows[y][10];
                                                xlsSheet.Cells[i + intStartRow, 35] = dtExternal.Rows[y][11];


                                            }

                                        }

                                    }

                                }

                            }//end for external

                            rsSum = QuickSalesReportDAL.getQuickSalesSupport(QuickSaleReportOBJ, true, true, trading); //External  return
                            dtExternal.Clear();
                            adapter.Fill(dtExternal, rsSum);
                        }

                        //Internal
                        rsSum = QuickSalesReportDAL.getQuickSalesSupport(QuickSaleReportOBJ, false, false, trading); //Internal no return
                        adapter.Fill(dtInternal, rsSum);
                        for (int loop = 0; loop < 2; loop++)
                        {

                            if (rsSum.RecordCount > 0)
                            {

                                for (int i = 0; i < getDate; i++)
                                {
                                    for (int y = 0; y < dtInternal.Rows.Count; y++)
                                    {

                                        DateTime T1 = Convert.ToDateTime(xlRangeLine.Cells[i + intStartRow, 1].Value.ToString());
                                        DateTime T2 = Convert.ToDateTime(dtInternal.Rows[y][0].ToString());
                                        if (String.Format("{0:yyyy/MM/dd}", T1.Date) == String.Format("{0:yyyy/MM/dd}", T2.Date))
                                        {
                                            {

                                                if (loop == 0)
                                                {
                                                    xlsSheet.Cells[i + intStartRow, 2] = dtInternal.Rows[y][1];
                                                    xlsSheet.Cells[i + intStartRow, 3] = dtInternal.Rows[y][2];
                                                }
                                                else
                                                {
                                                    xlsSheet.Cells[i + intStartRow, 4] = dtInternal.Rows[y][1];
                                                    xlsSheet.Cells[i + intStartRow, 5] = dtInternal.Rows[y][2];

                                                }
                                            }

                                        }

                                    }

                                }//end for Internal
                            }
                            rsSum = QuickSalesReportDAL.getQuickSalesSupport(QuickSaleReportOBJ, false, true, trading); //Internal no return
                            dtInternal.Clear();
                            adapter.Fill(dtInternal, rsSum);
                        }//end loop




                        //NOcome
                        rsSum = QuickSalesReportDAL.getQuickSalesSupportNocome(QuickSaleReportOBJ,"NOtrading"); //Internal no return
                        adapter.Fill(dtNocome, rsSum);
                            for (int i = 0; i < getDate; i++)
                            {
                                for (int y = 0; y < dtNocome.Rows.Count; y++)
                                {
                                    DateTime T1 = Convert.ToDateTime(xlRangeLine.Cells[i + intStartRow, 1].Value.ToString());
                                    DateTime T2 = Convert.ToDateTime(dtNocome.Rows[y][0].ToString());
                                    if (String.Format("{0:yyyy/MM/dd}", T1.Date) == String.Format("{0:yyyy/MM/dd}", T2.Date))
                                    {
                                        {
                                            xlsSheet.Cells[i + intStartRow, 38] = dtNocome.Rows[y][1];
                                         
                                            }

                                        }

                              }

                        }// end for end nocome



                    //===============================================================SUM====================================================================================//
                        rsSum = QuickSalesReportDAL.getQuickSalesSupportSum(QuickSaleReportOBJ,false,"NOtrading",true); //External-Internal no return -com

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["AK" + 41].CopyFromRecordset(rsSum);
                        }
                        row = rsSum.RecordCount;
                        rsSum = QuickSalesReportDAL.getQuickSalesSupportSum(QuickSaleReportOBJ, true,"NOtrading",true); //External-Internal  return -com
                       
                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["AK" + (41+row)].CopyFromRecordset(rsSum);
                        }

                       //No come
                         row += (41+rsSum.RecordCount);
                         xlsSheet.Cells[row, 37] = "NO-Comercial";

                         xlsSheet.Range[xlsSheet.Cells[row, 39], xlsSheet.Cells[row, 39]].Formula = "=AL" + (getDate +7);


                       // row += rsSum.RecordCount;
                       // rsSum = QuickSalesReportDAL.getQuickSalesSupportSum(QuickSaleReportOBJ, false, "NOtrading", false); //External-Internal  return -com

                       // if (rsSum.RecordCount > 0)
                       // {
                       //     xlsSheet.Range["AK" + (41 + row)].CopyFromRecordset(rsSum);
                       // }





                        if (QuickSaleReportOBJ.Factory == "GMO")
                        {
                            trading = true;
                        }

                       
                        if (trading)
                        {
                            //dtExternal.Clear();
                            //dtInternal.Clear();
                            xlsSheet = xlsBook.Sheets[3];
                            rsSum = QuickSalesReportDAL.getQuickSalesSupport(QuickSaleReportOBJ, true, false, trading); //External Trading
                            adapter.Fill(dtTrading, rsSum);
                            //External Trading
                            for (int loop = 0; loop < 2; loop++)
                            {
                                for (int i = 0; i < getDate; i++)
                                {
                                    for (int y = 0; y < dtTrading.Rows.Count; y++)
                                    {

                                        DateTime T1 = Convert.ToDateTime(xlRangeLine.Cells[i + intStartRow, 1].Value.ToString());
                                        DateTime T2 = Convert.ToDateTime(dtTrading.Rows[y][0].ToString());
                                        if (String.Format("{0:yyyy/MM/dd}", T1.Date) == String.Format("{0:yyyy/MM/dd}", T2.Date))
                                        {
                                            {
                                                if (loop == 0)
                                                {
                                                    xlsSheet.Cells[i + intStartRow, 8] = dtTrading.Rows[y][1];
                                                    xlsSheet.Cells[i + intStartRow, 9] = dtTrading.Rows[y][2];
                                                    xlsSheet.Cells[i + intStartRow, 10] = dtTrading.Rows[y][3];

                                                    xlsSheet.Cells[i + intStartRow, 11] = dtTrading.Rows[y][4];
                                                    xlsSheet.Cells[i + intStartRow, 12] = dtTrading.Rows[y][5];
                                                    xlsSheet.Cells[i + intStartRow, 13] = dtTrading.Rows[y][6];

                                                    xlsSheet.Cells[i + intStartRow, 14] = dtTrading.Rows[y][7];
                                                    xlsSheet.Cells[i + intStartRow, 15] = dtTrading.Rows[y][8];
                                                    xlsSheet.Cells[i + intStartRow, 16] = dtTrading.Rows[y][9];

                                                    xlsSheet.Cells[i + intStartRow, 17] = dtTrading.Rows[y][10]; //sony
                                                    xlsSheet.Cells[i + intStartRow, 18] = dtTrading.Rows[y][11];

                                                    xlsSheet.Cells[i + intStartRow, 19] = dtTrading.Rows[y][12];//ricoh
                                                    xlsSheet.Cells[i + intStartRow, 20] = dtTrading.Rows[y][13];

                                                    xlsSheet.Cells[i + intStartRow, 21] = dtTrading.Rows[y][14];//nicon
                                                    xlsSheet.Cells[i + intStartRow, 22] = dtTrading.Rows[y][15];

                                                    xlsSheet.Cells[i + intStartRow, 23] = dtTrading.Rows[y][16]; //nidec
                                                    xlsSheet.Cells[i + intStartRow, 24] = dtTrading.Rows[y][17];
                                                }
                                                else
                                                {
                                                    xlsSheet.Cells[i + intStartRow, 25] = dtTrading.Rows[y][1];
                                                    xlsSheet.Cells[i + intStartRow, 26] = dtTrading.Rows[y][2];
                                                    xlsSheet.Cells[i + intStartRow, 27] = dtTrading.Rows[y][3];

                                                    xlsSheet.Cells[i + intStartRow, 28] = dtTrading.Rows[y][4];
                                                    xlsSheet.Cells[i + intStartRow, 29] = dtTrading.Rows[y][5];
                                                    xlsSheet.Cells[i + intStartRow, 30] = dtTrading.Rows[y][6];

                                                    xlsSheet.Cells[i + intStartRow, 31] = dtTrading.Rows[y][7];
                                                    xlsSheet.Cells[i + intStartRow, 32] = dtTrading.Rows[y][8];
                                                    xlsSheet.Cells[i + intStartRow, 33] = dtTrading.Rows[y][9];

                                                    xlsSheet.Cells[i + intStartRow, 34] = dtTrading.Rows[y][10];
                                                    xlsSheet.Cells[i + intStartRow, 35] = dtTrading.Rows[y][11];


                                                }

                                            }

                                        }

                                    }

                                }//end for external

                                rsSum = QuickSalesReportDAL.getQuickSalesSupport(QuickSaleReportOBJ, true, true, trading); //External  return
                                dtTrading.Clear();
                                adapter.Fill(dtTrading, rsSum);
                            }



                            //Trading Internal
                            dtTrading.Clear();
                            rsSum = QuickSalesReportDAL.getQuickSalesSupport(QuickSaleReportOBJ, false, false, trading); //Internal no return
                            adapter.Fill(dtTrading, rsSum);
                            for (int loop = 0; loop < 2; loop++)
                            {

                                if (rsSum.RecordCount > 0)
                                {

                                    for (int i = 0; i < getDate; i++)
                                    {
                                        for (int y = 0; y < dtInternal.Rows.Count; y++)
                                        {

                                            DateTime T1 = Convert.ToDateTime(xlRangeLine.Cells[i + intStartRow, 1].Value.ToString());
                                            DateTime T2 = Convert.ToDateTime(dtTrading.Rows[y][0].ToString());
                                            if (String.Format("{0:yyyy/MM/dd}", T1.Date) == String.Format("{0:yyyy/MM/dd}", T2.Date))
                                            {
                                                {

                                                    if (loop == 0)
                                                    {
                                                        xlsSheet.Cells[i + intStartRow, 2] = dtTrading.Rows[y][1];
                                                        xlsSheet.Cells[i + intStartRow, 3] = dtTrading.Rows[y][2];
                                                    }
                                                    else
                                                    {
                                                        xlsSheet.Cells[i + intStartRow, 4] = dtTrading.Rows[y][1];
                                                        xlsSheet.Cells[i + intStartRow, 5] = dtTrading.Rows[y][2];

                                                    }
                                                }

                                            }

                                        }

                                    }//end for Internal
                                }
                                rsSum = QuickSalesReportDAL.getQuickSalesSupport(QuickSaleReportOBJ, false, true, trading); //Internal no return
                                dtTrading.Clear();
                                adapter.Fill(dtTrading, rsSum);
                            }//end loop



                            //========================Sum
                            rsSum = QuickSalesReportDAL.getQuickSalesSupportSum(QuickSaleReportOBJ, false,"trading",true); //Internal no return

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["AK" + 41].CopyFromRecordset(rsSum);
                            }
                            row = rsSum.RecordCount;
                            rsSum = QuickSalesReportDAL.getQuickSalesSupportSum(QuickSaleReportOBJ, true,"trading",true); //Internal  return

                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["AK" + (41 + row)].CopyFromRecordset(rsSum);
                            }


                            //No come
                            row += (41 + rsSum.RecordCount);
                            xlsSheet.Cells[row, 37] = "NO-Comercial";

                            xlsSheet.Range[xlsSheet.Cells[row, 39], xlsSheet.Cells[row, 39]].Formula = "=AL" + (getDate + 7);

                           



                            //NOcome trading
                            dtNocome.Clear();
                            rsSum = QuickSalesReportDAL.getQuickSalesSupportNocome(QuickSaleReportOBJ, "trading"); //Internal no return
                            adapter.Fill(dtNocome, rsSum);
                            for (int i = 0; i < getDate; i++)
                            {
                                for (int y = 0; y < dtNocome.Rows.Count; y++)
                                {
                                    DateTime T1 = Convert.ToDateTime(xlRangeLine.Cells[i + intStartRow, 1].Value.ToString());
                                    DateTime T2 = Convert.ToDateTime(dtNocome.Rows[y][0].ToString());
                                    if (String.Format("{0:yyyy/MM/dd}", T1.Date) == String.Format("{0:yyyy/MM/dd}", T2.Date))
                                    {
                                        {
                                            xlsSheet.Cells[i + intStartRow, 38] = dtNocome.Rows[y][1];

                                        }

                                    }

                                }

                            }// end for end nocome

                   }
      
                  


                    xlsApp.SheetsInNewWorkbook = 3;
                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;
                }

                rsSum = null;
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }// end geSupporttQuickSalesReport


        


    }//end class
}
