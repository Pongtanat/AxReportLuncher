using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace NewVersion.Report.SalesReturn.SummaryTransaction
{
    class ReturnTransactionBLL
    {
        ReturnTransactionDAL ReturnTransactionDAL = new ReturnTransactionDAL();

        public string ProcessBySection(ReturnTransactionOBJ ReturnTransactionOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                string strSystemPath = System.IO.Directory.GetCurrentDirectory();
                Excel.Application xlsApp = new Excel.Application();
                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                int intStartRow = 4;



                Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SummaryTransactionReturn\SummaryTransactionReturn.xls");
                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();
                xlsSheet = xlsBook.Sheets[1];

                string transType = "";
                if (ReturnTransactionOBJ.TransType == 3)
                {
                    transType = "Receive";
                }
                else
                {
                    transType = "Shipment";
                }


                if (ReturnTransactionOBJ.WareHouse=="All")
                {
                   

                    foreach (string cat in ReturnTransactionOBJ.Category.ToString().Split(','))
                    {
                        xlsSheet = xlsBook.Worksheets[1];
                        xlsSheet.Copy(After: xlsBook.Sheets[xlsBook.Sheets.Count]);
                        xlsSheet = xlsBook.Worksheets[xlsBook.Sheets.Count];
                        xlsSheet.Name = String.Format("{0} - {1}", cat, transType);
                        xlsSheet.Cells[1, 1] = String.Format("{0} -{1} ({2}) {3:dd/MM/yyyy} to {4:dd/MM/yyyy}", ReturnTransactionOBJ.WareHouse, ReturnTransactionOBJ.Factory, xlsSheet.Name, ReturnTransactionOBJ.DateFrom, ReturnTransactionOBJ.DateTo);
                       rsSum = ReturnTransactionDAL.getSummary2(ReturnTransactionOBJ,cat,"");

                    if (rsSum.RecordCount > 0)
                    {



                        xlsSheet.Range[xlsSheet.Cells[(4 + 1), 1], xlsSheet.Cells[4 + rsSum.RecordCount, 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                        xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "TOTAL" + (char)+34);
                        xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[1].Font.Bold = true;

                        xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "GRAND TOTAL" + (char)+34);
                        xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[2].Font.Bold = true;
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                        if (ReturnTransactionOBJ.TransType != 4)
                        {
                            xlsSheet.Range["C:C"].EntireColumn.Delete();
                            xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[4 + rsSum.RecordCount, 3]].NumberFormat = "dd/mm/yyyy";

                        }
                        else
                        {
                            xlsSheet.Range[xlsSheet.Cells[4, 4], xlsSheet.Cells[4 + rsSum.RecordCount, 4]].NumberFormat = "dd/mm/yyyy";
                        }

                        xlsSheet.Cells[(rsSum.RecordCount + 4), 6] = "TOTAL";
                        xlsSheet.Range["H"+(rsSum.RecordCount+4)+":J"+(rsSum.RecordCount+4)].Formula = "=SUM(H" + 4 + ":H" + (rsSum.RecordCount + 3) + ")";
                       
                        xlsSheet.Cells[(rsSum.RecordCount + 5), 6] = "GRAND TOTAL";
                        xlsSheet.Range["H" + (rsSum.RecordCount + 5) + ":J" + (rsSum.RecordCount + 5)].Formula = "=SUM(H" + (rsSum.RecordCount+ 4) + ":H" + (rsSum.RecordCount + 4) + ")";

                        xlsSheet.Range[xlsSheet.Cells[4, 10], xlsSheet.Cells[(rsSum.RecordCount + 5), 10]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";

                        xlsSheet.Range["B:J"].Columns.EntireColumn.AutoFit(); 
                    }
                }//split 


                    //========================================================== WH1 ================================================//
                }
                else if (ReturnTransactionOBJ.WareHouse == "WH1")
                {

                    foreach (string cat in ReturnTransactionOBJ.Category.ToString().Split(','))
                    {
                        xlsSheet = xlsBook.Worksheets[1];
                        xlsSheet.Copy(After: xlsBook.Sheets[xlsBook.Sheets.Count]);
                        xlsSheet = xlsBook.Worksheets[xlsBook.Sheets.Count];
                        xlsSheet.Name = String.Format("{0} - {1}", cat, transType);
                        xlsSheet.Cells[1, 1] = String.Format("{0} -{1} ({2}) {3:dd/MM/yyyy} to {4:dd/MM/yyyy}", ReturnTransactionOBJ.WareHouse, ReturnTransactionOBJ.Factory, xlsSheet.Name, ReturnTransactionOBJ.DateFrom, ReturnTransactionOBJ.DateTo);
                        rsSum = ReturnTransactionDAL.getSummary2(ReturnTransactionOBJ, cat,"F1");

                        if (rsSum.RecordCount > 0)
                        {



                            xlsSheet.Range[xlsSheet.Cells[(4 + 1), 1], xlsSheet.Cells[4 + rsSum.RecordCount, 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[1].Font.Bold = true;

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "GRAND TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[2].Font.Bold = true;
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                            if (ReturnTransactionOBJ.TransType != 4)
                            {
                                xlsSheet.Range["C:C"].EntireColumn.Delete();
                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[4 + rsSum.RecordCount, 3]].NumberFormat = "dd/mm/yyyy";

                            }
                            else
                            {
                                xlsSheet.Range[xlsSheet.Cells[4, 4], xlsSheet.Cells[4 + rsSum.RecordCount, 4]].NumberFormat = "dd/mm/yyyy";
                            }

                            xlsSheet.Cells[(rsSum.RecordCount + 4), 6] = "TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 4) + ":J" + (rsSum.RecordCount + 4)].Formula = "=SUM(H" + 4 + ":H" + (rsSum.RecordCount + 3) + ")";

                            xlsSheet.Cells[(rsSum.RecordCount + 5), 6] = "GRAND TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 5) + ":J" + (rsSum.RecordCount + 5)].Formula = "=SUM(H" + (rsSum.RecordCount + 4) + ":H" + (rsSum.RecordCount + 4) + ")";

                            xlsSheet.Range[xlsSheet.Cells[4, 10], xlsSheet.Cells[(rsSum.RecordCount + 5), 10]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";

                            xlsSheet.Range["B:J"].Columns.EntireColumn.AutoFit();
                        }
                    }//split 

                }
                    //========================================================= WH 2 =============================================================//

                

                else if (ReturnTransactionOBJ.WareHouse == "WH2")
                {

                    foreach (string cat in ReturnTransactionOBJ.Category.ToString().Split(','))
                    {
                        xlsSheet = xlsBook.Worksheets[1];
                        xlsSheet.Copy(After: xlsBook.Sheets[xlsBook.Sheets.Count]);
                        xlsSheet = xlsBook.Worksheets[xlsBook.Sheets.Count];
                        xlsSheet.Name = String.Format("{0} - {1}", cat, transType);
                        xlsSheet.Cells[1, 1] = String.Format("{0} -{1} ({2}) {3:dd/MM/yyyy} to {4:dd/MM/yyyy}", ReturnTransactionOBJ.WareHouse, ReturnTransactionOBJ.Factory, xlsSheet.Name, ReturnTransactionOBJ.DateFrom, ReturnTransactionOBJ.DateTo);
                        rsSum = ReturnTransactionDAL.getSummary2(ReturnTransactionOBJ, cat,"F2");

                        if (rsSum.RecordCount > 0)
                        {



                            xlsSheet.Range[xlsSheet.Cells[(4 + 1), 1], xlsSheet.Cells[4 + rsSum.RecordCount, 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[1].Font.Bold = true;

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "GRAND TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[2].Font.Bold = true;
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                            if (ReturnTransactionOBJ.TransType != 4)
                            {
                                xlsSheet.Range["C:C"].EntireColumn.Delete();
                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[4 + rsSum.RecordCount, 3]].NumberFormat = "dd/mm/yyyy";

                            }
                            else
                            {
                                xlsSheet.Range[xlsSheet.Cells[4, 4], xlsSheet.Cells[4 + rsSum.RecordCount, 4]].NumberFormat = "dd/mm/yyyy";
                            }

                            xlsSheet.Cells[(rsSum.RecordCount + 4), 6] = "TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 4) + ":J" + (rsSum.RecordCount + 4)].Formula = "=SUM(H" + 4 + ":H" + (rsSum.RecordCount + 3) + ")";

                            xlsSheet.Cells[(rsSum.RecordCount + 5), 6] = "GRAND TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 5) + ":J" + (rsSum.RecordCount + 5)].Formula = "=SUM(H" + (rsSum.RecordCount + 4) + ":H" + (rsSum.RecordCount + 4) + ")";

                            xlsSheet.Range[xlsSheet.Cells[4, 10], xlsSheet.Cells[(rsSum.RecordCount + 5), 10]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";

                            xlsSheet.Range["B:J"].Columns.EntireColumn.AutoFit();
                        }
                    }//split 

                }


                xlsSheet=xlsBook.Worksheets[1];
                xlsSheet.Delete();
                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;

                rsSum = null;
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        } /// End by Section
          /// 


        public string ProcessByVoucher(ReturnTransactionOBJ ReturnTransactionOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                string strSystemPath = System.IO.Directory.GetCurrentDirectory();
                Excel.Application xlsApp = new Excel.Application();
                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                int intStartRow = 4;



                Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SummaryTransactionReturn\SummaryTransactionReturn.xls");
                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();
                xlsSheet = xlsBook.Sheets[1];


                string transType = "";
                if (ReturnTransactionOBJ.TransType == 3)
                {
                    transType = "Receive";
                }
                else
                {
                    transType = "Shipment";
                }


                if (ReturnTransactionOBJ.WareHouse == "All")
                {


                    foreach (string cat in ReturnTransactionOBJ.Category.ToString().Split(','))
                    {
                        xlsSheet = xlsBook.Worksheets[1];
                        xlsSheet.Copy(After: xlsBook.Sheets[xlsBook.Sheets.Count]);
                        xlsSheet = xlsBook.Worksheets[xlsBook.Sheets.Count];
                        xlsSheet.Name = String.Format("{0} - {1}", cat, transType);
                        xlsSheet.Cells[1, 1] = String.Format("{0} -{1} ({2}) {3:dd/MM/yyyy} to {4:dd/MM/yyyy}", ReturnTransactionOBJ.WareHouse, ReturnTransactionOBJ.Factory, xlsSheet.Name, ReturnTransactionOBJ.DateFrom, ReturnTransactionOBJ.DateTo);
                        rsSum = ReturnTransactionDAL.getSummaryByVoucher(ReturnTransactionOBJ, cat, "");

                        if (rsSum.RecordCount > 0)
                        {



                            xlsSheet.Range[xlsSheet.Cells[(4 + 1), 1], xlsSheet.Cells[4 + rsSum.RecordCount, 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[1].Font.Bold = true;

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "GRAND TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[2].Font.Bold = true;
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                            if (ReturnTransactionOBJ.TransType != 4)
                            {
                                xlsSheet.Range["C:C"].EntireColumn.Delete();
                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[4 + rsSum.RecordCount, 3]].NumberFormat = "dd/mm/yyyy";

                            }
                            else
                            {
                                xlsSheet.Range[xlsSheet.Cells[4, 4], xlsSheet.Cells[4 + rsSum.RecordCount, 4]].NumberFormat = "dd/mm/yyyy";
                            }

                            xlsSheet.Cells[(rsSum.RecordCount + 4), 6] = "TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 4) + ":J" + (rsSum.RecordCount + 4)].Formula = "=SUM(H" + 4 + ":H" + (rsSum.RecordCount + 3) + ")";

                            xlsSheet.Cells[(rsSum.RecordCount + 5), 6] = "GRAND TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 5) + ":J" + (rsSum.RecordCount + 5)].Formula = "=SUM(H" + (rsSum.RecordCount + 4) + ":H" + (rsSum.RecordCount + 4) + ")";

                            //xlsSheet.Range[xlsSheet.Cells[4, 10], xlsSheet.Cells[(rsSum.RecordCount + 5), 10]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";

                            xlsSheet.Range["B:J"].Columns.EntireColumn.AutoFit();
                        }
                    }//split 


                    //========================================================== WH1 ================================================//
                }
                else if (ReturnTransactionOBJ.WareHouse == "WH1")
                {

                    foreach (string cat in ReturnTransactionOBJ.Category.ToString().Split(','))
                    {
                        xlsSheet = xlsBook.Worksheets[1];
                        xlsSheet.Copy(After: xlsBook.Sheets[xlsBook.Sheets.Count]);
                        xlsSheet = xlsBook.Worksheets[xlsBook.Sheets.Count];
                        xlsSheet.Name = String.Format("{0} - {1}", cat, transType);
                        xlsSheet.Cells[1, 1] = String.Format("{0} -{1} ({2}) {3:dd/MM/yyyy} to {4:dd/MM/yyyy}", ReturnTransactionOBJ.WareHouse, ReturnTransactionOBJ.Factory, xlsSheet.Name, ReturnTransactionOBJ.DateFrom, ReturnTransactionOBJ.DateTo);
                        rsSum = ReturnTransactionDAL.getSummaryByVoucher(ReturnTransactionOBJ, cat, "F1");

                        if (rsSum.RecordCount > 0)
                        {



                            xlsSheet.Range[xlsSheet.Cells[(4 + 1), 1], xlsSheet.Cells[4 + rsSum.RecordCount, 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[1].Font.Bold = true;

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "GRAND TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[2].Font.Bold = true;
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                            if (ReturnTransactionOBJ.TransType != 4)
                            {
                                xlsSheet.Range["C:C"].EntireColumn.Delete();
                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[4 + rsSum.RecordCount, 3]].NumberFormat = "dd/mm/yyyy";

                            }
                            else
                            {
                                xlsSheet.Range[xlsSheet.Cells[4, 4], xlsSheet.Cells[4 + rsSum.RecordCount, 4]].NumberFormat = "dd/mm/yyyy";
                            }

                            xlsSheet.Cells[(rsSum.RecordCount + 4), 6] = "TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 4) + ":J" + (rsSum.RecordCount + 4)].Formula = "=SUM(H" + 4 + ":H" + (rsSum.RecordCount + 3) + ")";

                            xlsSheet.Cells[(rsSum.RecordCount + 5), 6] = "GRAND TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 5) + ":J" + (rsSum.RecordCount + 5)].Formula = "=SUM(H" + (rsSum.RecordCount + 4) + ":H" + (rsSum.RecordCount + 4) + ")";

                           // xlsSheet.Range[xlsSheet.Cells[4, 10], xlsSheet.Cells[(rsSum.RecordCount + 5), 10]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";

                            xlsSheet.Range["B:J"].Columns.EntireColumn.AutoFit();
                        }
                    }//split 

                }
                //========================================================= WH 2 =============================================================//



                else if (ReturnTransactionOBJ.WareHouse == "WH2")
                {

                    foreach (string cat in ReturnTransactionOBJ.Category.ToString().Split(','))
                    {
                        xlsSheet = xlsBook.Worksheets[1];
                        xlsSheet.Copy(After: xlsBook.Sheets[xlsBook.Sheets.Count]);
                        xlsSheet = xlsBook.Worksheets[xlsBook.Sheets.Count];
                        xlsSheet.Name = String.Format("{0} - {1}", cat, transType);
                        xlsSheet.Cells[1, 1] = String.Format("{0} -{1} ({2}) {3:dd/MM/yyyy} to {4:dd/MM/yyyy}", ReturnTransactionOBJ.WareHouse, ReturnTransactionOBJ.Factory, xlsSheet.Name, ReturnTransactionOBJ.DateFrom, ReturnTransactionOBJ.DateTo);
                        rsSum = ReturnTransactionDAL.getSummaryByVoucher(ReturnTransactionOBJ, cat, "F2");

                        if (rsSum.RecordCount > 0)
                        {



                            xlsSheet.Range[xlsSheet.Cells[(4 + 1), 1], xlsSheet.Cells[4 + rsSum.RecordCount, 1]].EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[1].Font.Bold = true;

                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$F1=" + (char)+34 + "GRAND TOTAL" + (char)+34);
                            xlsSheet.Range["F1", "J" + (4 + rsSum.RecordCount)].FormatConditions[2].Font.Bold = true;
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                            if (ReturnTransactionOBJ.TransType != 4)
                            {
                                xlsSheet.Range["C:C"].EntireColumn.Delete();
                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[4 + rsSum.RecordCount, 3]].NumberFormat = "dd/mm/yyyy";

                            }
                            else
                            {
                                xlsSheet.Range[xlsSheet.Cells[4, 4], xlsSheet.Cells[4 + rsSum.RecordCount, 4]].NumberFormat = "dd/mm/yyyy";
                            }

                            xlsSheet.Cells[(rsSum.RecordCount + 4), 6] = "TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 4) + ":J" + (rsSum.RecordCount + 4)].Formula = "=SUM(H" + 4 + ":H" + (rsSum.RecordCount + 3) + ")";

                            xlsSheet.Cells[(rsSum.RecordCount + 5), 6] = "GRAND TOTAL";
                            xlsSheet.Range["H" + (rsSum.RecordCount + 5) + ":J" + (rsSum.RecordCount + 5)].Formula = "=SUM(H" + (rsSum.RecordCount + 4) + ":H" + (rsSum.RecordCount + 4) + ")";

                            //xlsSheet.Range[xlsSheet.Cells[4, 10], xlsSheet.Cells[(rsSum.RecordCount + 5), 10]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";

                            xlsSheet.Range["B:J"].Columns.EntireColumn.AutoFit();
                        }
                    }//split 

                }




                xlsSheet = xlsBook.Worksheets[1];
                xlsSheet.Delete();


                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;

                rsSum = null;
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }































        

        public DataTable getCategoryByType()
        {
            return ReturnTransactionDAL.getCategoryByType();
        }

        public DataTable getAllSectionByFactory(string strFactory)
        {
            return ReturnTransactionDAL.getAllSectionByFactory(strFactory);
        }

        public DataTable getAllSubSectionByFactory(string strFactory)
        {
            return ReturnTransactionDAL.getAllSubSectionByFactory(strFactory);
        }
         
    }
}
