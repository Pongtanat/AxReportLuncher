using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace NewVersion.Report.SalesReturn
{
    class SaleReturnBLL
    {
        SalesReturnDAL InvoiceReportDAL = new SalesReturnDAL();


        public string getNumberSequenceGroup(string strFac, int intShipmentLocation)
        {
            string strNumberSequenceGroup = "";
            DataTable dt = InvoiceReportDAL.getNumberSequenceGroup(strFac, intShipmentLocation);

            if (dt.Rows.Count > 0)
            {
                strNumberSequenceGroup = dt.Rows[0][0].ToString();

            }
            return strNumberSequenceGroup;
        }

        public DataTable getCustomerGroup()
        {
            return InvoiceReportDAL.getCustomerGroup();
        }


        public string getSaleReturnBook(SalesReturnOBJ SalesReturnOBJ)
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



                Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SalesReturn\SalesReturnBook.xls");
                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                if (!SalesReturnOBJ.ShowWH)
                {

                    rsSum = InvoiceReportDAL.getSalesReturnBook(SalesReturnOBJ, true, "");

                    if (rsSum.RecordCount > 0)
                    {
                    
                        xlsSheet = xlsBook.Sheets[1];

                        xlsSheet.Name = "Sales Return Book";
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;

                        xlsSheet.Cells[1, 1] = "Sales Return Book";
                        xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                        xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));
                       

                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                        xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 14]].Formula = "=(+M" + intStartRow + ")";
                        xlsSheet.Cells[rsSum.RecordCount + intStartRow, 9] = "Total";


                        xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 10], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 10]].Formula = "=SUM(J" + intStartRow + ":J" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 11], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 11]].Formula = "=SUM(K" + intStartRow + ":K" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 12], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 12]].Formula = "=SUM(L" + intStartRow + ":L" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 13], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 13]].Formula = "=SUM(M" + intStartRow + ":M" + (intStartRow + rsSum.RecordCount - 1) + ")";


                        xlsSheet.Range["A" + intStartRow, "R" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                        xlsSheet.Range["A1", "R" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$I1=" + (char)+34 + "Total" + (char)+34);
                        xlsSheet.Range["A1", "R" + (intStartRow + rsSum.RecordCount)].FormatConditions[1].Interior.Color = 14281213;

                        xlsSheet.Range["B:R"].Columns.EntireColumn.AutoFit();

                        //xlsApp.DisplayAlerts = true;
                       // xlsApp.Visible = true;
                    }
                }

                //===================================== F1==========================================//
                rsSum = InvoiceReportDAL.getSalesReturnBook(SalesReturnOBJ, true, "F1");
                if (rsSum.RecordCount > 0)
                {

                    intStartRow = 4;
                    xlsSheet = xlsBook.Sheets[2];

                    xlsSheet.Name = "Sales Return Book (F1)";
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlsSheet.Cells[1, 1] = "Sales Return Book";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));


                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                    xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 14]].Formula = "=(+M" + intStartRow + ")";
                    xlsSheet.Cells[rsSum.RecordCount + intStartRow, 9] = "Total";


                    xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 10], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 10]].Formula = "=SUM(J" + intStartRow + ":J" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 11], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 11]].Formula = "=SUM(K" + intStartRow + ":K" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 12], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 12]].Formula = "=SUM(L" + intStartRow + ":L" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 13], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 13]].Formula = "=SUM(M" + intStartRow + ":M" + (intStartRow + rsSum.RecordCount - 1) + ")";


                    xlsSheet.Range["A" + intStartRow, "R" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Range["A1", "R" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$I1=" + (char)+34 + "Total" + (char)+34);
                    xlsSheet.Range["A1", "R" + (intStartRow + rsSum.RecordCount)].FormatConditions[1].Interior.Color = 14281213;

                    xlsSheet.Range["B:R"].Columns.EntireColumn.AutoFit();
                }

                //===================================== F2==========================================//
                rsSum = InvoiceReportDAL.getSalesReturnBook(SalesReturnOBJ, true, "F2");
                if (rsSum.RecordCount > 0)
                {

                    intStartRow = 4;
                    xlsSheet = xlsBook.Sheets[3];

                    xlsSheet.Name = "Sales Return Book (F2)";
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlsSheet.Cells[1, 1] = "Sales Return Book";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));


                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                    xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 14]].Formula = "=(+M" + intStartRow + ")";
                    xlsSheet.Cells[rsSum.RecordCount + intStartRow, 9] = "Total";


                    xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 10], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 10]].Formula = "=SUM(J" + intStartRow + ":J" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 11], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 11]].Formula = "=SUM(K" + intStartRow + ":K" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 12], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 12]].Formula = "=SUM(L" + intStartRow + ":L" + (intStartRow + rsSum.RecordCount - 1) + ")";
                    xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 13], xlsSheet.Cells[intStartRow + rsSum.RecordCount, 13]].Formula = "=SUM(M" + intStartRow + ":M" + (intStartRow + rsSum.RecordCount - 1) + ")";


                    xlsSheet.Range["A" + intStartRow, "R" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Range["A1", "R" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$I1=" + (char)+34 + "Total" + (char)+34);
                    xlsSheet.Range["A1", "R" + (intStartRow + rsSum.RecordCount)].FormatConditions[1].Interior.Color = 14281213;

                    xlsSheet.Range["B:R"].Columns.EntireColumn.AutoFit();
                }

                if (SalesReturnOBJ.ShowWH)
                {
                    xlsBook.Sheets[1].delete();

                }

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

        public string getSaleReturnByItem(SalesReturnOBJ SalesReturnOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                ADODB.Recordset rsSumNOCom = new ADODB.Recordset();


                string strSystemPath = System.IO.Directory.GetCurrentDirectory();
                Excel.Application xlsApp = new Excel.Application();
                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                int intStartRow = 4;

                Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SalesReturn\SalesReturnByitem.xls");
                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                if (!SalesReturnOBJ.ShowWH)
                {

                    rsSum = InvoiceReportDAL.getSalesReturnByItem(SalesReturnOBJ, true, "");

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet = xlsBook.Sheets[1];

                        xlsSheet.Name = "Sales Return By Item";
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;

                        xlsSheet.Cells[1, 1] = "Sales Return By Item";
                        xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                        xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));
                       


                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "F" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                        xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$B1=" + (char)+34 + "Total" + (char)+34);
                        xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions[1].Interior.Color = 14281213;

                        xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$A1=" + (char)+34 + "Grand Total" + (char)+34);
                        xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions[2].Interior.Color = 14408946;

                        xlsSheet.Range["B:F"].Columns.EntireColumn.AutoFit();

                    }

                }


                //=============================================== F1 ==========================================//

                 rsSum = InvoiceReportDAL.getSalesReturnByItem(SalesReturnOBJ, true, "F1");
                 intStartRow = 4;

                 if (rsSum.RecordCount > 0)
                 {

                     xlsSheet = xlsBook.Sheets[2];

                     xlsSheet.Name = "Sales Return By Item (F1)";
                     xlsSheet.Cells.Font.Name = "Arial";
                     xlsSheet.Cells.Font.Size = 8;

                     xlsSheet.Cells[1, 1] = "Sales Return By Item";
                     xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                     xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));



                     xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                     xlsSheet.Range["A" + intStartRow, "F" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                     xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$B1=" + (char)+34 + "Total" + (char)+34);
                     xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions[1].Interior.Color = 14281213;

                     xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$A1=" + (char)+34 + "Grand Total" + (char)+34);
                     xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions[2].Interior.Color = 14408946;

                     xlsSheet.Range["B:F"].Columns.EntireColumn.AutoFit();

                 }

                     //================================================================ F2 ================================================//
                     rsSum = InvoiceReportDAL.getSalesReturnByItem(SalesReturnOBJ, true, "F2");
                     intStartRow = 4;
                 

                    if (rsSum.RecordCount > 0)
                    {

                        xlsSheet = xlsBook.Sheets[3];

                        xlsSheet.Name = "Sales Return By Item (F2)";
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;

                        xlsSheet.Cells[1, 1] = "Sales Return By Item";
                        xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                        xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));
                       


                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "F" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                        xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$B1=" + (char)+34 + "Total" + (char)+34);
                        xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions[1].Interior.Color = 14281213;

                        xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$A1=" + (char)+34 + "Grand Total" + (char)+34);
                        xlsSheet.Range["A1", "F" + (intStartRow + rsSum.RecordCount)].FormatConditions[2].Interior.Color = 14408946;

                        xlsSheet.Range["B:F"].Columns.EntireColumn.AutoFit();


                    }


                    if (SalesReturnOBJ.ShowWH)
                    {
                        xlsBook.Sheets[1].delete();

                    }

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

        public string getSaleReturnByCustomer(SalesReturnOBJ SalesReturnOBJ)
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
                Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SalesReturn\SalesReturnByCustomer.xls");
                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                if (!SalesReturnOBJ.ShowWH)
                {

                    rsSum = InvoiceReportDAL.getSaleReturnByCustomer(SalesReturnOBJ, true, "");

                    if (rsSum.RecordCount > 0)
                    {


                        xlsSheet = xlsBook.Sheets[1];

                        xlsSheet.Name = "Sales Return By Customer";
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;

                        xlsSheet.Cells[1, 1] = "Sales Return By Customer";
                        xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                        xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));



                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range["A" + intStartRow, "H" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                        xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$B1=" + (char)+34 + "Total" + (char)+34);
                        xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions[1].Interior.Color = 14281213;

                        xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$A1=" + (char)+34 + "Grand Total" + (char)+34);
                        xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions[2].Interior.Color = 14408946;

                        xlsSheet.Range["A:H"].Columns.EntireColumn.AutoFit();



                        //  xlsApp.DisplayAlerts = true;
                        //   xlsApp.Visible = true;
                    }
                }


                //==================================== F1 ========================================//
                rsSum = InvoiceReportDAL.getSaleReturnByCustomer(SalesReturnOBJ, true, "F1");
                intStartRow = 4;
                if (rsSum.RecordCount > 0)
                {
                    xlsSheet = xlsBook.Sheets[2];

                    xlsSheet.Name = "Sales Return By Customer (F1)";
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlsSheet.Cells[1, 1] = "Sales Return By Customer";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));



                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "H" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                    xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$B1=" + (char)+34 + "Total" + (char)+34);
                    xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions[1].Interior.Color = 14281213;

                    xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$A1=" + (char)+34 + "Grand Total" + (char)+34);
                    xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions[2].Interior.Color = 14408946;

                    xlsSheet.Range["A:H"].Columns.EntireColumn.AutoFit();

                }


                //====================================== F2 =====================================//

                rsSum = InvoiceReportDAL.getSaleReturnByCustomer(SalesReturnOBJ, true, "F2");
                intStartRow = 4;
                if (rsSum.RecordCount > 0)
                {
                    xlsSheet = xlsBook.Sheets[3];

                    xlsSheet.Name = "Sales Return By Customer (F2)";
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlsSheet.Cells[1, 1] = "Sales Return By Customer";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));



                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "H" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                    xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$B1=" + (char)+34 + "Total" + (char)+34);
                    xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions[1].Interior.Color = 14281213;

                    xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$A1=" + (char)+34 + "Grand Total" + (char)+34);
                    xlsSheet.Range["A1", "H" + (intStartRow + rsSum.RecordCount)].FormatConditions[2].Interior.Color = 14408946;

                    xlsSheet.Range["A:H"].Columns.EntireColumn.AutoFit();

                }

                if (SalesReturnOBJ.ShowWH)
                {
                    xlsBook.Sheets[1].delete();

                }

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

        public string getSalesReturnRemain(SalesReturnOBJ SalesReturnOBJ)
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
                Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SalesReturn\SalesReturnRemain.xls");
                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                if (!SalesReturnOBJ.ShowWH)
                {

                    rsSum = InvoiceReportDAL.getSalesReturnRemainReport(SalesReturnOBJ, true, "");

                    if (rsSum.RecordCount > 0)
                    {


                        xlsSheet = xlsBook.Sheets[1];

                        xlsSheet.Name = "Defective Receive & Remain";
                        xlsSheet.Cells.Font.Name = "Arial";
                        xlsSheet.Cells.Font.Size = 8;

                        xlsSheet.Cells[1, 1] = "Defective Receive & Remain";
                        xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                        xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));



                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range[xlsSheet.Cells[4, 15], xlsSheet.Cells[(rsSum.RecordCount + intStartRow - 1), 15]].FormulaR1C1 = "=IFERROR(R[0]C[-1]*R[0]C[-2],0)";
                        xlsSheet.Range[xlsSheet.Cells[4, 16], xlsSheet.Cells[(rsSum.RecordCount + intStartRow - 1), 16]].FormulaR1C1 = "=IFERROR(R[0]C[-6]-R[0]C[-2],0)";
                        xlsSheet.Range[xlsSheet.Cells[4, 17], xlsSheet.Cells[(rsSum.RecordCount + intStartRow - 1), 17]].FormulaR1C1 = "=IFERROR(R[0]C[-1]*R[0]C[-4],0)";

                        xlsSheet.Range["A" + intStartRow, "U" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                        xlsSheet.Range["B:U"].Columns.EntireColumn.AutoFit();



                        //  xlsApp.DisplayAlerts = true;
                        //   xlsApp.Visible = true;
                    }
                }


                //==================================== F1 ========================================//
                rsSum = InvoiceReportDAL.getSalesReturnRemainReport(SalesReturnOBJ, true, "F1");
                intStartRow = 4;
                if (rsSum.RecordCount > 0)
                {
                    xlsSheet = xlsBook.Sheets[2];

                    xlsSheet.Name = "Defective Receive & Remain (F1)";
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlsSheet.Cells[1, 1] = "Defective Receive & Remain";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));


                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[4, 15], xlsSheet.Cells[(rsSum.RecordCount + intStartRow - 1), 15]].FormulaR1C1 = "=IFERROR(R[0]C[-1]*R[0]C[-2],0)";
                    xlsSheet.Range[xlsSheet.Cells[4, 16], xlsSheet.Cells[(rsSum.RecordCount + intStartRow - 1), 16]].FormulaR1C1 = "=IFERROR(R[0]C[-6]-R[0]C[-2],0)";
                    xlsSheet.Range[xlsSheet.Cells[4, 17], xlsSheet.Cells[(rsSum.RecordCount + intStartRow - 1), 17]].FormulaR1C1 = "=IFERROR(R[0]C[-1]*R[0]C[-4],0)";

                    xlsSheet.Range["A" + intStartRow, "U" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Range["B:U"].Columns.EntireColumn.AutoFit();


                }


                //====================================== F2 =====================================//

                rsSum = InvoiceReportDAL.getSalesReturnRemainReport(SalesReturnOBJ, true, "F2");
                intStartRow = 4;
                if (rsSum.RecordCount > 0)
                {
                    xlsSheet = xlsBook.Sheets[3];

                    xlsSheet.Name = "Defective Receive & Remain (F2)";
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlsSheet.Cells[1, 1] = "Defective Receive & Remain";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} : {3}", SalesReturnOBJ.Factory, SalesReturnOBJ.DateFrom, SalesReturnOBJ.DateTo, SalesReturnOBJ.CustomerGroup.Replace("','", ", "));


                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[4, 15], xlsSheet.Cells[(rsSum.RecordCount + intStartRow - 1), 15]].FormulaR1C1 = "=IFERROR(R[0]C[-1]*R[0]C[-2],0)";
                    xlsSheet.Range[xlsSheet.Cells[4, 16], xlsSheet.Cells[(rsSum.RecordCount + intStartRow - 1), 16]].FormulaR1C1 = "=IFERROR(R[0]C[-6]-R[0]C[-2],0)";
                    xlsSheet.Range[xlsSheet.Cells[4, 17], xlsSheet.Cells[(rsSum.RecordCount + intStartRow - 1), 17]].FormulaR1C1 = "=IFERROR(R[0]C[-1]*R[0]C[-4],0)";
                   
                    xlsSheet.Range["A" + intStartRow, "U" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Range["B:U"].Columns.EntireColumn.AutoFit();
                }

                if (SalesReturnOBJ.ShowWH)
                {
                    xlsBook.Sheets[1].delete();

                }

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
    }//end class
}
