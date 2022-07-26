using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace NewVersion.Report.MaterialReport
{
    class MaterialBLL
    {

        MaterialDAL MaterialDAL = new MaterialDAL();

        public string getNumberSequenceGroup(string strFac, int intShipmentLocation)
        {
            string strNumberSequenceGroup = "";
            DataTable dt = MaterialDAL.getNumberSequenceGroup(strFac, intShipmentLocation);

            if (dt.Rows.Count > 0)
            {
                strNumberSequenceGroup = dt.Rows[0][0].ToString();

            }
            return strNumberSequenceGroup;
        }

        public string getReceiveReport(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);

                rsSum = MaterialDAL.getMaterialReceive(MaterialOBJ); //Com

                if (rsSum.RecordCount > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 4;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate; 

                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialRecei.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Receive Report";
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlsSheet.Cells[1, 1] = "Receive Report";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);
                   
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "M" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

        
                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount-1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$E1=" + (char)+34 + "Total" + (char)+34);
                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[1].Interior.Color = 14281213;

                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$E1=" + (char)+34 + "Grand Total" + (char)+34);
                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[2].Interior.Color = 14408946;

                    xlsSheet.Range["B:M"].Columns.EntireColumn.AutoFit();

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }
       
        }//End Receive report

        public string getShiptmentReport(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);

                rsSum = MaterialDAL.getMaterialShipment(MaterialOBJ); //Com

                if (rsSum.RecordCount > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 4;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;

                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialShipment.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Shipment Report";
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlsSheet.Cells[1, 1] = "Shipment Report";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);

                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "O" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;



                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$G1=" + (char)+34 + "TOTAL" + (char)+34);
                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[1].Interior.Color = 14281213;

                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$G1=" + (char)+34 + "GRAND TOTAL" + (char)+34);
                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[2].Interior.Color = 14408946;

                    xlsSheet.Range["B:O"].Columns.EntireColumn.AutoFit();

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Shipment

        public string getDetailMaterialUSED(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);



                string[] arrGlassType = { "EB", "FC", "GB", "HS", "OTHER" };
            

                if (dtMonthRange.Rows.Count > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 5;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\DetailOfMaterialUSED.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Name = "Detail Of Material Report";
                    xlsSheet.Cells.Font.Name = "Arial";
                    int Column = 0;
                    int indexMonth = 5;

                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);

                    if (dtMonthRange.Rows.Count > 1)
                    {
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 5], xlsSheet.Cells[10, 7]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 5];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }


                        //Column
                
                        //xlsSheet.Range[xlsSheet.Cells[3, (8 + (dtMonthRange.Rows.Count * 3)) - 6], xlsSheet.Cells[10, (7 + (dtMonthRange.Rows.Count * 3)) - 3]].EntireColumn.delete();
                        xlsSheet.Range[xlsSheet.Cells[3, 5], xlsSheet.Cells[10, 10]].EntireColumn.delete();


                        Column = (8 + (dtMonthRange.Rows.Count * 3)) - ((8 + (dtMonthRange.Rows.Count * 3)) - 6);
                        Column = Column + 8;

                        xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 3) + 6] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1]);

                      //  xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 3) + 6] = "Compare " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];
                       

                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 8], xlsSheet.Cells[10, 14]].EntireColumn.delete();
                        Column = 10;
                    }

                 

                    for (int i = 0; i < arrGlassType.Length; i++)
                    {
                        if (arrGlassType[i].ToString() == "EB")
                        {
                            
                                rsSum = MaterialDAL.getDetailMaterialReport(MaterialOBJ, arrGlassType[i].ToString(), "know");

                                //Row
                                rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                                rangeSource.EntireRow.Copy();
                                rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                                rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                                intStartRow += rsSum.RecordCount + 1;



                            //================GD=================
                                rsSum = MaterialDAL.getDetailMaterialReport(MaterialOBJ, arrGlassType[i].ToString(), "GD");
                                //Row
                                rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                                rangeSource.EntireRow.Copy();
                                rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                                rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                                intStartRow += rsSum.RecordCount + 1;




                                //================RD=================
                                rsSum = MaterialDAL.getDetailMaterialReport(MaterialOBJ, arrGlassType[i].ToString(), "RD");
                                //Row
                                rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                                rangeSource.EntireRow.Copy();
                                rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                                rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                                xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                                intStartRow += rsSum.RecordCount + 1;

                            
                            
                        }
                        else
                        {
                            rsSum = MaterialDAL.getDetailMaterialReport(MaterialOBJ, arrGlassType[i].ToString(), "unknow");

                            //Row
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount + 1;

                        }

                    }//end loop glasstype

                    //===================== NG ====================
                    rsSum = MaterialDAL.getDetailMaterialReport(MaterialOBJ, "EB", "NG");

                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    intStartRow += rsSum.RecordCount + 1;



                    xlsSheet.Range[xlsSheet.Cells[(intStartRow) , 3], xlsSheet.Cells[(intStartRow + 2), Column]].EntireRow.delete();


                    int CostQty = 7;
                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[3, indexMonth] = drr[0];
                        indexMonth += 3;

                        xlsSheet.Range[xlsSheet.Cells[5, CostQty], xlsSheet.Cells[intStartRow, CostQty]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        CostQty += 3;
                    }


                    //Sheet Summary or material Used


                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Summary Of Material Report";
                    xlsSheet.Cells.Font.Name = "Arial";
                    Column = 0;
                    indexMonth = 3;
                    intStartRow = 5;

                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy} ", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);

                    if (dtMonthRange.Rows.Count > 1)
                    {


                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 5]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 3];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }


                        xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 8]].EntireColumn.delete();


                        Column = (6 + (dtMonthRange.Rows.Count * 3)) - ((6 + (dtMonthRange.Rows.Count * 3)) - 6);
                        Column = Column + 6;

                        xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 3) + 4] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1]);

                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 6], xlsSheet.Cells[10, 12]].EntireColumn.delete();
                        Column = 10;
                    }




                      rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                      if (rsSum.RecordCount > 0)
                      {


                          //Row
                          rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                          rangeSource.EntireRow.Copy();
                          rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                          rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                          xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                         


                          xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount ), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount+1), Column]].EntireRow.delete();



                          CostQty = 5;
                          intStartRow += rsSum.RecordCount;
                          foreach (DataRow drr in dtMonthRange.Rows)
                          {
                              xlsSheet.Cells[3, indexMonth] = drr[0];
                              indexMonth += 3;

                              xlsSheet.Range[xlsSheet.Cells[5, CostQty], xlsSheet.Cells[(intStartRow-1), CostQty]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                              CostQty += 3;
                          }

                      }

                      xlsSheet = xlsBook.Sheets[4];
                      xlsSheet.Delete();
                      xlsSheet = xlsBook.Sheets[3];
                      xlsSheet.Delete();

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end MaterialUSED

        public string getMaterialBalanceByItem(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
               
                rsSum = MaterialDAL.getMaterialBalanceByItem(MaterialOBJ); //Com

                if (rsSum.RecordCount > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 4;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;

                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialBalanceByItem.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Material Balance By Item";
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Cells.Font.Size = 8;

                    xlsSheet.Cells[1, 1] = "Material Balance By Item";
                    xlsSheet.Range[xlsSheet.Cells[1, 1], xlsSheet.Cells[1, 1]].Font.Size = 16;
                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);

                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[4, 9], xlsSheet.Cells[rsSum.RecordCount + intStartRow, 9]].Formula = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                    xlsSheet.Range["A" + intStartRow, "I" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=$A1=" + (char)+34 + "GRAND TOTAL" + (char)+34);
                    xlsSheet.Range["A1", "M" + (intStartRow + rsSum.RecordCount - 1)].FormatConditions[1].Interior.Color = 14281213;

        
                    xlsSheet.Range["B:I"].Columns.EntireColumn.AutoFit();

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end BalanceByItem

        public string getLossSuri(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);




                if (dtMonthRange.Rows.Count > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 5;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\Losssuri.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "LOSS SURI";
                    xlsSheet.Cells.Font.Name = "Arial";
                    int Column = 0;
                    int indexMonth = 3;

                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);

 
                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                        //xlsSheet.Range[xlsSheet.Cells[3, 6], xlsSheet.Cells[10, 12]].EntireColumn.delete();

                       
                        for (int i = 0; i < dtMonthRange.Rows.Count;i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 5]];
                            rangeSource.EntireColumn.Copy();

                            rangeDest = xlsSheet.Cells[3,6];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }
                   
                        xlsSheet.Range[xlsSheet.Cells[3, (6 + (dtMonthRange.Rows.Count * 3)) - 6], xlsSheet.Cells[10, (5 + (dtMonthRange.Rows.Count * 3))]].EntireColumn.delete();
                        Column = (6 + (dtMonthRange.Rows.Count * 3)) - ((6 + (dtMonthRange.Rows.Count * 3)) - 6);
                        Column = Column + 6;

                        //xlsSheet.Cells[3, 10] = "Compare " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 6], xlsSheet.Cells[10, 12]].EntireColumn.delete();
                        Column = 10;
                    }

                    rsSum = MaterialDAL.getLossSuri(MaterialOBJ);
                    if (rsSum.RecordCount > 0)
                    {
          
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount + 1;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow ), 3], xlsSheet.Cells[(intStartRow +2), Column]].EntireRow.delete();

                           // xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 4), Column]].EntireRow.delete();
                         int  Price = 5;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[3, indexMonth] = drr[0];
                            indexMonth += 3;

                            xlsSheet.Range[xlsSheet.Cells[5, Price], xlsSheet.Cells[intStartRow - 1, Price]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                            Price += 3;
                        }

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


        }//end Lossuri

        public string getMaterialReport(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                Excel.Range xlRangeLine;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);




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
                    Excel.Workbook xlsBookTemplate;
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialReports-RP.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();
                    DataTable dt = new DataTable();
                    System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();
               

                    int indexMonth = 0;
                    int intStartRow = 5;


                    xlsSheet = xlsBook.Sheets[3];
                    xlsSheet.Cells.Font.Name = "Arial";
                    indexMonth = 0;


                    if (dtMonthRange.Rows.Count > 0)
                    {
                        rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "USED", "Qty",false);
                        xlsSheet.Name = "Mat_Used_kg";

                        if (rsSum.RecordCount > 0)
                        {
                            

                            xlsSheet.Range["B" + 17].CopyFromRecordset(rsSum); //USED


                            rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "SALE", "Qty",false);//not SALE
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 25].CopyFromRecordset(rsSum);
                            }


                            rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "OTHER", "Qty",false);//not OTHER
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 33].CopyFromRecordset(rsSum);
                            }


                            //Purchase
                            rsSum = MaterialDAL.getMaterialPurchase(MaterialOBJ,"Qty");//not sales
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 5].CopyFromRecordset(rsSum);
                            }

                            //Balance
                            rsSum = MaterialDAL.getMaterialReportRPBalance(MaterialOBJ, "Qty");//not sales
                            xlRangeLine = xlsSheet.UsedRange;
                            intStartRow = 51;
                            if (rsSum.RecordCount > 0)
                            {
                                //xlsSheet.Range["B" + 51].CopyFromRecordset(rsSum);

                                adapter.Fill(dt, rsSum);
                                //dt = Pivot(dt);

                                for (int i = 0; i < 5; i++)
                                {
                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[intStartRow + i, 2].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            xlsSheet.Cells[intStartRow + i, 3] = dt.Rows[y][1];
                                        }
                                    }
                                }

                            }


                        }//end Rssum

                    }// end KG



                    //==============================YEN
                    xlsSheet = xlsBook.Sheets[4];
                    xlsSheet.Cells.Font.Name = "Arial";
                    indexMonth = 0;


                    if (dtMonthRange.Rows.Count > 0)
                    {
                        rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "USED", "Cost",true);
                        xlsSheet.Name = "Mat_Used_Yen";

                        if (rsSum.RecordCount > 0)
                        {
                            /*
                            if (dtMonthRange.Rows.Count > 1)
                            {

                                
                                //Column
                                for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                                {
                                    rangeSource = xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[4, 3]];
                                    rangeSource.EntireColumn.Copy();
                                    rangeDest = xlsSheet.Cells[4, 4];
                                    rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                                }

                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[66, 4]].EntireColumn.delete();

                            }
                            else
                            {
                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[66, 3]].EntireColumn.delete();


                            }


                            foreach (DataRow drr in dtMonthRange.Rows)
                            {
                                xlsSheet.Cells[4, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];
                                xlsSheet.Cells[14, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];
                               
                            }
                                */

                            xlsSheet.Range["B" + 15].CopyFromRecordset(rsSum); //USED

                        }//end Rssum

                    }// end YEN


                   



                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Cells.Font.Name = "Arial";

                    if (dtMonthRange.Rows.Count > 0)
                    {
                        rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "USED", "Cost", false);
                        xlsSheet.Name = "Mat_Used_Bath";

                        if (rsSum.RecordCount > 0)
                        {
                            /*
                            if (dtMonthRange.Rows.Count > 1)
                            {
                                //Column
                                for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                                {
                                    rangeSource = xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[4, 3]];
                                    rangeSource.EntireColumn.Copy();
                                    rangeDest = xlsSheet.Cells[4, 4];
                                    rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                                }

                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[66, 4]].EntireColumn.delete();

                            }
                            else
                            {
                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[66, 3]].EntireColumn.delete();


                            }

                            */
                            Excel.Range find  = xlsSheet.Range["A:A"].Find(What: "Material Piece", LookIn: Excel.XlFindLookIn.xlFormulas,
                        LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);


                            xlsSheet.Range[xlsSheet.Cells[(find.Row + 1), (3)], xlsSheet.Cells[(find.Row + 5), (dtMonthRange.Rows.Count + 2)]].Formula = String.Format(@"=IFERROR(C{0}/'Mat_Used_kg'!C{0},""-"")", 17, (17 + (dtMonthRange.Rows.Count-1)));


                            /*

                            foreach (DataRow drr in dtMonthRange.Rows)
                            {
                                xlsSheet.Cells[4, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];
                                xlsSheet.Cells[15, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];
                                xlsSheet.Cells[43, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];
                                xlsSheet.Cells[54, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];
                                xlsSheet.Cells[59, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];
                            }
                             * */

                            xlsSheet.Range["B" + 17].CopyFromRecordset(rsSum); //USED


                            rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "SALE", "Cost", false);//not SALE
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 25].CopyFromRecordset(rsSum);
                            }


                            rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "OTHER", "Cost", false);//not OTHER
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 33].CopyFromRecordset(rsSum);
                            }


                            //Purchase
                            // rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "OTHER", "OTHER");//not OTHER
                            //  if (rsSum.RecordCount > 0)
                            // {
                            //     xlsSheet.Range["B" + 60].CopyFromRecordset(rsSum);
                            // }

                            rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "SALE", "OTHER", false);//not OTHER
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 67].CopyFromRecordset(rsSum);
                            }



                            rsSum = MaterialDAL.getMaterialPurchase(MaterialOBJ, "Cost");//not sale
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 5].CopyFromRecordset(rsSum);


                            }



                            //Balance
                            rsSum = MaterialDAL.getMaterialReportRPBalance(MaterialOBJ, "Cost");//not sales
                            xlRangeLine = xlsSheet.UsedRange;
                            intStartRow = 44;
                            if (rsSum.RecordCount > 0)
                            {
                                //xlsSheet.Range["B" + 51].CopyFromRecordset(rsSum);
                                dt.Clear();
                                adapter.Fill(dt, rsSum);
                                //dt = Pivot(dt);

                                for (int i = 0; i < 5; i++)
                                {
                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[intStartRow + i, 2].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            xlsSheet.Cells[intStartRow + i, 3] = dt.Rows[y][2];
                                        }
                                    }
                                }

                            }







                        }//end RsSum

                    }// end ==============================


                    

                    //==============================Unit Price
                    xlsSheet = xlsBook.Sheets[5];
                    xlsSheet.Cells.Font.Name = "Arial";
                    indexMonth = 0;

                    /*
                    if (dtMonthRange.Rows.Count > 0)
                    {
                        rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "USED", "OTHER", false);
                        xlsSheet.Name = "Mat_Used_Unit_Price";

                        if (rsSum.RecordCount > 0)
                        {
                            
                            if (dtMonthRange.Rows.Count > 1)
                            {

                                
                                //Column
                                for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                                {
                                    rangeSource = xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[4, 3]];
                                    rangeSource.EntireColumn.Copy();
                                    rangeDest = xlsSheet.Cells[4, 4];
                                    rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                                }

                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[66, 4]].EntireColumn.delete();
                                

                            }
                            else
                            {
                                xlsSheet.Range[xlsSheet.Cells[4, 3], xlsSheet.Cells[66, 3]].EntireColumn.delete();

                        
                            }
                                

                            Excel.Range find = xlsSheet.Range["A:A"].Find(What: "TOTAL MATERIAL", LookIn: Excel.XlFindLookIn.xlFormulas,
                     LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                            xlsSheet.Range[xlsSheet.Cells[(find.Row), (3)], xlsSheet.Cells[(find.Row), (dtMonthRange.Rows.Count + 2)]].Formula = String.Format(@"=IFERROR('Mat_Used_Bath'!C{0}/'Mat_Used_kg'!C{1},""-"")", 10, 10);

                            xlsSheet.Range[xlsSheet.Cells[(find.Row + 10), (3)], xlsSheet.Cells[(find.Row + 10), (dtMonthRange.Rows.Count + 2)]].Formula = String.Format(@"=IFERROR('Mat_Used_Bath'!C{0}/'Mat_Used_kg'!C{1},""-"")", 22, 22);

                            xlsSheet.Range[xlsSheet.Cells[(find.Row + 18), (3)], xlsSheet.Cells[(find.Row + 18), (dtMonthRange.Rows.Count + 2)]].Formula = String.Format(@"=IFERROR('Mat_Used_Bath'!C{0}/'Mat_Used_kg'!C{1},""-"")", 30, 30);

                            xlsSheet.Range[xlsSheet.Cells[(find.Row + 26), (3)], xlsSheet.Cells[(find.Row + 26), (dtMonthRange.Rows.Count + 2)]].Formula = String.Format(@"=IFERROR('Mat_Used_Bath'!C{0}/'Mat_Used_kg'!C{1},""-"")", 38, 38);

                            xlsSheet.Range[xlsSheet.Cells[(find.Row + 35), (3)], xlsSheet.Cells[(find.Row + 35), (dtMonthRange.Rows.Count + 2)]].Formula = String.Format(@"=IFERROR('Mat_Used_Bath'!C{0}/'Mat_Used_kg'!C{1},""-"")", 49, 56);


                            
                            foreach (DataRow drr in dtMonthRange.Rows)
                            {
                                xlsSheet.Cells[4, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];
                                xlsSheet.Cells[13, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];
                                xlsSheet.Cells[39, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];

                            }
                            

                            xlsSheet.Range["B" + 15].CopyFromRecordset(rsSum); //USED * YEN


                            rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "SALE", "OTHER", false);//not SALE
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 23].CopyFromRecordset(rsSum);
                            }


                            rsSum = MaterialDAL.getMaterialReport(MaterialOBJ, "OTHER", "OTHER", false);//not OTHER
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 31].CopyFromRecordset(rsSum);
                            }


                            //Purchase
                            rsSum = MaterialDAL.getMaterialPurchase(MaterialOBJ, "OTHER");//not sales
                            if (rsSum.RecordCount > 0)
                            {
                                xlsSheet.Range["B" + 5].CopyFromRecordset(rsSum);
                            }



                            //Balance
                            rsSum = MaterialDAL.getMaterialReportRPBalance(MaterialOBJ, "Know");//not sales
                            xlRangeLine = xlsSheet.UsedRange;
                            intStartRow = 40;
                            if (rsSum.RecordCount > 0)
                            {
                                //xlsSheet.Range["B" + 51].CopyFromRecordset(rsSum);

                                adapter.Fill(dt, rsSum);
                                //dt = Pivot(dt);

                                for (int i = 0; i < 5; i++)
                                {
                                    for (int y = 0; y < dt.Rows.Count; y++)
                                    {
                                        if (xlRangeLine.Cells[intStartRow + i, 2].Value2.ToString() == dt.Rows[y][0].ToString())
                                        {
                                            xlsSheet.Cells[intStartRow + i, 3] = dt.Rows[y][1];
                                        }
                                    }
                                }

                            }



                    

                        }//end Rssum

                    }// end Unit Price



                    */

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end MaterialReport

        public string getSummaryMaterialBalance(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);


                  string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 5;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\SummaryMaterialBalance.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Cells.Font.Name = "Arial";
                    xlsSheet.Name = "Detail";

                  // MaterialOBJ.DateTo = MaterialOBJ.DateTo.AddMonths(1).AddDays(-1); // last days
                   rsSum=MaterialDAL.getSummaryMaterialBalance(MaterialOBJ);


                   if(rsSum.RecordCount>0){

                       xlsSheet.Cells[2, 3] = String.Format("{0:dd-MMM-yyyy}", MaterialOBJ.DateTo);

                       
                       //xlsSheet.Range[xlsSheet.Cells[5, 6], xlsSheet.Cells[intStartRow - 1, 6]].FormulaR1C1 = "=IFERROR(R[0]C[-1]*R[0]C[+1],0)";


                       rangeSource = xlsSheet.Range[xlsSheet.Cells[4, 1], xlsSheet.Cells[(4 + 1) , 7]];
                       rangeSource.EntireRow.Copy();

                      
                       rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow)+ rsSum.RecordCount + 3, 1], xlsSheet.Cells[(intStartRow) + (rsSum.RecordCount) + 5, 20]];

                       rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                       xlsSheet.Cells[(intStartRow) + rsSum.RecordCount + 2, 1] = "Goods In";
                       xlsSheet.Range[xlsSheet.Cells[(intStartRow) + rsSum.RecordCount + 4, 7], xlsSheet.Cells[(intStartRow) + rsSum.RecordCount + 4, 7]].Formula = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";


                       xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum); //put
                       xlsSheet.Range[xlsSheet.Cells[5, 7], xlsSheet.Cells[rsSum.RecordCount + (intStartRow-1), 7]].Formula = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";

                       xlsSheet.Range["A" + intStartRow, "G" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                       xlsSheet.Range["A:G"].Columns.EntireColumn.AutoFit();

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end getSummaryMaterialBalance

        public string getSummaryMaterialCompare(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataTable dt = new DataTable();
     

                string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                Excel.Application xlsApp = new Excel.Application();
                System.Globalization.CultureInfo oldCI;
                oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                int intStartRow = 4;


                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                Excel.Workbook xlsBookTemplate;
                xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialCompare.xls");

                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                xlsSheet = xlsBook.Sheets[2];
                xlsSheet.Cells.Font.Name = "Arial";
                //xlsSheet.Name = "Receive Report";

                // MaterialOBJ.DateTo = MaterialOBJ.DateTo.AddMonths(1).AddDays(-1); // last days
                rsSum = MaterialDAL.getMaterialReceive(MaterialOBJ);


                if (rsSum.RecordCount > 0)
                {

                    xlsSheet.Cells[1, 1] = String.Format("{0:dd-MMM-yyyy}", MaterialOBJ.DateTo);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "M" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Range["A:M"].Columns.EntireColumn.AutoFit();
                }//end rsSum.RecordCount


                xlsSheet.Range[xlsSheet.Cells[5, 4], xlsSheet.Cells[34, 4]].NumberFormat = "#,##0.00";

           

                xlsSheet = xlsBook.Sheets[3];
                xlsSheet.Cells.Font.Name = "Arial";
               // xlsSheet.Name = "Shipment Report";

                // MaterialOBJ.DateTo = MaterialOBJ.DateTo.AddMonths(1).AddDays(-1); // last days
                rsSum = MaterialDAL.getMaterialShipment(MaterialOBJ);


                if (rsSum.RecordCount > 0)
                {

                   xlsSheet.Cells[1, 1] = String.Format("{0:dd-MMM-yyyy}", MaterialOBJ.DateTo);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range["A" + intStartRow, "O" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Range["A:O"].Columns.EntireColumn.AutoFit();
                }//end rsSum.RecordCount


                xlsSheet = xlsBook.Sheets[1];
                intStartRow = 5;
                rsSum = MaterialDAL.BOM(MaterialOBJ);
                System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();
                adapter.Fill(dt, rsSum);

                for (int y = 0; y < dt.Rows.Count; y++)
                {
                    switch(dt.Rows[y][0].ToString()) {
                        case "EB":
                        xlsSheet.Cells[5, 2] = dt.Rows[y][1];
                        xlsSheet.Cells[6, 2] = dt.Rows[y][2];
                        break;

                        case "GB":
                        xlsSheet.Cells[9, 2] = dt.Rows[y][1];
                        xlsSheet.Cells[10, 2] = dt.Rows[y][2];
                        break;

                        case "FC":
                        xlsSheet.Cells[14, 2] = dt.Rows[y][1];
                        xlsSheet.Cells[15, 2] = dt.Rows[y][2];
                        break;

                        case "HS":
                        xlsSheet.Cells[19, 2] = dt.Rows[y][1];
                        xlsSheet.Cells[20, 2] = dt.Rows[y][2];
                        break;
                    

                    }
                   
                }

                xlsSheet = xlsBook.Sheets[1];
                xlsSheet.Range[xlsSheet.Cells[5, 4], xlsSheet.Cells[34, 4]].NumberFormat = "#,##0.00";
                xlsSheet.Range[xlsSheet.Cells[5, 6], xlsSheet.Cells[34, 6]].NumberFormat = "#,##0.00";



                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end getSummaryMaterialCompare

        public string getMaterialMoveMentByItem(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();

                string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                Excel.Application xlsApp = new Excel.Application();
                System.Globalization.CultureInfo oldCI;
                oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                int intStartRow = 4;


                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                Excel.Workbook xlsBookTemplate;

                if (MaterialOBJ.Factory == "GMO")
                {
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MoveMentMaterialByItem.xls");
                }
                else
                {
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MoveMentMaterialByItem2.xls");
                }

                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();
                Excel.Range rangeSource, rangeDest;

                xlsSheet = xlsBook.Sheets[2];
                xlsSheet.Cells.Font.Name = "Arial";
                xlsSheet.Name = "Movement By Item";
              
                

                rsSum = MaterialDAL.getMaterailMoveMentByItem(MaterialOBJ);


                if (rsSum.RecordCount > 0)
                {

                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.Factory, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);


                   // rsSum = MaterialDAL.getDetailMaterialReport(MaterialOBJ, arrGlassType[i].ToString());

                    //Row
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 24]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 24]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    //intStartRow += rsSum.RecordCount + 1;



                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);


                    if (MaterialOBJ.Factory == "PO")
                    {
                        xlsSheet.Range[xlsSheet.Cells[intStartRow, 21], xlsSheet.Cells[intStartRow + rsSum.RecordCount - 1, 21]].Formula = @"=SUM(E" + intStartRow + "+G" + intStartRow + "-I" + intStartRow + " -K" + intStartRow + "-M" + intStartRow + "-O" + intStartRow + "-Q" + intStartRow + "-S" + intStartRow + " )";
                        xlsSheet.Range[xlsSheet.Cells[intStartRow, 22], xlsSheet.Cells[intStartRow + rsSum.RecordCount - 1, 22]].Formula = @"=SUM(F" + intStartRow + "+H" + intStartRow + "-J" + intStartRow + " -L" + intStartRow + "-N" + intStartRow + "-P" + intStartRow + "-R" + intStartRow + "-T" + intStartRow + " )";
                    }
                    else
                    {

                        xlsSheet.Range[xlsSheet.Cells[intStartRow, 20], xlsSheet.Cells[intStartRow + rsSum.RecordCount - 1, 20]].Formula = @"=SUM(D" + intStartRow + "+F" + intStartRow + "-H" + intStartRow + " -J" + intStartRow + "-L" + intStartRow + "-N" + intStartRow + "-P" + intStartRow + "-R" + intStartRow + " )";
                        xlsSheet.Range[xlsSheet.Cells[intStartRow, 21], xlsSheet.Cells[intStartRow + rsSum.RecordCount - 1, 21]].Formula = @"=SUM(E" + intStartRow + "+G" + intStartRow + "-I" + intStartRow + " -K" + intStartRow + "-M" + intStartRow + "-O" + intStartRow + "-Q" + intStartRow + "-S" + intStartRow + " )";


               

                    }

                    // xlsSheet.Cells[intStartRow + rsSum.RecordCount, 3] = "TOTAL";
                    //xlsSheet.Range[xlsSheet.Cells[intStartRow + rsSum.RecordCount, 4], xlsSheet.Cells[intStartRow + rsSum.RecordCount,21]].Formula = @"=SUM(D" + intStartRow + ":D" + (intStartRow + rsSum.RecordCount - 1) + ")";

                    
                   // xlsSheet.Range["A" + intStartRow, "U" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;

   

                    xlsSheet.Range["A:X"].Columns.EntireColumn.AutoFit();
                }//end rsSum.RecordCount

                xlsSheet = xlsBook.Sheets[1];
                xlsSheet.Cells.Font.Name = "Arial";
                intStartRow += rsSum.RecordCount;
                rsSum = MaterialDAL.getMaterialWip(MaterialOBJ);

                if (rsSum.RecordCount > 0)
                {

                    xlsSheet.Cells[2, 1] = String.Format("{0} : {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.Factory, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);
                   // xlsSheet.Range[xlsSheet.Cells[6, 3], xlsSheet.Cells[6, 20]].Formula = String.Format("='Movement By Item'!D${0}", intStartRow);
                    
                    xlsSheet.Range["E" + 7].CopyFromRecordset(rsSum);

                      xlsSheet.Range["A:U"].Columns.EntireColumn.AutoFit();
                }//end
               
             

                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end getMaterialMoveMentByItem

        public string getGroupMateCompare(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);




                if (dtMonthRange.Rows.Count > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 5;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialGroupMateCompare.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Group mat compare";
                    xlsSheet.Cells.Font.Name = "Arial";
                    int Column = 0;
                    int indexMonth = 4;

                    xlsSheet.Cells[2, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);


                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                        //xlsSheet.Range[xlsSheet.Cells[3, 6], xlsSheet.Cells[10, 12]].EntireColumn.delete();


                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 4], xlsSheet.Cells[13, 6]];
                            rangeSource.EntireColumn.Copy();

                            rangeDest = xlsSheet.Cells[3, 4];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }
                 


                        xlsSheet.Range[xlsSheet.Cells[3, 4], xlsSheet.Cells[10,9]].EntireColumn.delete();
                        Column = (7 + (dtMonthRange.Rows.Count * 3)) - ((7 + (dtMonthRange.Rows.Count * 3)) - 7);
                        Column = Column + 9;

                        xlsSheet.Cells[3, (dtMonthRange.Rows.Count * 3)+5] = "DIFF " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[10, 13]].EntireColumn.delete();
                        Column = 6;
                    }




                    rsSum = MaterialDAL.getGroupmateCompare(MaterialOBJ,true);
                    if (rsSum.RecordCount > 0)
                    {

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        intStartRow += rsSum.RecordCount;

                        //xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 2), Column]].EntireRow.delete();

                    }

                    rsSum = MaterialDAL.getGroupmateCompare(MaterialOBJ, false);
                    if (rsSum.RecordCount > 0)
                    {

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        intStartRow += rsSum.RecordCount + 1;

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 2), Column]].EntireRow.delete();



                    }




                    int Price = 6;
                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[3, indexMonth] = drr[0];
                        indexMonth += 3;

                        xlsSheet.Range[xlsSheet.Cells[5, Price], xlsSheet.Cells[intStartRow - 1, Price]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        Price += 3;
                    }

                    xlsSheet.Range["A:C"].Columns.EntireColumn.AutoFit();

                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end GroupMateCompare

        public string getStockCompare(MaterialOBJ MaterialOBJ,string filePath)
        {
            try
            {


                    ADODB.Recordset rsSum = new ADODB.Recordset();
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 4;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\StockCompare.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[1];
                    //xlsSheet.Name = "Group mat compare";
                    xlsSheet.Cells.Font.Name = "Arial";

                    xlsSheet.Cells[1, 1] = MaterialOBJ.Factory;
                    xlsSheet.Cells[2, 1] = String.Format("Compare Stock {0}",DateTime.Now);
                    rsSum = MaterialDAL.getCompareStock(MaterialOBJ, filePath);


                    if (rsSum.RecordCount > 0)
                    {
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range[xlsSheet.Cells[4, 7], xlsSheet.Cells[intStartRow +rsSum.RecordCount, 7]].FormulaR1C1 = "=R[0]C[-1]-R[0]C[-2]";

                        xlsSheet.Range["A" + intStartRow, "G" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    }


                    xlsSheet.Range["A:G"].Columns.EntireColumn.AutoFit();
                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


         

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end StockCoompare

        public string getMaterialReportMO(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                Excel.Range xlRangeLine;

                DataTable dt = new DataTable();
                System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();
               
               int Column = 0;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);




                if (dtMonthRange.Rows.Count > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 5;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;
                   
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialReport-MO.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();
                    Excel.Range rangeSource, rangeDest;
                    int indexMonth = 0;
                    

                    xlsSheet = xlsBook.Sheets[3];
                    xlRangeLine = xlsSheet.UsedRange;
                    xlsSheet.Cells.Font.Name = "Arial";

                    xlsSheet.Name = "Mat_Used_PCS";


                    //TEST
                   //  rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "PER-PCS", "Purchase", false); 

                        //Baht Purchase Row
                        Column = 12;
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "Qty", "Purchase", false); //NOCOME Purchase
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }


                        // OPTICAL LENS  COME
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "Qty", "OPTICAL LENS", true); //COME

                        if (rsSum.RecordCount > 0)
                        {

                            xlsSheet.Range["B" + (intStartRow + 2)].CopyFromRecordset(rsSum);

                        }

                        //OPTICAL LENS NO COME
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "Qty", "OPTICAL LENS", false); //NOCOME

                        if (rsSum.RecordCount > 0)
                        {

                            xlsSheet.Range["B" + (intStartRow + 1)].CopyFromRecordset(rsSum);

                        }


                        //DEAD STOCK
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "Qty", "DEAD", false); //NOCOME

                        if (rsSum.RecordCount > 0)
                        {

                            xlsSheet.Range["B" + (intStartRow + 7)].CopyFromRecordset(rsSum);

                        }



                        //Baht USED==============================================================================
                        intStartRow += 12;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Qty", "Normal", false, "USED");
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }

                    /*
                        // OPTICAL LENS  COME
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Qty", "OPTICAL LENS", true, "know"); //COME

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["B" + (intStartRow + 2)].CopyFromRecordset(rsSum);
                        }

                        //OPTICAL LENS NO COME
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Qty", "OPTICAL LENS", false, "know"); //NOCOME

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["B" + (intStartRow + 1)].CopyFromRecordset(rsSum);
                        }
                    */

                        //NG
                        intStartRow += 10;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Qty", "Normal", true, "NG"); //NG
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }
                        else
                        {
                            intStartRow += 2;

                        }



                        //SALE
                        intStartRow += 5;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Qty", "Normal", true, "SALE"); //SALE
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }
                        else
                        {
                            intStartRow += 2;

                        }



                        //DEAD
                        intStartRow += 5;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Qty", "Normal", true, "DS"); //DS
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }
                        else
                        {
                            intStartRow += 2;

                        }


                        //Baht Balance===========================================================================
                        intStartRow += 6;
                        rsSum = MaterialDAL.getMaterialReportMOBalance(MaterialOBJ, "Qty", "Normal", false);
                        if (rsSum.RecordCount > 0)
                        {

                            adapter.Fill(dt, rsSum);
                            //dt = Pivot(dt);

                            for (int i = 0; i < 16; i++)
                            {
                                for (int y = 0; y < dt.Rows.Count; y++)
                                {
                                    if (xlRangeLine.Cells[intStartRow + i, 2].Value2.ToString() == dt.Rows[y][0].ToString())
                                    {
                                        xlsSheet.Cells[intStartRow + i, 3] = dt.Rows[y][1];
                                    }
                                }
                            }
                        }


                         

                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[4, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];

                        }
                    

                
                

  //==================================Mat USED PCS=====================================//

                    xlsSheet = xlsBook.Sheets[2];
                    xlRangeLine = xlsSheet.UsedRange;
                    xlsSheet.Cells.Font.Name = "Arial";
                    intStartRow = 5;
                    xlsSheet.Name = "Mat_Used_Baht";

                        //Baht Purchase Row
                        Column = 12;
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "Cost", "Purchase", false); //NOCOME Purchase
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }

                   
                        // OPTICAL LENS  COME
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "Cost", "OPTICAL LENS", true); //COME

                        if (rsSum.RecordCount > 0)
                        {

                            xlsSheet.Range["B" + (intStartRow + 2)].CopyFromRecordset(rsSum);

                        }

                        //OPTICAL LENS NO COME
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "Cost", "OPTICAL LENS", false); //NOCOME

                        if (rsSum.RecordCount > 0)
                        {

                            xlsSheet.Range["B" + (intStartRow + 1)].CopyFromRecordset(rsSum);

                        }
                  

                        //DEAD STOCK
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "Cost", "DEAD", false); //NOCOME

                        if (rsSum.RecordCount > 0)
                        {

                            xlsSheet.Range["B" + (intStartRow + 7)].CopyFromRecordset(rsSum);

                        }



                        //Baht USED==============================================================================
                        intStartRow += 12;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Cost", "Normal", false, "USED");
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }

                    /*
                        // OPTICAL LENS  COME
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Cost", "OPTICAL LENS", true, "know"); //COME

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["B" + (intStartRow + 2)].CopyFromRecordset(rsSum);
                        }

                        //OPTICAL LENS NO COME
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Cost", "OPTICAL LENS", false, "know"); //NOCOME

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["B" + (intStartRow + 1)].CopyFromRecordset(rsSum);
                        }
*/

                        //NG
                        intStartRow += 10;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Cost", "Normal", true, "NG"); //NG
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }
                        else
                        {
                            intStartRow += 2;

                        }



                        //SALE
                        intStartRow += 5;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Cost", "Normal", true, "SALE"); //SALE
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }
                        else
                        {
                            intStartRow += 2;

                        }



                        //DEAD
                        intStartRow += 5;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "Cost", "Normal", true, "DS"); //DS
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }
                        else
                        {
                            intStartRow += 2;

                        }


                        //Baht Balance===========================================================================
                        intStartRow += 6;
                        rsSum = MaterialDAL.getMaterialReportMOBalance(MaterialOBJ, "Cost", "Normal", false);
                        if (rsSum.RecordCount > 0)
                        {
                            dt.Clear();
                            adapter.Fill(dt, rsSum);
                            //dt = Pivot(dt);

                            for (int i = 0; i < 16; i++)
                            {
                                for (int y = 0; y < dt.Rows.Count; y++)
                                {
                                    if (xlRangeLine.Cells[intStartRow + i, 2].Value2.ToString() == dt.Rows[y][0].ToString())
                                    {
                                        xlsSheet.Cells[intStartRow + i, 3] = dt.Rows[y][2];
                                    }
                                }
                            }
                        }



                        indexMonth = 0;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[4, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];

                        }



                        //==================================PER PCS=====================================//

                        xlsSheet = xlsBook.Sheets[4];
                        xlRangeLine = xlsSheet.UsedRange;
                        xlsSheet.Cells.Font.Name = "Arial";
                        intStartRow = 5;
                        xlsSheet.Name = "PER PCS";

                        //Baht Purchase Row
                        Column = 12;
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "PER-PCS", "Purchase", false); //NOCOME Purchase
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }

               

                        // OPTICAL LENS  COME
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "PER-PCS", "OPTICAL LENS", true); //COME

                        if (rsSum.RecordCount > 0)
                        {

                            xlsSheet.Range["B" + (intStartRow + 2)].CopyFromRecordset(rsSum);

                        }

                        //OPTICAL LENS NO COME
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "PER-PCS", "OPTICAL LENS", false); //NOCOME

                        if (rsSum.RecordCount > 0)
                        {

                            xlsSheet.Range["B" + (intStartRow + 1)].CopyFromRecordset(rsSum);

                        }
                   

                        //DEAD STOCK
                        rsSum = MaterialDAL.getMaterialReportMOPurchase(MaterialOBJ, "PER-PCS", "DEAD", false); //NOCOME

                        if (rsSum.RecordCount > 0)
                        {

                            xlsSheet.Range["B" + (intStartRow + 8)].CopyFromRecordset(rsSum);

                        }



                        //Baht USED==============================================================================
                        intStartRow += 13;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "PER-PCS", "Normal", false, "USED");
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }

                    /*
                        // OPTICAL LENS  COME
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "PER-PCS", "OPTICAL LENS", true, "know"); //COME

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["B" + (intStartRow + 2)].CopyFromRecordset(rsSum);
                        }

                        //OPTICAL LENS NO COME
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "PER-PCS", "OPTICAL LENS", false, "know"); //NOCOME

                        if (rsSum.RecordCount > 0)
                        {
                            xlsSheet.Range["B" + (intStartRow + 1)].CopyFromRecordset(rsSum);
                        }
                    */

                        //NG
                        intStartRow += 11;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "PER-PCS", "Normal", true, "NG"); //NG
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }
                        else
                        {
                            intStartRow += 2;

                        }



                        //SALE
                        intStartRow += 5;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "PER-PCS", "Normal", true, "SALE"); //SALE
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }
                        else
                        {
                            intStartRow += 2;

                        }



                        //DEAD
                        intStartRow += 5;
                        rsSum = MaterialDAL.getMaterialReportMO(MaterialOBJ, "PER-PCS", "Normal", true, "DS"); //DS
                        if (rsSum.RecordCount > 0)
                        {

                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["B" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount;

                            xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        }
                        else
                        {
                            intStartRow += 2;

                        }


                        //Baht Balance===========================================================================
                   /*
                        intStartRow += 6;
                        rsSum = MaterialDAL.getMaterialReportMOBalance(MaterialOBJ, "PER-PCS", "Normal", false);
                        if (rsSum.RecordCount > 0)
                        {
                            dt.Clear();
                            adapter.Fill(dt, rsSum);
                            //dt = Pivot(dt);

                            for (int i = 0; i < 16; i++)
                            {
                                for (int y = 0; y < dt.Rows.Count; y++)
                                {
                                    if (xlRangeLine.Cells[intStartRow + i, 2].Value2.ToString() == dt.Rows[y][0].ToString())
                                    {
                                        xlsSheet.Cells[intStartRow + i, 3] = dt.Rows[y][3];
                                    }
                                }
                            }
                        }

                */

                        indexMonth = 0;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[4, (dtMonthRange.Rows.IndexOf(drr) + 3) + indexMonth] = drr[0];

                        }





                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;

                }
                return null;
            }

            catch (Exception ex)
            {
                return ex.Message;
            }
        
           
        }
       
        public string getSummaryMaterialMO(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataTable dt = new DataTable();


                string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                Excel.Application xlsApp = new Excel.Application();
                System.Globalization.CultureInfo oldCI;
                oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                int intStartRow = 4;
                Excel.Range rangeSource, rangeDest;

                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                Excel.Workbook xlsBookTemplate;
                xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\SummaryOfMaterialUsedMO.xls");

                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                xlsSheet = xlsBook.Sheets[2];
                xlsSheet.Cells.Font.Name = "Arial";
                //xlsSheet.Name = "Receive Report";

                rsSum = MaterialDAL.getMaterialReceive(MaterialOBJ);


                if (rsSum.RecordCount > 0)
                {
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 13]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), 13]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        intStartRow += rsSum.RecordCount + 1;

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 2), 13]].EntireRow.delete();


                    xlsSheet.Cells[1, 1] = String.Format("{0:dd-MMM-yyyy}", MaterialOBJ.DateTo);
                    //xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    //xlsSheet.Range["A" + intStartRow, "M" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Range["A:M"].Columns.EntireColumn.AutoFit();
                }//end rsSum.RecordCount



                intStartRow = 4;
                xlsSheet = xlsBook.Sheets[3];
                xlsSheet.Cells.Font.Name = "Arial";
                // xlsSheet.Name = "Shipment Report";

                // MaterialOBJ.DateTo = MaterialOBJ.DateTo.AddMonths(1).AddDays(-1); // last days
                rsSum = MaterialDAL.getMaterialShipmentMO(MaterialOBJ);


                if (rsSum.RecordCount > 0)
                {
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 9]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), 9]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    intStartRow += rsSum.RecordCount + 1;

                    xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 2), 9]].EntireRow.delete();



                    xlsSheet.Cells[1, 1] = String.Format("{0:dd-MMM-yyyy}", MaterialOBJ.DateTo);
                   // xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    //xlsSheet.Range["A" + intStartRow, "O" + (intStartRow + rsSum.RecordCount)].Borders.LineStyle = 1;
                    xlsSheet.Range["A:I"].Columns.EntireColumn.AutoFit();
                }//end rsSum.RecordCount


                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end Summary Of material MO

        public string getMaterialPurchase(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);




                if (dtMonthRange.Rows.Count > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 7;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialPurchase.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Name = "Detail Purchase";
                    xlsSheet.Cells.Font.Name = "Arial";
                    int Column = 0;
                    int indexMonth = 5;

                    xlsSheet.Cells[2, 1] = String.Format("SUMMARY MATERIAL PURCHASE YEAR {0:yyyy}",  MaterialOBJ.DateTo);


                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                         for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[7, 5], xlsSheet.Cells[11, 9]];
                            rangeSource.EntireColumn.Copy();

                            rangeDest = xlsSheet.Cells[7, 10];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[7, (9 + (dtMonthRange.Rows.Count * 5)) - 9], xlsSheet.Cells[11, (9 + (dtMonthRange.Rows.Count * 5))]].EntireColumn.delete();
                        Column = (9 + (dtMonthRange.Rows.Count * 5)) - ((9 + (dtMonthRange.Rows.Count * 5)) - 9);
                        Column = Column + 9;

                        //xlsSheet.Cells[3, 10] = "Compare " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[7, 10], xlsSheet.Cells[11, 14]].EntireColumn.delete();
                        Column = 14;
                    }

                    rsSum = MaterialDAL.getMaterailPurchaseYear(MaterialOBJ);
                    if (rsSum.RecordCount > 0)
                    {

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount-1 ), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        intStartRow += rsSum.RecordCount + 1;

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow), 1], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        // xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 4), Column]].EntireRow.delete();
                        int KG = 9;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[5, indexMonth] = drr[0];
                            indexMonth += 5;

                            xlsSheet.Range[xlsSheet.Cells[7, KG], xlsSheet.Cells[intStartRow - 3, KG]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-4],0)";
                            KG += 5;
                        }

                    }

                    xlsSheet.Range["B:CM"].Columns.EntireColumn.AutoFit();




                    //============== Summary====================
                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Summary";
                    xlsSheet.Cells.Font.Name = "Arial";
                     Column = 0;
                     indexMonth = 3;
                     intStartRow = 7;
                     xlsSheet.Cells[2, 1] = String.Format("SUMMARY MATERIAL PURCHASE YEAR {0:yyyy}", MaterialOBJ.DateTo);
 

                     if (dtMonthRange.Rows.Count > 1)
                     {
                         //Column
                         for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                         {
                             rangeSource = xlsSheet.Range[xlsSheet.Cells[7, 3], xlsSheet.Cells[11, 7]];
                             rangeSource.EntireColumn.Copy();

                             rangeDest = xlsSheet.Cells[7, 8];
                             rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                         }

                         xlsSheet.Range[xlsSheet.Cells[7, (11 + (dtMonthRange.Rows.Count * 5)+3) - 11], xlsSheet.Cells[11, (11 + (dtMonthRange.Rows.Count * 5)+1)]].EntireColumn.delete();
                         Column = (11 + (dtMonthRange.Rows.Count * 5)) - ((11 + (dtMonthRange.Rows.Count * 5)) - 11);
                         Column = Column + 8;

                         //xlsSheet.Cells[3, 10] = "Compare " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];
                     }
                     else
                     {
                         xlsSheet.Range[xlsSheet.Cells[7, 8], xlsSheet.Cells[11, 12]].EntireColumn.delete();
                         Column = 12;
                     }

                     rsSum = MaterialDAL.getMaterailPurchaseYearSummary(MaterialOBJ);
                     if (rsSum.RecordCount > 0)
                     {

                         rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                         rangeSource.EntireRow.Copy();
                         rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount - 1), Column]];
                         rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                         xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                         intStartRow += rsSum.RecordCount + 1;

                         xlsSheet.Range[xlsSheet.Cells[(intStartRow), 1], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                         // xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 4), Column]].EntireRow.delete();
                         int KG = 7;
                         foreach (DataRow drr in dtMonthRange.Rows)
                         {
                             xlsSheet.Cells[5, indexMonth] = drr[0];
                             indexMonth += 5;

                             xlsSheet.Range[xlsSheet.Cells[7, KG], xlsSheet.Cells[intStartRow - 2, KG]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-4],0)";
                             KG += 5;
                         }

                     }

                  

                    xlsSheet.Range["B:CM"].Columns.EntireColumn.AutoFit();


                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end MaterialPurchase

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

        public string getDetailMaterialUSEDForGMO(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);



                //  string[] arrGlassType = { "EB", "FC", "GB", "HS", "OTHER" };


                if (dtMonthRange.Rows.Count > 0)
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
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\DetailOfMaterialUSED.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[4];
                    xlsSheet.Name = "Detail Of Material Used";
                    xlsSheet.Cells.Font.Name = "Arial";
                    int Column = 0;
                    int indexMonth = 4;
                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND)LTD. - {0} FACTORY", MaterialOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("Detail of Materail Used", MaterialOBJ.Factory);
                    xlsSheet.Cells[3, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);





                    if (dtMonthRange.Rows.Count > 1)
                    {
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 4], xlsSheet.Cells[10, 6]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 4];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[3, 4], xlsSheet.Cells[10, 9]].EntireColumn.delete();


                        Column = (5 + (dtMonthRange.Rows.Count * 3));

                        // String.Format("{0:y}", dt);

                        xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);





                        //xlsSheet.Cells[3, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[10, 13]].EntireColumn.delete();
                        Column = 10;
                    }


                    //USED-USED-RT-RETRUN 1,6,3  #SALE 0 #DEAD 4 #NG-NG-RT 2,7 


                    string[] arrType = { "USED", "NG" };

                    for (int i = 0; i < arrType.Length; i++)
                    {

                        rsSum = MaterialDAL.getDetailMaterialReportForGMO(MaterialOBJ, arrType[i].ToString());

                        // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                        if (rsSum.RecordCount > 0)
                        {

                            //if (i == 3)
                            //  {
                            //     intStartRow += 4;
                            // }

                            //Row
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), Column]].EntireRow.delete();
                            intStartRow += rsSum.RecordCount + 2;
                        }
                        else
                        {
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 3), Column]].EntireRow.delete();

                        }


                    }// end for


                    intStartRow += 5;

                    rsSum = MaterialDAL.getDetailMaterialReportForGMO2(MaterialOBJ, "O");
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]].EntireRow.delete();
                    intStartRow += rsSum.RecordCount + 4;

                    rsSum = MaterialDAL.getDetailMaterialReportForGMO2(MaterialOBJ, "I");
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), Column]].EntireRow.delete();



                    intStartRow += rsSum.RecordCount + 3;

                    arrType = new[] { "SALE", "DEAD" };
                    for (int i = 0; i < arrType.Length; i++)
                    {

                        rsSum = MaterialDAL.getDetailMaterialReportForGMO(MaterialOBJ, arrType[i].ToString());

                        // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                        if (rsSum.RecordCount > 0)
                        {

                            //if (i == 3)
                            //  {
                            //     intStartRow += 4;
                            // }

                            //Row
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]].EntireRow.delete();
                            intStartRow += rsSum.RecordCount + 3;
                        }
                        else
                        {
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 3), Column]].EntireRow.delete();

                        }


                    }



                    int CostQty = 6;

                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[4, indexMonth] = drr[0];
                        indexMonth += 3;

                        xlsSheet.Range[xlsSheet.Cells[6, CostQty], xlsSheet.Cells[(intStartRow + rsSum.RecordCount - 1), CostQty]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        CostQty += 3;
                    }




                    //==================================== Summary Used =================================//

                    xlsSheet = xlsBook.Sheets[3];
                    xlsSheet.Name = "Summary Of Material Used";
                    xlsSheet.Cells.Font.Name = "Arial";
                    Column = 0;
                    indexMonth = 3;
                    intStartRow = 6;

                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND)LTD. - {0} FACTORY", MaterialOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("Summary of Materail Used", MaterialOBJ.Factory);
                    xlsSheet.Cells[3, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);

                    if (dtMonthRange.Rows.Count > 1)
                    {
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 5]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 3];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 8]].EntireColumn.delete();


                        Column = (4 + (dtMonthRange.Rows.Count * 3));

                        // xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1]);
                        xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 6], xlsSheet.Cells[10, 12]].EntireColumn.delete();
                        Column = 5;
                    }



                    // string[] arrType = { "USED", "SALE", "DEAD", "NG"};
                    arrType = new[] { "USED", "NG" };

                    for (int i = 0; i < arrType.Length; i++)
                    {
                        rsSum = MaterialDAL.getSummaryMaterialReportForGMO(MaterialOBJ, arrType[i].ToString());

                        // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                        if (rsSum.RecordCount > 0)
                        {
                            //Row
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                            intStartRow += rsSum.RecordCount;

                        }
                    }

                    rsSum = MaterialDAL.getSummaryMaterialReportForGMO2(MaterialOBJ, "O");
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    intStartRow += rsSum.RecordCount;


                    Excel.Range find;
                    find = xlsSheet.Range["A:A"].Find(What: "TOTAL MATERIAL USED", LookIn: Excel.XlFindLookIn.xlFormulas,
                                                                   LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                    // xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[(find.Row-1), Column]].EntireColumn.delete();
                    xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[(find.Row - 2), Column]].EntireRow.delete();



                    intStartRow = (find.Row + 1);
                    rsSum = MaterialDAL.getSummaryMaterialReportForGMO2(MaterialOBJ, "I");
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                    find = xlsSheet.Range["A:A"].Find(What: "TOTAL LENS PO FOR BALSUM", LookIn: Excel.XlFindLookIn.xlFormulas,
                                                                   LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);


                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(find.Row - 1), Column]].EntireRow.delete();

                    intStartRow += rsSum.RecordCount + 1;

                    arrType = new[] { "SALE", "DEAD" };

                    for (int i = 0; i < arrType.Length; i++)
                    {
                        rsSum = MaterialDAL.getSummaryMaterialReportForGMO(MaterialOBJ, arrType[i].ToString());

                        // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                        if (rsSum.RecordCount > 0)
                        {
                            //Row
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                            intStartRow += rsSum.RecordCount;

                        }
                    }


                    find = xlsSheet.Range["A:A"].Find(What: "TOTAL SALE AND DEAD", LookIn: Excel.XlFindLookIn.xlFormulas,
                                                                  LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                    xlsSheet.Range[xlsSheet.Cells[(intStartRow), 1], xlsSheet.Cells[(find.Row - 1), Column]].EntireRow.delete();



                    CostQty = 5;

                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[4, indexMonth] = drr[0];
                        indexMonth += 3;

                        xlsSheet.Range[xlsSheet.Cells[6, CostQty], xlsSheet.Cells[(find.Row), CostQty]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        CostQty += 3;
                    }

                    xlsSheet = xlsBook.Sheets[6];
                    xlsSheet.Delete();
                    xlsSheet = xlsBook.Sheets[5];
                    xlsSheet.Delete();
                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Delete();
                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Delete();
                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end MaterialUSEDForGMO


/*
        public string getDetailMaterialUSEDForGMO(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);



              //  string[] arrGlassType = { "EB", "FC", "GB", "HS", "OTHER" };


                if (dtMonthRange.Rows.Count > 0)
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
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\DetailOfMaterialUSED.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[4];
                    xlsSheet.Name = "Detail Of Material Used";
                    xlsSheet.Cells.Font.Name = "Arial";
                    int Column = 0;
                    int indexMonth = 4;
                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND)LTD. - {0} FACTORY", MaterialOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("Detail of Materail Used", MaterialOBJ.Factory);
                    xlsSheet.Cells[3, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);
                




                    if (dtMonthRange.Rows.Count > 1)
                    {
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 4], xlsSheet.Cells[10, 6]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 4];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[3, 4], xlsSheet.Cells[10, 9]].EntireColumn.delete();


                        Column = (5+ (dtMonthRange.Rows.Count * 3));

                       // String.Format("{0:y}", dt);

 xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);


                       

                       
                        //xlsSheet.Cells[3, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);
                    }
                    else{
                        xlsSheet.Range[xlsSheet.Cells[3, 7], xlsSheet.Cells[10, 13]].EntireColumn.delete();
                        Column = 10;
                    }


                    //USED-USED-RT-RETRUN 1,6,3  #SALE 0 #DEAD 4 #NG-NG-RT 2,7 
                    

                    string[] arrType = { "USED", "NG"};

                    for (int i = 0; i < arrType.Length; i++)
                    {

                    rsSum = MaterialDAL.getDetailMaterialReportForGMO(MaterialOBJ,arrType[i].ToString());

                   // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                    if (rsSum.RecordCount > 0)
                    {

                        //if (i == 3)
                      //  {
                       //     intStartRow += 4;
                       // }

                        //Row
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), Column]].EntireRow.delete();
                        intStartRow += rsSum.RecordCount + 2;
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 3), Column]].EntireRow.delete();

                    }


                    }// end for


                    intStartRow += 5;

                    rsSum = MaterialDAL.getDetailMaterialReportForGMO2(MaterialOBJ, "O");
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]].EntireRow.delete();
                    intStartRow += rsSum.RecordCount+4;

                    rsSum = MaterialDAL.getDetailMaterialReportForGMO2(MaterialOBJ, "I");
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), Column]].EntireRow.delete();



                    intStartRow += rsSum.RecordCount + 3;

                    arrType = new [] { "SALE", "DEAD"};
                    for (int i = 0; i < arrType.Length; i++)
                    {

                        rsSum = MaterialDAL.getDetailMaterialReportForGMO(MaterialOBJ, arrType[i].ToString());

                        // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                        if (rsSum.RecordCount > 0)
                        {

                            //if (i == 3)
                            //  {
                            //     intStartRow += 4;
                            // }

                            //Row
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]].EntireRow.delete();
                            intStartRow += rsSum.RecordCount + 3;
                        }
                        else
                        {
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 3), Column]].EntireRow.delete();

                        }


                    }



                    int CostQty = 6;
   
                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[4, indexMonth] = drr[0];
                        indexMonth += 3;

                        xlsSheet.Range[xlsSheet.Cells[6, CostQty], xlsSheet.Cells[(intStartRow + rsSum.RecordCount-1), CostQty]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        CostQty += 3;
                    }




                    //==================================== Summary Used =================================//

                    xlsSheet = xlsBook.Sheets[3];
                    xlsSheet.Name = "Summary Of Material Used";
                    xlsSheet.Cells.Font.Name = "Arial";
                   Column = 0;
                   indexMonth = 3;
                   intStartRow = 6;

                   xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND)LTD. - {0} FACTORY", MaterialOBJ.Factory);
                   xlsSheet.Cells[2, 1] = String.Format("Summary of Materail Used", MaterialOBJ.Factory);
                   xlsSheet.Cells[3, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);

                    if (dtMonthRange.Rows.Count > 1)
                    {
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 5]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 3];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 8]].EntireColumn.delete();


                        Column = (4 + (dtMonthRange.Rows.Count * 3));

                       // xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1]);
                        xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 6], xlsSheet.Cells[10, 12]].EntireColumn.delete();
                        Column = 5;
                    }


                    
                   // string[] arrType = { "USED", "SALE", "DEAD", "NG"};
                    arrType = new[] { "USED", "NG" };

                    for (int i = 0; i < arrType.Length; i++)
                    {
                    rsSum = MaterialDAL.getSummaryMaterialReportForGMO(MaterialOBJ,arrType[i].ToString());

                    // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                    if (rsSum.RecordCount > 0)
                    {
                        //Row
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                        intStartRow += rsSum.RecordCount;

                    }
                   }

                    rsSum = MaterialDAL.getSummaryMaterialReportForGMO2(MaterialOBJ,"O");
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    intStartRow += rsSum.RecordCount;


                    Excel.Range find;
                    find = xlsSheet.Range["A:A"].Find(What: "TOTAL MATERIAL USED", LookIn: Excel.XlFindLookIn.xlFormulas,
                                                                   LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                   // xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[(find.Row-1), Column]].EntireColumn.delete();
                    xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[(find.Row - 2), Column]].EntireRow.delete();



                    intStartRow = (find.Row + 1);
                    rsSum = MaterialDAL.getSummaryMaterialReportForGMO2(MaterialOBJ,"I");
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                    find = xlsSheet.Range["A:A"].Find(What: "TOTAL LENS PO FOR BALSUM", LookIn: Excel.XlFindLookIn.xlFormulas,
                                                                   LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);


                    xlsSheet.Range[xlsSheet.Cells[(intStartRow+rsSum.RecordCount), 1], xlsSheet.Cells[(find.Row - 1), Column]].EntireRow.delete();

                    intStartRow += rsSum.RecordCount + 1;

                    arrType = new[] { "SALE", "DEAD" };

                    for (int i = 0; i < arrType.Length; i++)
                    {
                        rsSum = MaterialDAL.getSummaryMaterialReportForGMO(MaterialOBJ, arrType[i].ToString());

                        // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                        if (rsSum.RecordCount > 0)
                        {
                            //Row
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                            intStartRow += rsSum.RecordCount;

                        }
                    }


                    find = xlsSheet.Range["A:A"].Find(What: "TOTAL SALE AND DEAD", LookIn: Excel.XlFindLookIn.xlFormulas,
                                                                  LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                    xlsSheet.Range[xlsSheet.Cells[(intStartRow), 1], xlsSheet.Cells[(find.Row - 1), Column]].EntireRow.delete();


                 
                    CostQty = 5;

                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[4, indexMonth] = drr[0];
                        indexMonth += 3;

                        xlsSheet.Range[xlsSheet.Cells[6, CostQty], xlsSheet.Cells[(find.Row), CostQty]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        CostQty += 3;
                    }

                    xlsSheet = xlsBook.Sheets[6];
                    xlsSheet.Delete();
                    xlsSheet = xlsBook.Sheets[5];
                    xlsSheet.Delete();
                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Delete();
                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Delete();
                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end MaterialUSEDForGMO


        */
        public string getDetailMaterialUSEDForPO(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);



                //  string[] arrGlassType = { "EB", "FC", "GB", "HS", "OTHER" };


                if (dtMonthRange.Rows.Count > 0)
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
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\DetailOfMaterialUSED.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[6];
                    xlsSheet.Name = "Detail Of Material Used";
                    xlsSheet.Cells.Font.Name = "Arial";
                    int Column = 0;
                    int indexMonth = 3;

                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND)LTD. - {0} FACTORY", MaterialOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("Detail of Materail Used", MaterialOBJ.Factory);
                    xlsSheet.Cells[3, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);
                




                    if (dtMonthRange.Rows.Count > 1)
                    {
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 5]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 3];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 8]].EntireColumn.delete();


                        Column = (4 + (dtMonthRange.Rows.Count * 3));

                        //xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1]);
                        xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 6], xlsSheet.Cells[10, 12]].EntireColumn.delete();
                        Column = 10;
                    }


                        rsSum = MaterialDAL.getDetailMaterialReportForPO(MaterialOBJ,"1");

                        // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                        if (rsSum.RecordCount > 0)
                        {


                            //Row
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                            intStartRow += rsSum.RecordCount + 1;

                        }

                    xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 2), Column]].EntireRow.delete();

                    int CostQty = 5;

                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[4, indexMonth] = drr[0];
                        indexMonth += 3;

                        xlsSheet.Range[xlsSheet.Cells[6, CostQty], xlsSheet.Cells[(intStartRow - 1), CostQty]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        CostQty += 3;
                    }



                    //==================================== Summary Used =================================//

                    xlsSheet = xlsBook.Sheets[5];
                    xlsSheet.Name = "Summary Of Material Used";
                    xlsSheet.Cells.Font.Name = "Arial";
                    Column = 0;
                    indexMonth = 3;
                    intStartRow = 6;

                    
                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS (THAILAND)LTD. - {0} FACTORY", MaterialOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("Summary of Materail Used", MaterialOBJ.Factory);
                    xlsSheet.Cells[3, 1] = String.Format("{0} from {1:dd-MMM-yyyy} to {2:dd-MMM-yyyy}", MaterialOBJ.NumberSequenceGroup, MaterialOBJ.DateFrom, MaterialOBJ.DateTo);
                

                    if (dtMonthRange.Rows.Count > 1)
                    {
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 5]];
                            rangeSource.EntireColumn.Copy();
                            rangeDest = xlsSheet.Cells[3, 3];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[10, 8]].EntireColumn.delete();


                        Column = (4 + (dtMonthRange.Rows.Count * 3));

                       // xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1]);
                        xlsSheet.Cells[4, Column] = "Compare " + String.Format("{0:MMM-yyyy} to {1:MMM-yyyy}", dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][0], dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][0]);
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[3, 6], xlsSheet.Cells[10, 12]].EntireColumn.delete();
                        Column = 5;
                    }



                       rsSum = MaterialDAL.getSummaryMaterialReportForPO(MaterialOBJ, "1");

                        // rsSum = MaterialDAL.getSummaryMaterialUSED(MaterialOBJ);
                        if (rsSum.RecordCount > 0)
                        {
                            //Row
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                            rangeSource.EntireRow.Copy();
                            rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 1), Column]];
                            rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                            xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);

                            intStartRow += rsSum.RecordCount;

                        }
                    

                    xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount ), Column]].EntireRow.delete();
                    CostQty = 5;

                    foreach (DataRow drr in dtMonthRange.Rows)
                    {
                        xlsSheet.Cells[4, indexMonth] = drr[0];
                        indexMonth += 3;

                        xlsSheet.Range[xlsSheet.Cells[6, CostQty], xlsSheet.Cells[(intStartRow - 1), CostQty]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-2],0)";
                        CostQty += 3;
                    }


                    xlsSheet = xlsBook.Sheets[4];
                    xlsSheet.Delete();
                    xlsSheet = xlsBook.Sheets[3];
                    xlsSheet.Delete();
                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Delete();
                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Delete();
                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end MaterialUSEDForPO



        public string getMaterialPurchaseForPO(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);




                if (dtMonthRange.Rows.Count > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 7;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialPurchase-PO.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Name = "Detail Purchase";
                    xlsSheet.Cells.Font.Name = "Arial";
                    int Column = 0;
                    int indexMonth = 5;
                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS(THAILAND)LTD.-{0}", MaterialOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("SUMMARY MATERIAL PURCHASE YEAR {0}", MaterialOBJ.DateTo.Year);


                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[7, 5], xlsSheet.Cells[11, 10]];
                            rangeSource.EntireColumn.Copy();

                            rangeDest = xlsSheet.Cells[7, 11];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[7, 11], xlsSheet.Cells[11,22]].EntireColumn.delete();
                        Column = (5 + (dtMonthRange.Rows.Count * 6));


                        //xlsSheet.Cells[3, 10] = "Compare " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[7, 11], xlsSheet.Cells[11, 22]].EntireColumn.delete();
                        Column = 5;
                    }


                     //rsSum = MaterialDAL.getMaterailPurchaseSummaryForPO(MaterialOBJ);

                    rsSum = MaterialDAL.getMaterailPurchaseForPO(dtMonthRange, MaterialOBJ);
                    if (rsSum.RecordCount > 0)
                    {

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount - 1), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        intStartRow += rsSum.RecordCount + 1;

                       // xlsSheet.Range[xlsSheet.Cells[(intStartRow), 1], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();

                        // xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 4), Column]].EntireRow.delete();
                        int KG = 10;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[5, indexMonth] = drr[0];
                            indexMonth += 6;

                            xlsSheet.Range[xlsSheet.Cells[7, KG], xlsSheet.Cells[intStartRow - 2, KG]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-5],0)";
                            KG += 6;
                        }

                    }

                    xlsSheet.Range["C:CM"].Columns.EntireColumn.AutoFit();

                    intStartRow +=  4; //TOTAL

                    rsSum = MaterialDAL.getMaterailPurchaseVender(MaterialOBJ);
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount - 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["C" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, 3], xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, Column]].EntireRow.delete();

                    intStartRow += rsSum.RecordCount + 4; // %

                   // rsSum = MaterialDAL.getMaterailPurchaseVender(MaterialOBJ);
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount - 1), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["C" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, 3], xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, Column]].EntireRow.delete();




                    //============== Summary====================
                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Summary";
                    xlsSheet.Cells.Font.Name = "Arial";
                    Column = 0;
                    indexMonth = 3;
                    intStartRow = 7;
                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS(THAILAND)LTD.-{0}", MaterialOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("SUMMARY MATERIAL PURCHASE YEAR {0}", MaterialOBJ.DateTo.Year);


                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[7, 3], xlsSheet.Cells[11, 8]];
                            rangeSource.EntireColumn.Copy();

                            rangeDest = xlsSheet.Cells[7, 9];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[7, 9], xlsSheet.Cells[11, 20]].EntireColumn.delete();
                        Column = (2 + (dtMonthRange.Rows.Count * 6));


                        //xlsSheet.Cells[3, 10] = "Compare " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[7, 9], xlsSheet.Cells[11, 20]].EntireColumn.delete();
                        Column = 5;
                    }


                    rsSum = MaterialDAL.getMaterailPurchaseSummaryForPO(MaterialOBJ);
                    if (rsSum.RecordCount > 0)
                    {

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount+1), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        intStartRow += rsSum.RecordCount + 1;

                         xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 2), Column]].EntireRow.delete();
                        int KG = 8;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[5, indexMonth] = drr[0];
                            indexMonth += 6;

                            xlsSheet.Range[xlsSheet.Cells[7, KG], xlsSheet.Cells[intStartRow - 2, KG]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-5],0)";
                            KG += 6;
                        }

                    }



                    xlsSheet.Range["C:CM"].Columns.EntireColumn.AutoFit();


                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end MaterialPurchase for PO


        public string getMaterialPurchaseForGMO(MaterialOBJ MaterialOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable dtMonthRange = new DataTable();
                dtMonthRange.Columns.Add("dt", typeof(System.DateTime));
                dtMonthRange.Columns.Add("YearMonth");
                dtMonthRange.Columns.Add("SalesData", typeof(System.Boolean));
                DateTime dateRunning = MaterialOBJ.DateFrom;
                DataRow dr;
                do
                {
                    dr = dtMonthRange.NewRow();
                    dr["dt"] = dateRunning;
                    dr["YearMonth"] = String.Format("{0:dd/MM/yyyy}", dateRunning);
                    dr["SalesData"] = false;
                    dtMonthRange.Rows.Add(dr);
                    dateRunning = dateRunning.AddMonths(1);
                } while (dateRunning < MaterialOBJ.DateTo);




                if (dtMonthRange.Rows.Count > 0)
                {
                    string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                    Excel.Application xlsApp = new Excel.Application();
                    System.Globalization.CultureInfo oldCI;
                    oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    int intStartRow = 7;


                    xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                    xlsApp.SheetsInNewWorkbook = 1;
                    xlsApp.DisplayAlerts = false;
                    xlsApp.Visible = false;
                    Excel.Workbook xlsBookTemplate;
                    Excel.Range rangeSource, rangeDest;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Material\MaterialPurchase-GMO.xls");

                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                    xlsSheet = xlsBook.Sheets[2];
                    xlsSheet.Name = "Detail Purchase";
                    xlsSheet.Cells.Font.Name = "Arial";
                    int Column = 0;
                    int indexMonth = 6;
     
                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS(THAILAND)LTD.-{0}", MaterialOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("SUMMARY MATERIAL PURCHASE YEAR {0}", MaterialOBJ.DateTo.Year);


                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[7, 6], xlsSheet.Cells[11, 11]];
                            rangeSource.EntireColumn.Copy();

                            rangeDest = xlsSheet.Cells[7, 12];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[7, 12], xlsSheet.Cells[11, 23]].EntireColumn.delete();
                        Column = (6 + (dtMonthRange.Rows.Count * 6));


                        //xlsSheet.Cells[3, 10] = "Compare " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[7, 12], xlsSheet.Cells[11, 23]].EntireColumn.delete();
                        Column = 6;
                    }


                    //rsSum = MaterialDAL.getMaterailPurchaseSummaryForGMO(MaterialOBJ);
                    rsSum = MaterialDAL.getMaterailPurchaseForGMO(dtMonthRange,MaterialOBJ);
                    if (rsSum.RecordCount > 0)
                    {

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount - 1), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                       

                        //xlsSheet.Range[xlsSheet.Cells[(intStartRow+rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow +rsSum.RecordCount), Column]].EntireRow.delete();
                        
                        intStartRow += rsSum.RecordCount + 1;
                        // xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 4), Column]].EntireRow.delete();
                        int KG = 11;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[5, indexMonth] = drr[0];
                            indexMonth += 6;

                            xlsSheet.Range[xlsSheet.Cells[7, KG], xlsSheet.Cells[intStartRow - 2, KG]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-5],0)";
                            KG += 6;
                        }

                    }

                    


                    intStartRow += 4; //TOTAL

                    rsSum = MaterialDAL.getMaterailPurchaseVender(MaterialOBJ);
                    //rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[(intStartRow+5), Column]];
                    //rangeSource.EntireRow.Copy();
                    //rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 3], xlsSheet.Cells[(intStartRow + (rsSum.RecordCount*5) - 1), Column]];
                    //rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);



                    //xlsSheet.Range["D" + intStartRow].CopyFromRecordset(rsSum);
                    //xlsSheet.Range[xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, 3], xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, Column]].EntireRow.delete();


                    DataTable dt = new DataTable();
                    System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();
                    adapter.Fill(dt, rsSum);

                    for (int row = 0; row < (dt.Rows.Count-2); row++)
                    {

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[(intStartRow + 6), Column]]; //Edit
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 7), 3], xlsSheet.Cells[(intStartRow +7), Column]]; //Edit
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    }



                        //Get Subgroup
                    rsSum = MaterialDAL.getMaterailPurchaseSubGroup(MaterialOBJ);
                    int x =0;
                    int temp = 0;

                    for (int y = 0; y < (dt.Rows.Count); y++)
                    {
                        
                      xlsSheet.Cells[y+(intStartRow+temp),3] = dt.Rows[x][0];
                      xlsSheet.Range["D" + (y + (intStartRow + temp))].CopyFromRecordset(rsSum);

                      //y += (rsSum.RecordCount + 1)+intStartRow;
                      x += 1;
                      temp += rsSum.RecordCount;
                    }



                    intStartRow += temp + 14; // % Edit

                    rsSum = MaterialDAL.getMaterailPurchaseVender(MaterialOBJ);
                    rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 3], xlsSheet.Cells[intStartRow, Column]];
                    rangeSource.EntireRow.Copy();
                    rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 3], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), Column]];
                    rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                    xlsSheet.Range["D" + intStartRow].CopyFromRecordset(rsSum);
                    xlsSheet.Range[xlsSheet.Cells[(intStartRow) + rsSum.RecordCount, 3], xlsSheet.Cells[(intStartRow) + (rsSum.RecordCount+2), Column]].EntireRow.delete();


                    xlsSheet.Range["D:CM"].Columns.EntireColumn.AutoFit();


                    //============== Summary====================
                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Summary";
                    xlsSheet.Cells.Font.Name = "Arial";
                    Column = 0;
                    indexMonth = 3;
                    intStartRow = 7;
                    xlsSheet.Cells[1, 1] = String.Format("HOYA OPTICS(THAILAND)LTD.-{0}", MaterialOBJ.Factory);
                    xlsSheet.Cells[2, 1] = String.Format("SUMMARY MATERIAL PURCHASE YEAR {0}", MaterialOBJ.DateTo.Year);


                    if (dtMonthRange.Rows.Count > 1)
                    {
                        //Column
                        for (int i = 0; i < dtMonthRange.Rows.Count; i++)
                        {
                            rangeSource = xlsSheet.Range[xlsSheet.Cells[7, 3], xlsSheet.Cells[11, 8]];
                            rangeSource.EntireColumn.Copy();

                            rangeDest = xlsSheet.Cells[7, 9];
                            rangeDest.EntireColumn.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                        }

                        xlsSheet.Range[xlsSheet.Cells[7, 9], xlsSheet.Cells[11, 20]].EntireColumn.delete();
                        Column = (2 + (dtMonthRange.Rows.Count * 6));


                        //xlsSheet.Cells[3, 10] = "Compare " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 2][1] + " - " + dtMonthRange.Rows[dtMonthRange.Rows.Count - 1][1];
                    }
                    else
                    {
                        xlsSheet.Range[xlsSheet.Cells[7, 9], xlsSheet.Cells[11, 20]].EntireColumn.delete();
                        Column = 5;
                    }


                    rsSum = MaterialDAL.getMaterailPurchaseSummaryForGMO(MaterialOBJ);
                    if (rsSum.RecordCount > 0)
                    {

                        rangeSource = xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, Column]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount - 1), Column]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                        intStartRow += rsSum.RecordCount + 1;

                         // xlsSheet.Range[xlsSheet.Cells[(intStartRow), 1], xlsSheet.Cells[(intStartRow + 1), Column]].EntireRow.delete();
                        // xlsSheet.Range[xlsSheet.Cells[(intStartRow), 3], xlsSheet.Cells[(intStartRow + 4), Column]].EntireRow.delete();
                        int KG = 8;
                        foreach (DataRow drr in dtMonthRange.Rows)
                        {
                            xlsSheet.Cells[5, indexMonth] = drr[0];
                            indexMonth += 6;

                            xlsSheet.Range[xlsSheet.Cells[7, KG], xlsSheet.Cells[intStartRow - 2, KG]].FormulaR1C1 = "=IFERROR(R[0]C[-1]/R[0]C[-5],0)";
                            KG += 6;
                        }

                    }


                    xlsSheet.Range["B:CM"].Columns.EntireColumn.AutoFit();


                    xlsApp.DisplayAlerts = true;
                    xlsApp.Visible = true;


                }//end rsSum.RecordCount

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end MaterialPurchase For GMO



    }//end Class
}
