using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace NewVersion.Report.SaleByCustomerByGaikei
{
    class SaleByGaikeiBLL
    {
        SalesByGaikeiDAL SalesByGaikeiDAL = new SalesByGaikeiDAL();



        public string getRequisitionList(SalesByGaikeiOBJ SalesByGaikeiOBJ)
        {
            try
            {

                ADODB.Recordset rsSum = new ADODB.Recordset();
             

                rsSum = SalesByGaikeiDAL.getSaleByCustomer(SalesByGaikeiOBJ);

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
                    int intStartRow = 2;  //StartRow
                    Excel.Workbook xlsBookTemplate;
                    xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\SaleByCustomerByGaikei\SalesByCustomer.xlsx");


                    Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                    Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                    Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                    xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                    xlsBookTemplate.Close();
                    xlsBook.Sheets[xlsBook.Sheets.Count].delete();

                    xlsSheet = xlsBook.Sheets[1];
                    xlsSheet.Name = "Summary";
                    xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                    //xlsSheet.Range[xlsSheet.Cells[intStartRow, 14], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 6]].FormulaR1C1 = @"=IFERROR(R[0]C[-1]/R[0]C[-4],""-"")";

                    xlsSheet.Range["A:K"].EntireColumn.AutoFit();
                    

                    xlsSheet = xlsBook.Sheets[2];
                    //xlsSheet.Name = "Sales by Item by Gaikei";

                    string[] Gaikei = { "(tb_Sales.GAIKEI>=63)"
                                          , "(tb_Sales.GAIKEI BETWEEN 40 AND 62.99)"
                                          , "(tb_Sales.GAIKEI BETWEEN 28 AND 39.99)"
                                          , "(tb_Sales.GAIKEI BETWEEN 16 AND 27.99)"
                                          , "(tb_Sales.GAIKEI <=15.99)"                                     
                                      };


                    string[] GaikeiT = { "TOTAL Summary for ME63"
                                          , "TOTAL Summary for ME40"
                                          , "TOTAL Summary for ME28"
                                          , "TOTAL Summary for ME16"
                                          , "TOTAL Summary for L16"                                     
                                      };

                    int c = 0;
                    foreach (string s in Gaikei)
                    {
                        
                        rsSum = SalesByGaikeiDAL.getSaleByCustomerByGaikei2(SalesByGaikeiOBJ, s); //All RP1/RP2
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);
                       

                        xlsSheet.Cells[(intStartRow+rsSum.RecordCount), 6].Formula = "=SUM(F" + intStartRow + ":F" + (intStartRow+rsSum.RecordCount-1) + ")";
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 7].Formula = "=SUM(G" + intStartRow + ":G" + (intStartRow + rsSum.RecordCount-1) + ")";
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 10].Formula = "=SUM(J" + intStartRow + ":J" + (intStartRow + rsSum.RecordCount-1) + ")";
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount ), 11].Formula = "=SUM(K" + intStartRow + ":K" + (intStartRow + rsSum.RecordCount-1) + ")";

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 5]].Merge();

                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1] = GaikeiT[c].ToString() + " (" + rsSum.RecordCount + " detail record )";


                        intStartRow += rsSum.RecordCount;
                        xlsSheet.Range["A" + (intStartRow), "K" + (intStartRow)].Interior.Color = System.Drawing.Color.Azure;
                        intStartRow += 1;
                        c++;

                    }


                    xlsSheet.Cells[intStartRow, 6].Formula = @"=SUMIF(A2:A"+(intStartRow-1)+@",""TOTAL*"",F2:F"+(intStartRow-1)+")";
                    xlsSheet.Cells[intStartRow, 7].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",G2:G" + (intStartRow - 1) + ")";
                    xlsSheet.Cells[intStartRow, 10].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",J2:J" + (intStartRow - 1) + ")";
                    xlsSheet.Cells[intStartRow, 11].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",K2:K" + (intStartRow - 1) + ")";

                    xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 5]].Merge();
                    xlsSheet.Cells[intStartRow, 1] = " GRAND TOTAL Summary for All Item (" + (intStartRow - 2) + " detail record )";
                    xlsSheet.Range["A" + (intStartRow), "K" + (intStartRow)].Interior.Color = System.Drawing.Color.Beige;


                    //RP1-------------------------------------------------------------------------------------
                    xlsSheet = xlsBook.Sheets[3];
                    intStartRow = 2;
                    c = 0;
                    foreach (string s in Gaikei)
                    {
                       
                        rsSum = SalesByGaikeiDAL.getSaleByCustomerByGaikei(SalesByGaikeiOBJ, s,"RP1"); //External
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);


                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 7].Formula = "=SUM(G" + intStartRow + ":G" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 8].Formula = "=SUM(H" + intStartRow + ":H" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 11].Formula = "=SUM(K" + intStartRow + ":K" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 12].Formula = "=SUM(L" + intStartRow + ":L" + (intStartRow + rsSum.RecordCount - 1) + ")";

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 5]].Merge();

                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1] = GaikeiT[c].ToString() + " (" + rsSum.RecordCount + " detail record )";

                        intStartRow += rsSum.RecordCount;
                        xlsSheet.Range["A" + (intStartRow), "K" + (intStartRow)].Interior.Color = System.Drawing.Color.Azure;
                        intStartRow += 1;
                        c++;

                    }



                    xlsSheet.Cells[intStartRow, 7].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",G2:G" + (intStartRow - 1) + ")";
                    xlsSheet.Cells[intStartRow, 8].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",H2:H" + (intStartRow - 1) + ")";
                    xlsSheet.Cells[intStartRow, 11].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",K2:K" + (intStartRow - 1) + ")";
                    xlsSheet.Cells[intStartRow, 12].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",L2:L" + (intStartRow - 1) + ")";

                    xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 5]].Merge();

                    xlsSheet.Cells[intStartRow, 1] = " GRAND TOTAL Summary for All Item (" + (intStartRow - 2) + " detail record )";
                    xlsSheet.Range["A" + (intStartRow), "K" + (intStartRow)].Interior.Color = System.Drawing.Color.Beige;

                    //RP2-------------------------------------------------------------------------------------------------
                    xlsSheet = xlsBook.Sheets[4];
                    intStartRow = 2;
                    
                    c=0;
                    foreach (string s in Gaikei)
                    {

                        rsSum = SalesByGaikeiDAL.getSaleByCustomerByGaikei(SalesByGaikeiOBJ, s,"RP2"); //External
                        xlsSheet.Range["A" + intStartRow].CopyFromRecordset(rsSum);


                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 7].Formula = "=SUM(G" + intStartRow + ":G" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 8].Formula = "=SUM(H" + intStartRow + ":H" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 11].Formula = "=SUM(K" + intStartRow + ":K" + (intStartRow + rsSum.RecordCount - 1) + ")";
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 12].Formula = "=SUM(L" + intStartRow + ":L" + (intStartRow + rsSum.RecordCount - 1) + ")";

                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 5]].Merge();

                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 1] = GaikeiT[c].ToString() + " (" + rsSum.RecordCount + " detail record )";

                        intStartRow += rsSum.RecordCount;
                        xlsSheet.Range["A" + (intStartRow), "K" + (intStartRow)].Interior.Color = System.Drawing.Color.Azure;
                        intStartRow += 1;
                        c++;
                    }


                    xlsSheet.Cells[intStartRow, 7].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",G2:G" + (intStartRow - 1) + ")";
                    xlsSheet.Cells[intStartRow, 8].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",H2:H" + (intStartRow - 1) + ")";
                    xlsSheet.Cells[intStartRow, 11].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",K2:K" + (intStartRow - 1) + ")";
                    xlsSheet.Cells[intStartRow, 12].Formula = @"=SUMIF(A2:A" + (intStartRow - 1) + @",""TOTAL*"",L2:L" + (intStartRow - 1) + ")";
         
                    xlsSheet.Range[xlsSheet.Cells[intStartRow, 1], xlsSheet.Cells[intStartRow, 5]].Merge();
                    xlsSheet.Cells[intStartRow, 1] = " GRAND TOTAL Summary for All Item (" + (intStartRow - 2) + " detail record )";
                    xlsSheet.Range["A" + (intStartRow), "K" + (intStartRow)].Interior.Color = System.Drawing.Color.Beige;







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
