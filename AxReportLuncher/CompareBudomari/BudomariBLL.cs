using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace NewVersion.CompareBudomari
{
    class BudomariBLL
    {
        BudomariDAL BudomariDAL = new BudomariDAL();
        public string GetBodomari(BudomariOBJ BudomariOBJ)
        {
            try
            {

                
                ADODB.Recordset rsSum = new ADODB.Recordset();
                DataTable MasterTable = new DataTable();

                string strSystemPath = System.IO.Directory.GetCurrentDirectory();

                Excel.Application xlsApp = new Excel.Application();
                System.Globalization.CultureInfo oldCI;
                oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                int intStartRow =8;
                int RowsSet = 0;
                 

                xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
                xlsApp.SheetsInNewWorkbook = 1;
                xlsApp.DisplayAlerts = false;
                xlsApp.Visible = false;
                Excel.Workbook xlsBookTemplate;
                Excel.Range rangeSource, rangeDest;
                xlsBookTemplate = xlsApp.Workbooks.Open(strSystemPath + @"\ExcelTemplate\Budomari\Budomari.xlsx");

                Excel.Worksheet xlsSheetTemplate = xlsBookTemplate.Worksheets[1];
                Excel.Workbook xlsBook = xlsApp.Workbooks.Add();
                Excel.Worksheet xlsSheet = xlsBook.Worksheets[1];
                xlsBookTemplate.Sheets.Copy(Before: xlsBook.Sheets[1]);
                xlsBookTemplate.Close();
                xlsBook.Sheets[xlsBook.Sheets.Count].delete();


                xlsSheet = xlsBook.Sheets[1];
                //xlsSheet.Name = "Group mat compare";
                xlsSheet.Cells.Font.Name = "Arial";

                xlsSheet.Cells[1, 1] = String.Format("タイＭＯ工場{0}・{1}の仕掛の物量変動と歩留 -   ", BudomariOBJ.GetSheet1, BudomariOBJ.GetSheet2);
                xlsSheet.Cells[4, 4] = BudomariOBJ.DateFrom;
                xlsSheet.Cells[4, 21] = BudomariOBJ.DateTo;
              
                MasterTable = BudomariDAL.GetMasterGroup();
                  


                foreach (DataRow Group in MasterTable.Rows)
                {
                    rsSum = BudomariDAL.GetBudomari(Group[0].ToString(), BudomariOBJ);
                    if (rsSum.RecordCount > 0)
                    {

                      
                        rangeSource = xlsSheet.Range[xlsSheet.Cells[(intStartRow+1), 1], xlsSheet.Cells[(intStartRow+1), 38]];
                        rangeSource.EntireRow.Copy();
                        rangeDest = xlsSheet.Range[xlsSheet.Cells[(intStartRow+2), 1], xlsSheet.Cells[(intStartRow + rsSum.RecordCount+4), 38]];
                        rangeDest.EntireRow.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);

                        xlsSheet.Range["B" + (intStartRow+1)].CopyFromRecordset(rsSum);



                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount+2), 4], xlsSheet.Cells[(intStartRow + rsSum.RecordCount+2), 18]].Formula =  String.Format(@"=SUM(D$" + intStartRow + ":D$" +( intStartRow + rsSum.RecordCount) + ")");
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), 21], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), 35]].Formula = String.Format(@"=SUM(U$" + intStartRow + ":U$" + (intStartRow + rsSum.RecordCount) + ")");


                        //Budomari                                                                                                                                                                                                                 //  =IF(J9<=0,0,IF((D9+G9-R9)=0,0,J9/(D9+G9-L9-M9-N9-O9-R9)))
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 19], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), 19]].Formula = "=IF(J" + (intStartRow + 1) + "<=0,0,IF((D" + (intStartRow + 1) + "+G" + (intStartRow + 1) + "-R" + (intStartRow + 1) + ")=0,0,J" + (intStartRow + 1) + "/(D" + (intStartRow + 1) + "+G" + (intStartRow + 1) + "-L" + (intStartRow + 1) + "-M" + (intStartRow + 1) + "-N" + (intStartRow + 1) + "-O" + (intStartRow + 1) + "-R" + (intStartRow + 1) + ")))";
                                                                                                                                                                                                                                                  //=IF(AA9<=0,0,IF((U9+X9-AI9)=0,0,AA9/(U9+X9-AC9-AD9-AE9-AF9-AI9)))     
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 36], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), 36]].Formula = "=IF(AA" + (intStartRow + 1) + "<=0,0,IF((U" + (intStartRow + 1) + "+X" + (intStartRow + 1) + "-AI" + (intStartRow + 1) + ")=0,0,AA" + (intStartRow + 1) + "/(U" + (intStartRow + 1) + "+X" + (intStartRow + 1) + "-AC" + (intStartRow + 1) + "-AD" + (intStartRow + 1) + "-AE" + (intStartRow + 1) + "-AF" + (intStartRow + 1) + "-AI" + (intStartRow + 1) + ")))";
                   

                        xlsSheet.Cells[(intStartRow+rsSum.RecordCount+2), 2] = Group[1].ToString();
                        xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), 3] = "TOTAL";
                       


                        xlsSheet.Range[xlsSheet.Cells[(intStartRow+1),7], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 7]].FormulaR1C1 = "=R[0]C[-2]-R[0]C[-1]";
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow+1), 10], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 10]].FormulaR1C1 = "=R[0]C[-2]-R[0]C[-1]";
                        // NET NG  =+K10+P10-L10
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 17], xlsSheet.Cells[(intStartRow + rsSum.RecordCount ), 17]].Formula = "=K" + (intStartRow + 1) + "+P" + (intStartRow + 1) + "-L" + (intStartRow + 1);
                       
                        //-======================
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 24], xlsSheet.Cells[(intStartRow + rsSum.RecordCount ), 24]].FormulaR1C1 = "=R[0]C[-2]-R[0]C[-1]";
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 27], xlsSheet.Cells[(intStartRow + rsSum.RecordCount ), 27]].FormulaR1C1 = "=R[0]C[-2]-R[0]C[-1]";
                        //NET NG =+AB9+AG9-AC9
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 34], xlsSheet.Cells[(intStartRow + rsSum.RecordCount), 34]].Formula = "=AB" + (intStartRow + 1) + "+AG" + (intStartRow + 1) + "-AC" + (intStartRow + 1);
                  

                        //====================== AJ - S
                        //xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 38], xlsSheet.Cells[(intStartRow + rsSum.RecordCount - 1), 38]].FormulaR1C1 = "=R[0]C[-2]-R[0]C[-19]";


                        //=============================================================================================================================+D11+G11-J11-Q11-R11-L11-M11-N11-O11
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 20], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), 20]].Formula = "=D" + (intStartRow + 1) + "+G" + (intStartRow + 1) + "-J" + (intStartRow + 1) + "-Q" + (intStartRow + 1) + "-R" + (intStartRow + 1) + "-L" + (intStartRow + 1) + "-M" + (intStartRow + 1) + "-N" + (intStartRow + 1) + "-O" + (intStartRow + 1);
                        //=============================================================================================================================+U9+X9-AA9-AH9-AI9-AC9-AD9-AE9-AF9
                        xlsSheet.Range[xlsSheet.Cells[(intStartRow + 1), 37], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 2), 37]].Formula = "=U" + (intStartRow + 1) + "+X" + (intStartRow + 1) + "-AA" + (intStartRow + 1) + "-AH" + (intStartRow + 1) + "-AI" + (intStartRow + 1) + "-AC" + (intStartRow + 1) + "-AD" + (intStartRow + 1) + "-AE" + (intStartRow + 1) + "-AF" + (intStartRow + 1);
               

                       

                        if (Group[0].ToString() == "SLR")
                        {
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 4], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 18]].Formula = "=SUMIF($C$" + 9 + ":$C$" + (intStartRow + rsSum.RecordCount + 2) + @",""TOTAL"",D$" + 9 + ":D$" + (intStartRow + rsSum.RecordCount + 2) + ")";
                            //Budomari
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 19], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 19]].Formula
                                = "=IF(J" + (intStartRow + rsSum.RecordCount + 4) + "<=0,0,IF((D" + (intStartRow + rsSum.RecordCount + 4) + "+G" + (intStartRow + rsSum.RecordCount + 4) +
                                "-R" + (intStartRow + rsSum.RecordCount + 4) + ")=0,0,J" + (intStartRow + rsSum.RecordCount + 4) + "/(D" + (intStartRow + rsSum.RecordCount + 4) + "+G"
                                + (intStartRow + rsSum.RecordCount + 4) + "-L" + (intStartRow + rsSum.RecordCount + 4) + "-M" + (intStartRow + rsSum.RecordCount + 4) + "-N" + (intStartRow + rsSum.RecordCount + 4)
                                + "-O" + (intStartRow + rsSum.RecordCount + 4) + "-R" + (intStartRow + rsSum.RecordCount + 4) + ")))";



                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 21], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 35]].Formula = "=SUMIF($C$" + 9 + ":$C$" + (intStartRow + rsSum.RecordCount + 2) + @",""TOTAL"",U$" + 9 + ":U$" + (intStartRow + rsSum.RecordCount + 2) + ")";
                            //Budomari
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 36], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 36]].Formula
                                = "=IF(AA" + (intStartRow + rsSum.RecordCount + 4) + "<=0,0,IF((U" + (intStartRow + rsSum.RecordCount + 4) + "+X" + (intStartRow + rsSum.RecordCount + 4) + "-AI" +
                                (intStartRow + rsSum.RecordCount + 4) + ")=0,0,AA" + (intStartRow + rsSum.RecordCount + 4) + "/(U" + (intStartRow + rsSum.RecordCount + 4) + "+X" + (intStartRow + rsSum.RecordCount + 4) +
                                "-AC" + (intStartRow + rsSum.RecordCount + 4) + "-AD" + (intStartRow + rsSum.RecordCount + 4) + "-AE" + (intStartRow + rsSum.RecordCount + 4) + "-AF" + (intStartRow + rsSum.RecordCount + 4) +
                                "-AI" + (intStartRow + rsSum.RecordCount + 4) + ")))";



                            xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 2] = "通常流動　計";
                            xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 3] = "①+②+③";
                            RowsSet = (intStartRow + rsSum.RecordCount + 4) + 1;
                            intStartRow += 2;
                        }

                        if (Group[0].ToString() == "DEFECTIVE SLR")
                        {
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 4], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 18]].Formula = "=SUMIF($C$" + RowsSet + ":$C$" + (intStartRow + rsSum.RecordCount + 2) + @",""TOTAL"",D$" + RowsSet + ":D$" + (intStartRow + rsSum.RecordCount + 2) + ")";
                            //Budomai
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 19], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 19]].Formula =
                                "=IF(J" + (intStartRow + rsSum.RecordCount + 4) + "<=0,0,IF((D" + (intStartRow + rsSum.RecordCount + 4) + "+G" + (intStartRow + rsSum.RecordCount + 4) +
                                "-R" + (intStartRow + rsSum.RecordCount + 4) + ")=0,0,J" + (intStartRow + rsSum.RecordCount + 4) + "/(D" + (intStartRow + rsSum.RecordCount + 4) + "+G" +
                               (intStartRow + rsSum.RecordCount + 4) + "-L" + (intStartRow + rsSum.RecordCount + 4) + "-M" + (intStartRow + rsSum.RecordCount + 4) + "-N" + (intStartRow + rsSum.RecordCount + 4) + "-O" +
                             (intStartRow + rsSum.RecordCount + 4) + "-R" + (intStartRow + rsSum.RecordCount + 4) + ")))";


                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 21], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 35]].Formula = "=SUMIF($C$" + RowsSet + ":$C$" + (intStartRow + rsSum.RecordCount + 2) + @",""TOTAL"",U$" + RowsSet + ":U$" + (intStartRow + rsSum.RecordCount + 2) + ")";
                          
                            //Budomari
                            xlsSheet.Range[xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 36], xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 36]].Formula =
                               "=IF(AA" + (intStartRow + rsSum.RecordCount + 4) + "<=0,0,IF((U" + (intStartRow + rsSum.RecordCount + 4) + "+X" + (intStartRow + rsSum.RecordCount + 4) + "-AI" + (intStartRow + rsSum.RecordCount + 4) + ")=0,0,AA" + (intStartRow + rsSum.RecordCount + 4) + "/(U" + (intStartRow + rsSum.RecordCount + 4) + "+X" + (intStartRow + rsSum.RecordCount + 4) + "-AC" + (intStartRow + rsSum.RecordCount + 4) + "-AD" + (intStartRow + rsSum.RecordCount + 4) + "-AE" + (intStartRow + rsSum.RecordCount + 4) + "-AF" + (intStartRow + rsSum.RecordCount + 4) + "-AI" + (intStartRow + rsSum.RecordCount + 4) + ")))";
                   

                            
                            xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 2] = "通常流動　計";
                            xlsSheet.Cells[(intStartRow + rsSum.RecordCount + 4), 3] = "④+⑤+⑥";

                            intStartRow += 2;
                        }

               intStartRow += rsSum.RecordCount+3;
                    }
                }


                Excel.Range find,Exclude;
                Exclude = xlsSheet.Range["B:B"].Find(What: "Budo Excluded Defective SH IN/Mat used", LookIn: Excel.XlFindLookIn.xlFormulas,
                                                               LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

                find = xlsSheet.Range["C:C"].Find(What: "①+②+③", LookIn: Excel.XlFindLookIn.xlFormulas,
                                                           LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlNext, MatchCase: false, SearchFormat: false);

               // xlsSheet.Cells[Exclude.Row, 16]  // xlsSheet.Cells[find.Row, 4];
                //=+D18+E18-R16-L16-M16-N16-O16

                xlsSheet.Range[xlsSheet.Cells[Exclude.Row, 17], xlsSheet.Cells[Exclude.Row, 17]].Formula = "=J" + find.Row;
                xlsSheet.Range[xlsSheet.Cells[Exclude.Row, 18], xlsSheet.Cells[Exclude.Row, 18]].Formula = "=D" + find.Row + "+E" + find.Row + "-R" + find.Row + "-L" + find.Row + "-M" + find.Row + "-N" + find.Row + "-O" + find.Row;

                xlsSheet.Range[xlsSheet.Cells[Exclude.Row, 34], xlsSheet.Cells[Exclude.Row, 34]].Formula = "=AA" + find.Row;
                xlsSheet.Range[xlsSheet.Cells[Exclude.Row, 35], xlsSheet.Cells[Exclude.Row, 35]].Formula = "=U" + find.Row + "+V" + find.Row + "-AI" + find.Row + "-AC" + find.Row + "-AD" + find.Row + "-AE" + find.Row + "-AF" + find.Row;




                //IF(J9<=0,0,IF((D9+G9-R9)=0,0,J9/(D9+G9-L9-M9-N9-O9-R9)))

                //"Budo Excluded Defective SH IN/Mat used";

                
                xlsSheet.Range["D:AF"].Columns.EntireColumn.AutoFit();
                xlsSheet.Range["AH:AL"].Columns.EntireColumn.AutoFit();
                xlsApp.DisplayAlerts = true;
                xlsApp.Visible = true;


                //.Close(0);
                //xlsApp.Quit();

                return null;

            }

            catch (Exception ex)
            {
                return ex.Message;
            }


        }//end StockCoompare





    }
}
