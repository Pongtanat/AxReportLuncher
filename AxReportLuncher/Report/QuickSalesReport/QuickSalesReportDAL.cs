using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.Report.QuickSales_Report
{
    class QuickSalesReportDAL
    {

        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();

        string _strFac;
        int ShipLoc;
        public DataTable getNumberSequenceGroup(string strFactory, int intShipmentLocation)
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(" SELECT DISTINCT NUMBERSEQUENCEGROUPID");
            sbSql.AppendLine(" FROM ECL_SalesImportSetup");

            if (intShipmentLocation == 1)
            {
                if (strFactory == "GMO")
                {
                    strFactory = "MO";
                }
                else
                {
                    _strFac = strFactory;
                }
            }
            else
            {
                sbSql.AppendLine(" WHERE InventSiteID='" + strFactory + "' AND ShipmentLoc=" + intShipmentLocation);

            }

            ShipLoc = intShipmentLocation;
            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }


        public ADODB.Recordset getQuickSalesReport(QuickSaleReportOBJ QuickSaleReportOBJ,DataTable strNumberSequence,bool trading )
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom  = new DateTime(QuickSaleReportOBJ.DateFrom.Year,QuickSaleReportOBJ.DateFrom.Month,1);
            DateTime dtTo = new DateTime(QuickSaleReportOBJ.DateTo.Year, QuickSaleReportOBJ.DateTo.Month, 1);

            DateTime CheckDate = QuickSaleReportOBJ.DateFrom.AddMonths(1).AddDays(-1);
           // CheckDate.AddMonths(1).AddDays(-1);

            String strFac = QuickSaleReportOBJ.strFactory;
            if (QuickSaleReportOBJ.strFactory == "GMO")
            {

                strFac = "MO";
            }
            else
            {
                strFac = QuickSaleReportOBJ.strFactory;

            }


            sbSql.AppendLine("SELECT NUMBERSEQUENCEGROUP2 [Customer],SUM(THB),SUM(PCS)");


            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN NUMBERSEQUENCEGROUP='" + strFac + "-EXT' OR NUMBERSEQUENCEGROUP='" + strFac + "-REXT' OR NUMBERSEQUENCEGROUP='" + strFac + "-CEXT' OR NUMBERSEQUENCEGROUP='" + strFac + "-TRD'  OR NUMBERSEQUENCEGROUP='" + strFac + "-CTRD' OR NUMBERSEQUENCEGROUP='" + strFac + "-RTRD'   THEN 'EXTERNAL SALE' + CASE WHEN NOT(NAMEALIAS IS NULL) THEN ' ('+NAMEALIAS+')' ELSE ' TOTAL' END ELSE ");
            sbSql.AppendLine("CASE WHEN NUMBERSEQUENCEGROUP='" + strFac + "-INT' OR NUMBERSEQUENCEGROUP='" + strFac + "-RINT' OR NUMBERSEQUENCEGROUP='" + strFac + "-CINT' THEN 'INTERNAL SALE' + CASE WHEN NOT(NAMEALIAS IS NULL) THEN ' ('+NAMEALIAS+')' ELSE ' TOTAL' END ELSE ");
            sbSql.AppendLine("NUMBERSEQUENCEGROUP END  END NUMBERSEQUENCEGROUP2");

            while (dtFrom <= dtTo)
            {

                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [LineAmountMST]/1000 ELSE 0 END)[THB]",dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS]/1000 ELSE 0 END) [PCS]",dtFrom.Month));
                //sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN '' ELSE 0 END) [BAHT/PCS] ", dtFrom.Month));
                //sbSql.AppendLine(",''[BAHT/PCS]");
                dtFrom = dtFrom.AddMonths(1);
              

            }
           

            
            sbSql.AppendLine("FROM (");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",NAMEALIAS");



            if (QuickSaleReportOBJ.strFactory == "GMO")
            {
                sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM([SET]) ELSE 0 END [PCS]");
            }
            else
            {
                sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(PCS) ELSE 0 END [PCS]");
            }

            

            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
            sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + QuickSaleReportOBJ.Factory + "'");


            if (trading)
            {
                sbSql.AppendLine(" AND (HOYA_TRADING = 1");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CTRD') ");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-TRD')  ");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD'))  ");
            }
            else
            {
                if (QuickSaleReportOBJ.Factory == "RP")
                {
                    if (QuickSaleReportOBJ.DateTo < CheckDate)
                    {
                        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') ");
                        sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-INT') ");
                        sbSql.AppendLine("OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT') ");
                        sbSql.AppendLine("OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT'))");

                    }
                    else
                    {
                        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT')  ");
                        sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-INT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT'))  ");
                    }

                }
                else
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT')  ");
                    sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-INT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT'))  ");

                }

                
            }
    

            sbSql.AppendLine(" AND CUSTGROUP IN ('" + QuickSaleReportOBJ.CustomerGroup + "')");

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateTo) + "',103)");
            sbSql.AppendLine("AND INVOICEACCOUNT !='ARAF001'");
            sbSql.AppendLine("AND  ECL_SALESCOMERCIAL = 1");
            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,ECL_SALESCOMERCIAL,INVOICEDATE");
            sbSql.AppendLine(")salesTotal");


            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS  ");
            sbSql.AppendLine("HAVING NOT NUMBERSEQUENCEGROUP IS NULL  )Data ");

         
            /*
            if (strNumberSequence.Rows.Count>0)
            {

                for (int i = 0; i < strNumberSequence.Rows.Count; i++)
                {
                    QuickSaleReportOBJ.numbersequence2 +="'" +strNumberSequence.Rows[i][0]+"',";

                }
                QuickSaleReportOBJ.numbersequence2 += "''";
                sbSql.AppendLine("WHERE NUMBERSEQUENCEGROUP2 IN (" + QuickSaleReportOBJ.numbersequence2 + ") ");
            }

            */
            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP2 ");
            //sbSql.AppendLine("ORDER BY NUMBERSEQUENCEGROUP2 ASC");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getQuickSalesReportNOCOME(QuickSaleReportOBJ QuickSaleReportOBJ,bool trading)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(QuickSaleReportOBJ.DateFrom.Year, QuickSaleReportOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(QuickSaleReportOBJ.DateTo.Year, QuickSaleReportOBJ.DateTo.Month, 1);

            DateTime CheckDate = QuickSaleReportOBJ.DateFrom.AddMonths(1).AddDays(-1);

            String strFac = QuickSaleReportOBJ.strFactory;
            if (QuickSaleReportOBJ.strFactory == "GMO")
            {

                strFac = "MO";
            }
            else
            {
                strFac = QuickSaleReportOBJ.strFactory;

            }


            sbSql.AppendLine("SELECT");
        
            while (dtFrom <= dtTo)
            {
                  sbSql.AppendLine(String.Format("SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET]/1000 ELSE 0 END) [SET]", dtFrom.Month));
                  dtFrom = dtFrom.AddMonths(1);


            }



            sbSql.AppendLine("FROM (");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",NAMEALIAS");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=2 THEN SUM([SET]) ELSE 0 END [SET]");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=2 THEN SUM(PCS) ELSE 0 END [PCS]");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=2 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=2 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
            sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + QuickSaleReportOBJ.Factory + "'");


            if (trading)
            {
                sbSql.AppendLine(" AND (HOYA_TRADING = 1)");
            }else
            {

                if (QuickSaleReportOBJ.DateTo < CheckDate)
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT')  ");
                    sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-INT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT'))  ");

                }
                else
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT')  ");
                    sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-INT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT'))  ");
                    //sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-RNOC'))  ");
                } 


            }

            sbSql.AppendLine(" AND CUSTGROUP IN ('" + QuickSaleReportOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateTo) + "',103)");
            sbSql.AppendLine("AND INVOICEACCOUNT !='ARAF001'");
            sbSql.AppendLine("AND  ECL_SALESCOMERCIAL = 2");
            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,ECL_SALESCOMERCIAL,INVOICEDATE");
            sbSql.AppendLine(")salesTotal");

            /*
           // sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS  ");
           // sbSql.AppendLine("HAVING NOT NUMBERSEQUENCEGROUP IS NULL  )Data ");

            if (strNumberSequence.Rows.Count > 0)
            {

                for (int i = 0; i < strNumberSequence.Rows.Count; i++)
                {
                    QuickSaleReportOBJ.numbersequence2 += "'" + strNumberSequence.Rows[i][0] + "',";

                }
                QuickSaleReportOBJ.numbersequence2 += "''";
                sbSql.AppendLine("WHERE NUMBERSEQUENCEGROUP2 IN (" + QuickSaleReportOBJ.numbersequence2 + ") ");
            }


           // sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP2 ");
            //sbSql.AppendLine("ORDER BY NUMBERSEQUENCEGROUP2 ASC");
            */


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getQuickSalesHeader(QuickSaleReportOBJ QuickSaleReportOBJ,bool trading)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(QuickSaleReportOBJ.DateFrom.Year, QuickSaleReportOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(QuickSaleReportOBJ.DateTo.Year, QuickSaleReportOBJ.DateTo.Month, 1);

            DateTime CheckDate = QuickSaleReportOBJ.DateFrom.AddMonths(1).AddDays(-1);
            // CheckDate.AddMonths(1).AddDays(-1);

            String strFac = QuickSaleReportOBJ.strFactory;
            if (QuickSaleReportOBJ.strFactory == "GMO")
            {

                strFac = "MO";
            }
            else
            {
                strFac = QuickSaleReportOBJ.strFactory;

            }


            sbSql.AppendLine("SELECT NUMBERSEQUENCEGROUP2 [Customer]");
            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN NUMBERSEQUENCEGROUP='" + strFac + "-EXT' OR NUMBERSEQUENCEGROUP='" + strFac + "-REXT' OR NUMBERSEQUENCEGROUP='" + strFac + "-CEXT' OR NUMBERSEQUENCEGROUP='" + strFac + "-TRD'  OR NUMBERSEQUENCEGROUP='" + strFac + "-CTRD' OR NUMBERSEQUENCEGROUP='" + strFac + "-RTRD'   THEN 'EXTERNAL SALE' + CASE WHEN NOT(NAMEALIAS IS NULL) THEN ' ('+NAMEALIAS+')' ELSE ' TOTAL' END ELSE ");
            sbSql.AppendLine("CASE WHEN NUMBERSEQUENCEGROUP='" + strFac + "-INT' OR NUMBERSEQUENCEGROUP='" + strFac + "-RINT' OR NUMBERSEQUENCEGROUP='" + strFac + "-CINT' THEN 'INTERNAL SALE' + CASE WHEN NOT(NAMEALIAS IS NULL) THEN ' ('+NAMEALIAS+')' ELSE ' TOTAL' END ELSE ");
            sbSql.AppendLine("NUMBERSEQUENCEGROUP END  END NUMBERSEQUENCEGROUP2");

         

            sbSql.AppendLine("FROM (");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",NAMEALIAS");
          
            sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + QuickSaleReportOBJ.Factory + "'");


            if (trading)
            {
                sbSql.AppendLine(" AND (HOYA_TRADING = 1)");
            }
            else
            {
                if (QuickSaleReportOBJ.Factory == "RP")
                {
                    if (QuickSaleReportOBJ.DateTo < CheckDate)
                    {
                        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') ");
                        sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-INT'))  ");

                    }
                    else
                    {
                        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT')  ");
                        sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-INT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT'))  ");
                    }

                }
                else
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT')  ");
                    sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-INT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT'))  ");

                }


            }


            sbSql.AppendLine(" AND CUSTGROUP IN ('" + QuickSaleReportOBJ.CustomerGroup + "')");

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateTo) + "',103)");
            sbSql.AppendLine("AND INVOICEACCOUNT !='ARAF001'");
            sbSql.AppendLine("AND  ECL_SALESCOMERCIAL = 1");
            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,ECL_SALESCOMERCIAL,INVOICEDATE");
            sbSql.AppendLine(")salesTotal");


            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS  ");
            sbSql.AppendLine("HAVING NOT NUMBERSEQUENCEGROUP IS NULL  )Data ");




            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP2 ");
            //sbSql.AppendLine("ORDER BY NUMBERSEQUENCEGROUP2 ASC");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getQuickSalesAllHeader(QuickSaleReportOBJ QuickSaleReportOBJ, bool trading)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(QuickSaleReportOBJ.DateFrom.Year, QuickSaleReportOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(QuickSaleReportOBJ.DateTo.Year, QuickSaleReportOBJ.DateTo.Month, 1);

            DateTime CheckDate = QuickSaleReportOBJ.DateFrom.AddMonths(1).AddDays(-1);

            String strFac = QuickSaleReportOBJ.strFactory;
           

            sbSql.AppendLine("SELECT NUMBERSEQUENCEGROUP2 [Customer]");
            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("'EXTERNAL SALE' + CASE WHEN NOT(NAMEALIAS IS NULL) THEN ' ('+NAMEALIAS+')' ELSE 'TOTAL'END  NUMBERSEQUENCEGROUP2");

            sbSql.AppendLine("FROM (");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",NAMEALIAS");
            sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId IN ('RP','GMO','PO')");

            sbSql.AppendLine("  AND (NUMBERSEQUENCEGROUP LIKE ('%-EXT') OR  NUMBERSEQUENCEGROUP LIKE ('%-REXT') OR  NUMBERSEQUENCEGROUP LIKE ('%-CEXT'))   ");

            sbSql.AppendLine(" AND CUSTGROUP IN ('" + QuickSaleReportOBJ.CustomerGroup + "')");

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateTo) + "',103)");
            sbSql.AppendLine("AND INVOICEACCOUNT !='ARAF001'");
            sbSql.AppendLine("AND  ECL_SALESCOMERCIAL = 1");
            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,ECL_SALESCOMERCIAL,INVOICEDATE");
            sbSql.AppendLine(")salesTotal");


            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS  ");
            sbSql.AppendLine("HAVING NOT NUMBERSEQUENCEGROUP IS NULL  )Data ");

            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP2 ");
            //sbSql.AppendLine("ORDER BY NUMBERSEQUENCEGROUP2 ASC");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getQuickSalesSupport(QuickSaleReportOBJ QuickSaleReportOBJ,bool NumberSequen,bool Return,bool trading)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(QuickSaleReportOBJ.DateFrom.Year, QuickSaleReportOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(QuickSaleReportOBJ.DateTo.Year, QuickSaleReportOBJ.DateTo.Month, 1);

            String strFac = QuickSaleReportOBJ.strFactory;
            if (QuickSaleReportOBJ.strFactory == "GMO")
            {

                strFac = "MO";
            }
            else
            {
                strFac = QuickSaleReportOBJ.strFactory;

            }


            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("INVOICEDATE [InvoiceDate]");

            if (NumberSequen)//External
            {
                if (strFac == "MO")
                {
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='USD' THEN [SET] ELSE 0 END)[PCS]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='USD' THEN [LineAmount] ELSE 0 END)[LineAmount]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='USD' THEN [LineAmountMST] ELSE 0 END)[LineAmountMST]");

                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='JPY' THEN [SET] ELSE 0 END)[PCS]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='JPY' THEN [LineAmount] ELSE 0 END)[LineAmount]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='JPY' THEN [LineAmountMST] ELSE 0 END)[LineAmountMST]");

                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='CNY' THEN [SET] ELSE 0 END)[PCS]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='CNY' THEN [LineAmount] ELSE 0 END)[LineAmount]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='CNY' THEN [LineAmountMST] ELSE 0 END)[LineAmountMST]");

                    if (Return)
                    {
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN  [SET] ELSE 0 END)[PCS]");
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN  [LineAmountMST] ELSE 0  END)[LineAmountMST]");
                    }
                    else
                    {

                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX006' THEN [SET] ELSE 0 END ELSE 0  END)[PCS]"); //sony
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX006' THEN [LineAmountMST] ELSE 0 END ELSE 0  END)[LineAmountMST]");

                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX014' THEN [SET] ELSE 0 END ELSE 0  END)[PCS]");//ricoh
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX014' THEN [LineAmountMST] ELSE 0 END ELSE 0  END)[LineAmountMST]");

                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX008' THEN [SET] ELSE 0 END ELSE 0  END)[PCS]");//nikon
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX008' THEN [LineAmountMST] ELSE 0 END ELSE 0  END)[LineAmountMST]");

                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX005' THEN [SET] ELSE 0 END ELSE 0  END)[PCS]");//nidec
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX005' THEN [LineAmountMST] ELSE 0 END ELSE 0  END)[LineAmountMST]");
                    }
                }
                else
                {

                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='USD' THEN [PCS] ELSE 0 END)[PCS]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='USD' THEN [LineAmount] ELSE 0 END)[LineAmount]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='USD' THEN [LineAmountMST] ELSE 0 END)[LineAmountMST]");

                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='JPY' THEN [PCS] ELSE 0 END)[PCS]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='JPY' THEN [LineAmount] ELSE 0 END)[LineAmount]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='JPY' THEN [LineAmountMST] ELSE 0 END)[LineAmountMST]");

                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='CNY' THEN [PCS] ELSE 0 END)[PCS]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='CNY' THEN [LineAmount] ELSE 0 END)[LineAmount]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='CNY' THEN [LineAmountMST] ELSE 0 END)[LineAmountMST]");

                    if (Return)
                    {
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN  [PCS] ELSE 0 END)[PCS]");
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN  [LineAmountMST] ELSE 0  END)[LineAmountMST]");
                    }
                    else
                    {

                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX006' THEN [PCS] ELSE 0 END ELSE 0  END)[PCS]");//sony
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX006' THEN [LineAmountMST] ELSE 0 END ELSE 0  END)[LineAmountMST]");

                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX014' THEN [PCS] ELSE 0 END ELSE 0  END)[PCS]");//ricoh
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX014' THEN [LineAmountMST] ELSE 0 END ELSE 0  END)[LineAmountMST]");

                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX008' THEN [PCS] ELSE 0 END ELSE 0  END)[PCS]");//nicon
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX008' THEN [LineAmountMST] ELSE 0 END ELSE 0  END)[LineAmountMST]");

                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX005' THEN [PCS] ELSE 0 END ELSE 0  END)[PCS]");//nidec
                        sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN CASE WHEN INVOICEACCOUNT ='AREX005' THEN [LineAmountMST] ELSE 0 END ELSE 0  END)[LineAmountMST]");
                    }
                }

            }
            else//Internal
            {
                if (strFac == "MO")
                {
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN [SET] ELSE 0 END)[PCS]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN [LineAmountMST] ELSE 0 END)[LineAmountMST]");
                }
                else
                {
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN [PCS] ELSE 0 END)[PCS]");
                    sbSql.AppendLine(",SUM(CASE WHEN CURRENCYCODEISO ='THB' THEN [LineAmountMST] ELSE 0 END)[LineAmountMST]");

                }
              
            }
        

            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + QuickSaleReportOBJ.Factory + "'");

            if (trading)
            {
                if (Return)
                {
                 
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-CTRD') ");
                    // sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-TRD'))  ");
                    sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD'))  ");

                }
                else
                {
                    sbSql.AppendLine(" AND (HOYA_TRADING = 1)");
                    //sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CTRD') ");
                    //sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-TRD'))  ");
                    //sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD'))  ");
                    // sbSql.AppendLine(" AND (HOYA_TRADING = 1");

                }
            } //end Trading
            else
            {

                if (NumberSequen)
                {
                    if (Return)
                    {
                        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT')) ");
                    }
                    else
                    {
                        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT')) ");
                    }

                }
                else
                {
                    if (Return)
                    {
                        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT'))  ");

                    }
                    else
                    {
                        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-INT') OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT'))  ");

                    }
                }

            }



         

            sbSql.AppendLine(" AND CUSTGROUP IN ('" + QuickSaleReportOBJ.CustomerGroup + "')");

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateTo) + "',103)");

            sbSql.AppendLine("AND INVOICEACCOUNT !='ARAF001'");
            sbSql.AppendLine("AND  ECL_SALESCOMERCIAL = 1");
            sbSql.AppendLine("GROUP BY INVOICEDATE");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getQuickSalesSupportNocome(QuickSaleReportOBJ QuickSaleReportOBJ,string trading)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(QuickSaleReportOBJ.DateFrom.Year, QuickSaleReportOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(QuickSaleReportOBJ.DateTo.Year, QuickSaleReportOBJ.DateTo.Month, 1);

            String strFac = QuickSaleReportOBJ.strFactory;
            if (QuickSaleReportOBJ.strFactory == "GMO")
            {

                strFac = "MO";
            }
            else
            {
                strFac = QuickSaleReportOBJ.strFactory;

            }


            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("INVOICEDATE [InvoiceDate]");

            if (QuickSaleReportOBJ.strFactory == "GMO")
            {
                sbSql.AppendLine(",SUM([SET])[PCS]");
            }
            else
            {
                sbSql.AppendLine(", SUM(PCS) [PCS]");
            }


           // sbSql.AppendLine(",SUM([PCS])[PCS]");
              
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + QuickSaleReportOBJ.Factory + "'");

            if (trading == "trading")
            {
                sbSql.AppendLine(" AND (HOYA_TRADING = 1");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD')  ");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-TRD'))  ");

            }
            else
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT') ");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT')  ");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-INT')  ");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT')  ");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT')  ");

                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT'))  ");
            }
            

            sbSql.AppendLine(" AND CUSTGROUP IN ('" + QuickSaleReportOBJ.CustomerGroup + "')");

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateTo) + "',103)");

            sbSql.AppendLine("AND INVOICEACCOUNT !='ARAF001'");
            sbSql.AppendLine("AND  ECL_SALESCOMERCIAL = 2");
            sbSql.AppendLine("GROUP BY INVOICEDATE");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getQuickSalesSupportSum(QuickSaleReportOBJ QuickSaleReportOBJ,bool Return,string trading,bool commercial)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(QuickSaleReportOBJ.DateFrom.Year, QuickSaleReportOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(QuickSaleReportOBJ.DateTo.Year, QuickSaleReportOBJ.DateTo.Month, 1);

            String strFac = QuickSaleReportOBJ.strFactory;
            if (QuickSaleReportOBJ.strFactory == "GMO")
            {

                strFac = "MO";
            }
            else
            {
                strFac = QuickSaleReportOBJ.strFactory;

            }


            sbSql.AppendLine("SELECT NUMBERSEQUENCEGROUP2 [Customer],'',SUM(PCS),SUM(THB)");


            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN NUMBERSEQUENCEGROUP='" + strFac + "-EXT' OR NUMBERSEQUENCEGROUP='" + strFac + "-INT' OR NUMBERSEQUENCEGROUP='" + strFac + "-NOC' OR NUMBERSEQUENCEGROUP='" + strFac + "-CEXT' OR NUMBERSEQUENCEGROUP='" + strFac + "-CINT' OR NUMBERSEQUENCEGROUP='" + strFac + "-TRD'  OR NUMBERSEQUENCEGROUP='" + strFac + "-CTRD'  THEN  CASE WHEN NOT(NAMEALIAS IS NULL) THEN NAMEALIAS ELSE ' TOTAL' END ELSE ");
            sbSql.AppendLine("CASE WHEN NUMBERSEQUENCEGROUP='" + strFac + "-RINT' OR NUMBERSEQUENCEGROUP='" + strFac + "-REXT' OR NUMBERSEQUENCEGROUP='" + strFac + "-RNOC' OR NUMBERSEQUENCEGROUP='" + strFac + "-RTRD' THEN 'SALE RETURN' + CASE WHEN NOT(NAMEALIAS IS NULL) THEN  NAMEALIAS ELSE ' TOTAL' END ELSE ");
            sbSql.AppendLine("NUMBERSEQUENCEGROUP END  END NUMBERSEQUENCEGROUP2");

            while (dtFrom <= dtTo)
            {

                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [LineAmountMST] ELSE 0 END)[THB]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                //sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN '' ELSE 0 END) [BAHT/PCS] ", dtFrom.Month));
                //sbSql.AppendLine(",''[BAHT/PCS]");
                dtFrom = dtFrom.AddMonths(1);


            }

            sbSql.AppendLine("FROM (");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",NAMEALIAS");

            if (QuickSaleReportOBJ.strFactory == "GMO")
            {
                sbSql.AppendLine(",SUM([SET])[PCS]");
            }
            else
            {
                sbSql.AppendLine(", SUM(PCS) [PCS]");
            }



            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
            sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + QuickSaleReportOBJ.Factory + "'");

            if (trading == "trading")
            {

                if (Return)
                {
                        sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD')  ");
                }
                else
                {
                    sbSql.AppendLine(" AND (HOYA_TRADING = 1");
                    sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CTRD') ");
                    sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-TRD')) ");
                  
                }
            }///// end Trading
            else
            {

                if (Return)
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT') ");
                    sbSql.AppendLine(" OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT'))  ");
                }
                else
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT')");
                    sbSql.AppendLine(" OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-INT')");
                    sbSql.AppendLine(" OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT') ");
                    sbSql.AppendLine(" OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT'))");
                  

                }

            }
            


            sbSql.AppendLine(" AND CUSTGROUP IN ('" + QuickSaleReportOBJ.CustomerGroup + "')");

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", QuickSaleReportOBJ.DateTo) + "',103)");
            sbSql.AppendLine("AND INVOICEACCOUNT !='ARAF001'");

            if (commercial)
            {
                sbSql.AppendLine("AND  ECL_SALESCOMERCIAL IN(1)");
            }
            else
            {
                sbSql.AppendLine("AND  ECL_SALESCOMERCIAL IN(2)");

            }

            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,ECL_SALESCOMERCIAL,INVOICEDATE");
            sbSql.AppendLine(")salesTotal");


            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS  ");
            sbSql.AppendLine("HAVING NOT NUMBERSEQUENCEGROUP IS NULL  )Data ");


            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP2 ");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

       
    }
}
