using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Globalization;

namespace NewVersion.Report
{
    class SalesSummaryDAL
    {
        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();

        string _strFac;

        public ADODB.Recordset getSaleSummaryByCustomer(DataTable dt,SalesSummaryOBJ SalesSummaryOBJ,string strSheetName ,bool salesTye)
        {

            StringBuilder sbSql = new StringBuilder();

            sbSql.AppendLine(" SELECT sales.INVOICEACCOUNT,sales.CustName,sales.ECL_Reason");

            if (salesTye){

                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(" ,sales.[" + String.Format("{0:yyMM}", dr[0]) + "]");
                }

            }else{
                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(" ,salesReturn.[" + String.Format("{0:yyMM}", dr[0]) + "] *- 1");
                }
            }


            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT,CustName,ECL_Reason");
          
            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM([" + String.Format("{0:yyMM}", dr[0]) + "]) [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,[InvoiceDate],108) [InvoiceDate]");
            sbSql.AppendLine(" ,SUM(BAHT) BAHT");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,CASE WHEN CONVERT(CHAR(4),[InvoiceDate],12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN SUM(BAHT) END [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT");
            sbSql.AppendLine(" ,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) [InvoiceDate]");
            sbSql.AppendLine(" ,CURRENCYCODEISO [Curr]");
            sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END [Baht]");
            sbSql.AppendLine(" FROM hoya_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + SalesSummaryOBJ.Factory + "'");

            if (SalesSummaryOBJ.Factory == "GMO")
            {
                _strFac = "MO";
            }
            else
            {
                _strFac = SalesSummaryOBJ.Factory;
            }

            if (strSheetName == "Trading")
            {
                 sbSql.AppendLine(" AND (HOYA_TRADING=1)");
         
            }else if(strSheetName=="Normal"){

                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' )");
       
               
                sbSql.AppendLine(" AND HOYA_TRADING=0");  

           }else if (strSheetName == "Total"){

                   sbSql.AppendLine("AND HOYA_TRADING = 1 OR");
                   sbSql.AppendLine("  (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT')");
            }

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}",SalesSummaryOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateTo) + "',103)");
           
            sbSql.AppendLine(" ) SALEBYMONTH");
            if (SalesSummaryOBJ.Factory == "GMO")
            {
                sbSql.AppendLine("WHERE SALEBYMONTH.Baht!=0");
            }


            sbSql.AppendLine(" GROUP BY INVOICEACCOUNT, CustName,ECL_Reason,[InvoiceDate]");
            sbSql.AppendLine(" ");
            sbSql.AppendLine(" ) Summary");
            sbSql.AppendLine(" GROUP BY Summary.INVOICEACCOUNT,Summary.CustName,Summary.ECL_Reason)sales ");
            
            sbSql.AppendLine("LEFT JOIN"); 
            //Leftjoin

            sbSql.AppendLine(" (SELECT INVOICEACCOUNT,CustName,ECL_Reason");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM([" + String.Format("{0:yyMM}", dr[0]) + "]) [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,[InvoiceDate],108) [InvoiceDate]");
            sbSql.AppendLine(" ,SUM(BAHT) BAHT");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,CASE WHEN CONVERT(CHAR(4),[InvoiceDate],12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN SUM(BAHT) END [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT INVOICEACCOUNT,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) [InvoiceDate]");
            sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END [Baht]");

            sbSql.AppendLine("FROM hoya_vwSalesDetail");
            sbSql.AppendLine("WHERE");

            if (strSheetName == "Trading")
            {
                sbSql.AppendLine("NUMBERSEQUENCEGROUP = '" + _strFac + "-RTRD'");

            }
            else if (strSheetName == "Normal")
            {
                sbSql.AppendLine("  (NUMBERSEQUENCEGROUP = '" + _strFac + "-REXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-RINT' )");
                sbSql.AppendLine(" AND HOYA_TRADING=0");

            }
            else if (strSheetName == "Total")
            {
                sbSql.AppendLine(" (NUMBERSEQUENCEGROUP = '" + _strFac + "-REXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-RINT' )");
            }

            sbSql.AppendLine(" ) SALEBYMONTH");
            if (SalesSummaryOBJ.Factory == "GMO")
            {
                sbSql.AppendLine("WHERE SALEBYMONTH.Baht!=0");
            }

                sbSql.AppendLine("GROUP BY INVOICEACCOUNT,Custname,ECL_Reason,[InvoiceDate]");
                sbSql.AppendLine(")Summary");
                sbSql.AppendLine("GROUP BY Summary.INVOICEACCOUNT,Summary.CustName,Summary.ECL_Reason");
                sbSql.AppendLine(")salesReturn ON sales.INVOICEACCOUNT = salesReturn.INVOICEACCOUNT");
                sbSql.AppendLine("ORDER BY INVOICEACCOUNT,Custname");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();
            object ret = null;
            rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

            return rs;

        }

        public ADODB.Recordset getSaleSummaryByCustomer2(DataTable dt, SalesSummaryOBJ SalesSummaryOBJ, string strSheetName, bool salesTye)
        {

            StringBuilder sbSql = new StringBuilder();

            sbSql.AppendLine(" SELECT sales.INVOICEACCOUNT,sales.CustName,sales.ECL_Reason");

            if (salesTye)
            {

                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(" ,sales.[" + String.Format("{0:yyMM}", dr[0]) + "]");
                }

            }
            else
            {
                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(" ,salesReturn.[" + String.Format("{0:yyMM}", dr[0]) + "] *- 1");
                }
            }


            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT,CustName,ECL_Reason");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM([" + String.Format("{0:yyMM}", dr[0]) + "]) [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,[InvoiceDate],108) [InvoiceDate]");
            sbSql.AppendLine(" ,SUM(BAHT) BAHT");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,CASE WHEN CONVERT(CHAR(4),[InvoiceDate],12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN SUM(BAHT) END [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT");
            sbSql.AppendLine(" ,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) [InvoiceDate]");
            sbSql.AppendLine(" ,CURRENCYCODEISO [Curr]");
            sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmount ELSE 0 END AmtCurr");
            sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END [Baht]");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + SalesSummaryOBJ.Factory + "'");

            if (SalesSummaryOBJ.Factory == "GMO")
            {
                _strFac = "MO";
            }
            else
            {
                _strFac = SalesSummaryOBJ.Factory;
            }

            if (strSheetName == "Trading")
            {
                sbSql.AppendLine(" AND (HOYA_TRADING=1");

            }
            else if (strSheetName == "Normal")
            {

               // sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' )");
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' )");
                sbSql.AppendLine(" AND HOYA_TRADING=0");

            }
            else if (strSheetName == "Total")
            {
                 // sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-TRD' )");
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-TRD'  )");
                sbSql.AppendLine(" AND ECL_REASON != 'A811'");

              }

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateTo) + "',103)");

            sbSql.AppendLine(" ) SALEBYMONTH");
            if (SalesSummaryOBJ.Factory == "GMO")
            {
                sbSql.AppendLine("WHERE SALEBYMONTH.Baht!=0");
            }


            sbSql.AppendLine(" GROUP BY INVOICEACCOUNT, CustName,ECL_Reason,[InvoiceDate]");
            sbSql.AppendLine(" ");
            sbSql.AppendLine(" ) Summary");
            sbSql.AppendLine(" GROUP BY Summary.INVOICEACCOUNT,Summary.CustName,Summary.ECL_Reason)sales ");

            sbSql.AppendLine("LEFT JOIN"); ///Leftjoin

            sbSql.AppendLine(" (SELECT INVOICEACCOUNT,CustName,ECL_Reason");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM([" + String.Format("{0:yyMM}", dr[0]) + "]) [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,[InvoiceDate],108) [InvoiceDate]");
            sbSql.AppendLine(" ,SUM(BAHT) BAHT");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,CASE WHEN CONVERT(CHAR(4),[InvoiceDate],12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN SUM(BAHT) END [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT");
            sbSql.AppendLine(" ,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) [InvoiceDate]");
            sbSql.AppendLine(" ,CURRENCYCODEISO [Curr]");
             sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmount ELSE 0 END AmtCurr");
             sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END [Baht]");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + SalesSummaryOBJ.Factory + "'");

            if (SalesSummaryOBJ.Factory == "GMO")
            {
                _strFac = "MO";
            }
            else
            {
                _strFac = SalesSummaryOBJ.Factory;
            }

            if (strSheetName == "Trading")
            {
                sbSql.AppendLine(" AND (HOYA_TRADING=1");

            }
            else if (strSheetName == "Normal")
            {

                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-REXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-RINT' )");
                sbSql.AppendLine(" AND HOYA_TRADING=0");

            }
            else if (strSheetName == "Total")
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-REXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-RINT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-RTRD' )");
                sbSql.AppendLine(" AND ECL_REASON != 'A811'");

            }

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateTo) + "',103)");

            sbSql.AppendLine(" ) SALEBYMONTH");
            if (SalesSummaryOBJ.Factory == "GMO")
            {
               // sbSql.AppendLine("WHERE SALEBYMONTH.Baht!=0");
            }


            sbSql.AppendLine(" GROUP BY INVOICEACCOUNT, CustName,ECL_Reason,[InvoiceDate]");
            sbSql.AppendLine(" ");
            sbSql.AppendLine(" ) Summary");
            sbSql.AppendLine(" GROUP BY Summary.INVOICEACCOUNT,Summary.CustName,Summary.ECL_Reason ");
            sbSql.AppendLine(") salesReturn ON sales.INVOICEACCOUNT = salesReturn.INVOICEACCOUNT");
            sbSql.AppendLine("ORDER BY sales.INVOICEACCOUNT");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;
        }

        public ADODB.Recordset getSaleSummaryByCustomerRP(DataTable dt, SalesSummaryOBJ SalesSummaryOBJ, string strSheetName, bool salesTye)
        {

            StringBuilder sbSql = new StringBuilder();

            sbSql.AppendLine(" SELECT INVOICEACCOUNT,CustName,ECL_Reason");

            if (salesTye)
            {

                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(" ,SUM([" + String.Format("{0:yyMM}", dr[0]) + "]) [" + String.Format("{0:yyMM}", dr[0]) + "] ");
                }

            }
            else
            {
                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(" ,SUM([" + String.Format("{0:yyMM}", dr[0]) + "]) *-1[" + String.Format("{0:yyMM}", dr[0]) + "] ");
                }

            }


            sbSql.AppendLine(" FROM (");

            sbSql.AppendLine(" SELECT INVOICEACCOUNT,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,[InvoiceDate],108) [InvoiceDate]");
            sbSql.AppendLine(" ,SUM(BAHT) BAHT");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,CASE WHEN CONVERT(CHAR(4),[InvoiceDate],12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN SUM(BAHT) END [" + String.Format("{0:yyMM}", dr[0]) + "]");

            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT INVOICEACCOUNT");
            sbSql.AppendLine(" ,CustName,ECL_Reason");
            sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) [InvoiceDate]");
            sbSql.AppendLine(" ,CURRENCYCODEISO [Curr]");
       
           sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmount ELSE 0 END AmtCurr");
           sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END [Baht]");

            sbSql.AppendLine(" FROM hoya_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + SalesSummaryOBJ.Factory + "'");

            if (SalesSummaryOBJ.Factory == "GMO")
            {
                _strFac = "MO";
            }
            else
            {
                _strFac = SalesSummaryOBJ.Factory;
            }

            if (strSheetName == "Trading")
            {
                sbSql.AppendLine(" AND HOYA_TRADING=1");
            }
            else if (strSheetName == "Normal")
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' )");
                sbSql.AppendLine(" AND HOYA_TRADING=0");
            }
            else if (strSheetName == "Total")
            {
                if (salesTye)
                {
                    //sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-TRD' )");
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-TRD'  )");
            

                }
                else
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-REXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-RINT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-RTRD' )");

                }
                
                sbSql.AppendLine(" AND ECL_REASON = 'A811'");
                //sbSql.AppendLine(" AND NOT (ECL_LENTYPE IN ('RS') AND NOT ECL_LENTYPE IN('MRS'))");
            }

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateTo) + "',103)");

            sbSql.AppendLine(" ) SALEBYMONTH");
            sbSql.AppendLine(" GROUP BY INVOICEACCOUNT, CustName,ECL_Reason,[InvoiceDate]");
            sbSql.AppendLine(" ");
            sbSql.AppendLine(" ) Summary");
            sbSql.AppendLine(" GROUP BY INVOICEACCOUNT,CustName,ECL_Reason");
            sbSql.AppendLine(" ORDER BY INVOICEACCOUNT");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSaleSummaryByCustomerRS_MRS(DataTable dt, SalesSummaryOBJ SalesSummaryOBJ, string strSheetName, bool boolRs, bool salesTye)
        {

        StringBuilder sbSql = new StringBuilder();


        if(boolRs){

            sbSql.AppendLine(" SELECT CASE WHEN ECL_LENTYPE ='MRS' THEN 'HOYA OPTICS (THAILAND) LTD. (PO FACTORY) - Glass Stick' ");
            sbSql.AppendLine(" WHEN ECL_LENTYPE ='RS' THEN 'HOYA OPTICS (THAILAND) LTD. (PO FACTORY) - Rod Slice' END [ECL_LENTYPE],''");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM([" + String.Format("{0:yyMM}", dr[0]) + "]) [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

           sbSql.AppendLine(" FROM (");

           sbSql.AppendLine(" SELECT ");
            sbSql.AppendLine(" CONVERT(DATETIME,[InvoiceDate],108) [InvoiceDate]");
            sbSql.AppendLine(" ,SUM(BAHT) BAHT");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,CASE WHEN CONVERT(CHAR(4),[InvoiceDate],12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN SUM(BAHT) END [" + String.Format("{0:yyMM}", dr[0]) + "]");

            }

            sbSql.AppendLine(" ,ECL_LENTYPE");
            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT ");
            sbSql.AppendLine(" CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) [InvoiceDate]");
            sbSql.AppendLine(" ,CURRENCYCODEISO [Curr]");
            sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmount ELSE 0 END AmtCurr");
            sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END [Baht],ECL_LENTYPE");
           sbSql.AppendLine(" FROM hoya_vwSalesDetail");
           sbSql.AppendLine(" WHERE ");
        
            sbSql.AppendLine(" InventSiteId='" + SalesSummaryOBJ.Factory + "'");

            if (SalesSummaryOBJ.Factory == "GMO")
            {
                _strFac = "MO";
            }
            else
            {
                _strFac = SalesSummaryOBJ.Factory;
            }

            if (strSheetName == "Trading")
            {
                sbSql.AppendLine(" AND HOYA_TRADING=1");

            }
            else if (strSheetName == "Normal")
            {
               // sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT')");
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT')");
            
                sbSql.AppendLine(" AND HOYA_TRADING=0");
            }
            else if (strSheetName == "Total")
            {
                //sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-TRD' )");
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-TRD'  )");
            
                sbSql.AppendLine(" AND (ECL_LENTYPE IN ('RS') OR ECL_LENTYPE IN('MRS'))");
            }

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateTo) + "',103)");

            sbSql.AppendLine(" ) SALEBYMONTH");
            sbSql.AppendLine(" GROUP BY [InvoiceDate],ECL_LENTYPE");
            sbSql.AppendLine(" ");
            sbSql.AppendLine(" ) Summary");
            sbSql.AppendLine(" GROUP BY ECL_LENTYPE");

        }
        else if (!boolRs)
        {
            sbSql.AppendLine(" SELECT '','',''");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM([" + String.Format("{0:yyMM}", dr[0]) + "]) [" + String.Format("{0:yyMM}", dr[0]) + "]");
            }

             sbSql.AppendLine(" FROM (");

            sbSql.AppendLine(" SELECT ");
            sbSql.AppendLine(" CONVERT(DATETIME,[InvoiceDate],108) [InvoiceDate]");
            sbSql.AppendLine(" ,SUM(BAHT) BAHT");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,CASE WHEN CONVERT(CHAR(4),[InvoiceDate],12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN SUM(BAHT) END [" + String.Format("{0:yyMM}", dr[0]) + "]");

            }

            sbSql.AppendLine(" ,ECL_LENTYPE");
            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT ");
            sbSql.AppendLine(" CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) [InvoiceDate]");
            sbSql.AppendLine(" ,CURRENCYCODEISO [Curr]");
            sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmount ELSE 0 END AmtCurr");
            sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END [Baht],ECL_LENTYPE");
            sbSql.AppendLine(" FROM hoya_vwSalesDetail");
            sbSql.AppendLine(" WHERE ");
            sbSql.AppendLine(" InventSiteId='" + SalesSummaryOBJ.Factory + "'");

            if (SalesSummaryOBJ.Factory == "GMO")
            {
                _strFac = "MO";
            }
            else
            {
                _strFac = SalesSummaryOBJ.Factory;
            }

            if (strSheetName == "Trading")
            {
                sbSql.AppendLine(" AND HOYA_TRADING=1");

            }
            else if (strSheetName == "Normal")
            {
               // sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' )");
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT')");
            
                sbSql.AppendLine(" AND HOYA_TRADING=0");
            }
            else if (strSheetName == "Total")
            {
                //sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-TRD' )");
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' OR  NUMBERSEQUENCEGROUP = '" + _strFac + "-TRD'  )");
            
                sbSql.AppendLine(" AND (ECL_LENTYPE IN ('RS') OR ECL_LENTYPE IN('MRS'))");
            }

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateTo) + "',103)");

            sbSql.AppendLine(" ) SALEBYMONTH");
            sbSql.AppendLine(" GROUP BY [InvoiceDate],ECL_LENTYPE");
            sbSql.AppendLine(" ");
            sbSql.AppendLine(" ) Summary");

        }

        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;
    }

        public ADODB.Recordset getSalesTradingByApplication(DataTable dtMonthRange, SalesSummaryOBJ SalesSummaryOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(" SELECT NAMEALIAS");
            DateTime dt = Convert.ToDateTime(dtMonthRange.Rows[0]["dt"]);
            for (int k = 0; k < 2; k++)
            {
                for (int j = 0; j < 2; j++)
                {

                    for (int i = 0; i < 3; i++)
                    {
                        sbSql.AppendLine(String.Format(",ISNULL(SUM([{0} K.SET]),0)[{0}K.SET]", String.Format("{0:MMM}", dt)).ToUpper());
                        sbSql.AppendLine(String.Format(",ISNULL(SUM([{0} K.PCS]),0)[{0}K.PCS]", String.Format("{0:MMM}", dt)).ToUpper());
                        sbSql.AppendLine(String.Format(",ISNULL(SUM([{0} K.Baht]),0)[{0}K.Baht]", String.Format("{0:MMM}", dt)).ToUpper());
                        dt = dt.AddMonths(1);
               

                    }
                    sbSql.AppendLine(" ,'','',''");
                }
                sbSql.AppendLine(" ,'','',''");
            }

            
        sbSql.AppendLine(" ");
        sbSql.AppendLine(" ");
        sbSql.AppendLine(" FROM ( SELECT ");
        sbSql.AppendLine(" CASE WHEN ECL_SALESCOMERCIAL=2 THEN 'ZZZZZZ' ELSE NAMEALIAS END NAMEALIAS");

        for (int i = 0; i < 12; i++)
        {
            sbSql.AppendLine(String.Format(",CASE WHEN MONTH([InvoiceDate])={0} THEN [SET]/1000 END [{1} K.SET]", dt.Month, String.Format("{0:MMM}", dt)).ToUpper());
            sbSql.AppendLine(String.Format(",CASE WHEN MONTH([InvoiceDate])={0} THEN [PCS]/1000 END [{1} K.PCS]", dt.Month, String.Format("{0:MMM}", dt)).ToUpper());
            sbSql.AppendLine(String.Format(",CASE WHEN MONTH([InvoiceDate])={0} THEN LineAmountMST/1000 END [{1} K.Baht]", dt.Month, String.Format("{0:MMM}", dt)).ToUpper());
            dt = dt.AddMonths(-1);
        }

        sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
        sbSql.AppendLine(" WHERE INVENTSITEID='" + SalesSummaryOBJ.Factory + "'");
        sbSql.AppendLine(" AND HOYA_TRADING = 1");
        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateTo) + "',103)");
        sbSql.AppendLine(" ) SaleTrading");
        sbSql.AppendLine(" GROUP BY NAMEALIAS ORDER BY NAMEALIAS");

        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;
        }

        public ADODB.Recordset getSalesSummaryByApplicationList(DataTable dtmonthRange, SalesSummaryOBJ SalesSummaryOBJ)
        {

        StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine(" SELECT CASE WHEN ECL_APPCODE='' THEN 'ZZZZZZZZ' ELSE ECL_APPCODE END ECL_APPCODE");
        sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
        sbSql.AppendLine(" WHERE ECL_SALESCOMERCIAL=1");
        sbSql.AppendLine("  AND INVENTSITEID='" + SalesSummaryOBJ.Factory + "'");
        sbSql.AppendLine("  AND HOYA_Trading=0");

        if (SalesSummaryOBJ.Factory == "GMO")
        {
            _strFac = "MO";
        }
        else
        {
            _strFac = SalesSummaryOBJ.Factory;
        }

        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT')");
        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateTo) + "',103)");
        sbSql.AppendLine("  GROUP BY ECL_APPCODE");
        sbSql.AppendLine("  ORDER BY ECL_APPCODE");

        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;

        }

        public ADODB.Recordset getSalesSummaryByApplication(DataTable dtMonthRange, SalesSummaryOBJ SalesSummaryOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine(" SELECT ");
        sbSql.AppendLine("  CASE WHEN NUMBERSEQUENCEGROUP IS NULL THEN 'TOTAL COMMERCIAL SALE' ELSE ");
        sbSql.AppendLine("      CASE WHEN INVOICEACCOUNT IS NULL THEN 'TOTAL ' ELSE '' END");
        sbSql.AppendLine(" 	    + CASE WHEN NUMBERSEQUENCEGROUP='EXT' THEN 'EXTERNAL SALE' ELSE");
        sbSql.AppendLine(" 		    CASE WHEN NUMBERSEQUENCEGROUP='INT' THEN 'INTERNAL SALE' ELSE");
        sbSql.AppendLine(" 			            NUMBERSEQUENCEGROUP");
        sbSql.AppendLine(" 		    END");
        sbSql.AppendLine(" 		END");
        sbSql.AppendLine("      + CASE WHEN NOT(NAMEALIAS IS NULL) THEN ' ('+NAMEALIAS+')' ELSE '' END");
        sbSql.AppendLine("  END NUMBERSEQUENCEGROUP2");
        sbSql.AppendLine(" ,CASE WHEN NUMBERSEQUENCEGROUP IS NULL THEN 'GRAND TOTAL' ELSE ");
        sbSql.AppendLine("      CASE WHEN (INVOICEACCOUNT IS NULL) AND NOT(NUMBERSEQUENCEGROUP IS NULL) AND GROUPING(NAMEALIAS)=0 THEN 'TOTAL' ELSE INVOICEACCOUNT END ");
        sbSql.AppendLine("  END INVOICEACCOUNT");
        sbSql.AppendLine(" ,ECL_APPCODE");

        DateTime dt = Convert.ToDateTime(dtMonthRange.Rows[0]["dt"]);
        for (int k = 0; k < 2; k++)
        {
            for (int j = 0; j < 2; j++)
            {

                for (int i = 0; i < 3; i++)
                {
                    sbSql.AppendLine(String.Format(",ISNULL(SUM([{0} K.SET]),0)[{0}K.SET]", String.Format("{0:MMM}", dt)).ToUpper());
                    sbSql.AppendLine(String.Format(",ISNULL(SUM([{0} K.PCS]),0)[{0}K.PCS]", String.Format("{0:MMM}", dt)).ToUpper());
                    sbSql.AppendLine(String.Format(",ISNULL(SUM([{0} K.Baht]),0)[{0}K.Baht]", String.Format("{0:MMM}", dt)).ToUpper());
                    dt = dt.AddMonths(1);

                }
                sbSql.AppendLine(" ,'','',''");
            }
            sbSql.AppendLine(" ,'','',''");
        }

             sbSql.AppendLine(" FROM (");
             sbSql.AppendLine("      SELECT NAMEALIAS,INVOICEACCOUNT,NUMBERSEQUENCEGROUP,ECL_APPCODE");

             for (int i = 0; i < 12; i++)
             {
                 sbSql.AppendLine(String.Format(",CASE WHEN MONTH([InvoiceDate])={0} THEN SUM([SET])/1000 END [{1} K.SET]",dt.Month,  String.Format("{0:MMM}", dt)).ToUpper());
                 sbSql.AppendLine(String.Format(",CASE WHEN MONTH([InvoiceDate])={0} THEN SUM([PCS])/1000 END [{1} K.PCS]",dt.Month, String.Format("{0:MMM}", dt)).ToUpper());
                 sbSql.AppendLine(String.Format(",CASE WHEN MONTH([InvoiceDate])={0} THEN SUM(BAHT)/1000 END [{1} K.Baht]",dt.Month, String.Format("{0:MMM}", dt)).ToUpper());
                 dt = dt.AddMonths(-1);


                 //  sbSql.AppendLine(String.Format(",ISNULL(SUM[{0} K.SET]),0)[{0}K.SET]", String.Format("{0:MM}", dt)).ToUpper());
             }

             sbSql.AppendLine("      FROM (");

        sbSql.AppendLine("    SELECT SUBSTRING(NUMBERSEQUENCEGROUP,LEN(NUMBERSEQUENCEGROUP)-2,3) NUMBERSEQUENCEGROUP");
        sbSql.AppendLine("   ,CASE WHEN ECL_APPCODE='' THEN 'ZZZZZZZZ' ELSE ECL_APPCODE END ECL_APPCODE");
        sbSql.AppendLine("   ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) [InvoiceDate]");
        sbSql.AppendLine("   ,ECL_SALESCOMERCIAL");
        sbSql.AppendLine("   ,NAMEALIAS");
        sbSql.AppendLine("   ,INVOICEACCOUNT ,INVOICEID");
        sbSql.AppendLine("   ,CUSTNAME");
        sbSql.AppendLine("   ,CURRENCYCODEISO [Curr]");
        sbSql.AppendLine("   ,ITEMID");
        sbSql.AppendLine("   ,[SET]");
        sbSql.AppendLine("   ,PCS");
        sbSql.AppendLine("   ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END [Baht]");
        sbSql.AppendLine("   FROM HOYA_vwSalesDetail");
        sbSql.AppendLine("   WHERE INVENTSITEID='" + SalesSummaryOBJ.Factory + "'");

        if (SalesSummaryOBJ.Factory == "GMO")
        {
            _strFac = "MO";

        }
        else
        {
            _strFac = SalesSummaryOBJ.Factory;

        }
        //sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT')");
        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' )");
        sbSql.AppendLine("    AND HOYA_Trading = 0");


              

        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesSummaryOBJ.DateTo) + "',103)");
         sbSql.AppendLine(" AND ECL_SALESCOMERCIAL=1");

        sbSql.AppendLine("  ) SALEBYMONTH");
        sbSql.AppendLine("  GROUP BY NUMBERSEQUENCEGROUP,INVOICEACCOUNT,NAMEALIAS,ECL_APPCODE,[InvoiceDate]");
        sbSql.AppendLine("  ) SALESUMMARY");
        sbSql.AppendLine("  GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,INVOICEACCOUNT");
        sbSql.AppendLine("  ,ECL_APPCODE WITH ROLLUP");
        sbSql.AppendLine("  HAVING INVOICEACCOUNT IS NULL OR NOT(ECL_APPCODE IS NULL)");
        sbSql.AppendLine("  ORDER BY GROUPING(NUMBERSEQUENCEGROUP),NUMBERSEQUENCEGROUP,GROUPING(NAMEALIAS),NAMEALIAS,GROUPING(INVOICEACCOUNT),INVOICEACCOUNT,ECL_APPCODE");

        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;
             
        }

         public ADODB.Recordset getSalesResultBySalesGroup(DateTime dt, SalesSummaryOBJ SalesSummaryOBJ, bool boolTotal)
        {
            StringBuilder sbSql = new StringBuilder();


            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN Summary.NUMBERSEQUENCEGROUP IS NULL THEN 'TOTAL COMMERCIAL SALE' ELSE  ");
            sbSql.AppendLine("      CASE WHEN Summary.INVOICEACCOUNT IS NULL THEN 'TOTAL ' ELSE '' END");
            sbSql.AppendLine("   + CASE WHEN Summary.NUMBERSEQUENCEGROUP='EXT' THEN 'EXTERNAL SALE' ELSE");
            sbSql.AppendLine("CASE WHEN Summary.NUMBERSEQUENCEGROUP='INT' THEN 'INTERNAL SALE' ELSE");
            sbSql.AppendLine("Summary.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("END");
            sbSql.AppendLine("END");
            sbSql.AppendLine(" + CASE WHEN NOT(Summary.NAMEALIAS IS NULL) THEN ' ('+Summary.NAMEALIAS+')' ELSE '' END");
            sbSql.AppendLine(" END NUMBERSEQUENCEGROUP2");
            sbSql.AppendLine(" ,CASE WHEN Summary.NUMBERSEQUENCEGROUP IS NULL THEN 'GRAND TOTAL' ELSE  ");
            sbSql.AppendLine("CASE WHEN (Summary.INVOICEACCOUNT IS NULL) AND NOT(Summary.NUMBERSEQUENCEGROUP IS NULL) AND GROUPING(Summary.NAMEALIAS)=0 THEN 'TOTAL' ELSE Summary.INVOICEACCOUNT END  ");
            sbSql.AppendLine("  END INVOICEACCOUNT");
            sbSql.AppendLine(",Summary.ECL_GROUPCODE,'',''");
           

            if (!boolTotal)
            {
                sbSql.AppendLine(",SUM(Summary.[K.SETSALE]) [K.SET]");
                sbSql.AppendLine(",SUM(Summary.[K.PCSSALE]) [K.PCS]");
                sbSql.AppendLine(" ,SUM(Summary.[K.BAHTSALE]) [K.BAHT]");
                sbSql.AppendLine(",''");
                sbSql.AppendLine("  ,SUM(Summary.[K.SETRETURN]) [K.SET]");
                sbSql.AppendLine(",SUM(Summary.[K.PCSRETURN]) [K.PCS]");
                sbSql.AppendLine(",SUM(Summary.[K.BAHTRETURN]) [K.BAHT]");
            }


            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN SALES.NUMBERSEQUENCEGROUP IS NULL THEN SALESRETURN.ReturnNumSequen ELSE SALES.NUMBERSEQUENCEGROUP END  [NUMBERSEQUENCEGROUP]");
            sbSql.AppendLine(",CASE WHEN SALES.INVOICEACCOUNT IS NULL THEN SALESRETURN.ReturnInvoiceAccount ELSE SALES.INVOICEACCOUNT END [INVOICEACCOUNT]");
            sbSql.AppendLine(",CASE WHEN SALES.ECL_GROUPCODE IS NULL THEN SALESRETURN.ReturnEcl_groupCode ELSE SALES.ECL_GROUPCODE END [ECL_GROUPCODE]");
            sbSql.AppendLine(",CASE WHEN SALES.NAMEALIAS IS NULL THEN SALESRETURN.ReturnNAMEALIAS ELSE SALES.NAMEALIAS END [NAMEALIAS]");
            sbSql.AppendLine(",SALES.[K.SETSALE],SALES.[K.PCSSALE],SALES.[K.BAHTSALE]");
            sbSql.AppendLine(",SALESRETURN.[K.SETRETURN],SALESRETURN.[K.PCSRETURN],SALESRETURN.[K.BAHTRETURN]");

            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT ");
            sbSql.AppendLine("NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",INVOICEACCOUNT");
            sbSql.AppendLine(",ECL_GROUPCODE");
            sbSql.AppendLine(",NAMEALIAS");
            sbSql.AppendLine(",SUM([SET])/1000 [K.SETSALE]");
            sbSql.AppendLine(" ,SUM([PCS])/1000 [K.PCSSALE]");
            sbSql.AppendLine(" ,SUM([BAHT])/1000 [K.BAHTSALE]");

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine("SELECT SUBSTRING(NUMBERSEQUENCEGROUP,LEN(NUMBERSEQUENCEGROUP)-2,3) NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("  ,CASE WHEN ECL_GROUPCODE='OTH' THEN 'ZZZZZZ' ELSE ECL_GROUPCODE END AS ECL_GROUPCODE");
            sbSql.AppendLine("  ,NAMEALIAS");
            sbSql.AppendLine(" ,INVOICEACCOUNT");
            sbSql.AppendLine("  ,SUM(LineAmountMST) [Baht]");
            sbSql.AppendLine(" ,SUM([SET])[SET]");
            sbSql.AppendLine("  ,SUM(PCS)[PCS]");
            sbSql.AppendLine("  FROM HOYA_vwSalesDetail");

            if (SalesSummaryOBJ.Factory == "GMO")
            {
                _strFac = "MO";
            }
            else
            {
                _strFac = SalesSummaryOBJ.Factory;

            }
            sbSql.AppendLine("  WHERE INVENTSITEID='" + SalesSummaryOBJ.Factory + "'");
            sbSql.AppendLine("   AND HOYA_TRADING=0");
            sbSql.AppendLine("   AND ECL_SALESCOMERCIAL=1");


            //sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT')");
            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' )");


            if (boolTotal)
            {
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year, 4, 1)) + "',103) ");
                sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year + 1, 3, 31)) + "',103)");

            }
            else
            {
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103) ");
                sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt.AddMonths(1).AddDays(-1)) + "',103)");

            }


            sbSql.AppendLine(" GROUP BY NUMBERSEQUENCEGROUP,ECL_GROUPCODE,NAMEALIAS,INVOICEACCOUNT)S2");
            sbSql.AppendLine(" GROUP BY NUMBERSEQUENCEGROUP,ECL_GROUPCODE,NAMEALIAS,INVOICEACCOUNT) as SALES");

            sbSql.AppendLine("  FULL JOIN ");

            sbSql.AppendLine(" (SELECT ");
            sbSql.AppendLine(" NUMBERSEQUENCEGROUP [ReturnNumSequen]");
            sbSql.AppendLine(",INVOICEACCOUNT [ReturnInvoiceAccount]");
            sbSql.AppendLine(" ,ECL_GROUPCODE [ReturnEcl_groupCode]");
            sbSql.AppendLine(",NAMEALIAS ReturnNAMEALIAS");
            sbSql.AppendLine(" ,SUM([SET])/1000 [K.SETRETURN]");
            sbSql.AppendLine(" ,SUM([PCS])/1000 [K.PCSRETURN]");
            sbSql.AppendLine(" ,SUM([BAHT])/1000 [K.BAHTRETURN]");
          
             sbSql.AppendLine("FROM (");
            sbSql.AppendLine(" SELECT SUBSTRING(NUMBERSEQUENCEGROUP,LEN(NUMBERSEQUENCEGROUP)-2,3) NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",CASE WHEN ECL_GROUPCODE='OTH' THEN 'ZZZZZZ' ELSE ECL_GROUPCODE END AS ECL_GROUPCODE");
            sbSql.AppendLine(",NAMEALIAS");
            sbSql.AppendLine(",INVOICEACCOUNT");
            sbSql.AppendLine(",SUM(LineAmountMST) [Baht]");
            sbSql.AppendLine(",SUM([SET])[SET]");
            sbSql.AppendLine(",SUM(PCS)[PCS]");

            sbSql.AppendLine("FROM HOYA_vwSalesDetail");

            if (SalesSummaryOBJ.Factory == "GMO")
            {
                _strFac = "MO";
            }
            else
            {
                _strFac = SalesSummaryOBJ.Factory;

            }
            sbSql.AppendLine("  WHERE INVENTSITEID='" + SalesSummaryOBJ.Factory + "'");

            sbSql.AppendLine("AND HOYA_TRADING=0");
            sbSql.AppendLine(" AND ECL_SALESCOMERCIAL=1");
            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-REXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-RINT' )");

            if (boolTotal)
            {
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year, 4, 1)) + "',103) ");
                sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year + 1, 3, 31)) + "',103)");

            }
            else
            {
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103) ");
                sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt.AddMonths(1).AddDays(-1)) + "',103)");

            }


            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,ECL_GROUPCODE,NAMEALIAS,INVOICEACCOUNT)S2");
            sbSql.AppendLine(" GROUP BY NUMBERSEQUENCEGROUP,ECL_GROUPCODE,NAMEALIAS,INVOICEACCOUNT)SALESRETURN");


            sbSql.AppendLine(" ON SALES.INVOICEACCOUNT=SALESRETURN.ReturnInvoiceAccount");
            sbSql.AppendLine("  AND SALES.ECL_GROUPCODE=SALESRETURN.ReturnEcl_groupCode) AS Summary");
            sbSql.AppendLine("GROUP BY Summary.NUMBERSEQUENCEGROUP,Summary.NAMEALIAS, Summary.INVOICEACCOUNT,Summary.ECL_GROUPCODE WITH ROLLUP");
            sbSql.AppendLine("HAVING Summary.INVOICEACCOUNT IS NULL OR NOT(Summary.ECL_GROUPCODE IS NULL)");
            sbSql.AppendLine("ORDER BY GROUPING(NUMBERSEQUENCEGROUP),NUMBERSEQUENCEGROUP,GROUPING(Summary.NAMEALIAS),Summary.NAMEALIAS,GROUPING(Summary.INVOICEACCOUNT),INVOICEACCOUNT,ECL_GROUPCODE");






            /*
             sbSql.AppendLine(" SELECT ");
             sbSql.AppendLine("   CASE WHEN SALES.NUMBERSEQUENCEGROUP IS NULL THEN 'TOTAL COMMERCIAL SALE' ELSE  ");
             sbSql.AppendLine("      CASE WHEN SALES.INVOICEACCOUNT IS NULL THEN 'TOTAL ' ELSE '' END");
             sbSql.AppendLine(" 	    + CASE WHEN SALES.NUMBERSEQUENCEGROUP='EXT' THEN 'EXTERNAL SALE' ELSE");
             sbSql.AppendLine(" 		    CASE WHEN SALES.NUMBERSEQUENCEGROUP='INT' THEN 'INTERNAL SALE' ELSE");
             sbSql.AppendLine(" 			           SALES.NUMBERSEQUENCEGROUP");
             sbSql.AppendLine(" 		    END");
             sbSql.AppendLine(" 		END");
             sbSql.AppendLine("     + CASE WHEN NOT(SALES.NAMEALIAS IS NULL) THEN ' ('+SALES.NAMEALIAS+')' ELSE '' END");
             sbSql.AppendLine("  END NUMBERSEQUENCEGROUP2");
             sbSql.AppendLine("  ,CASE WHEN SALES.NUMBERSEQUENCEGROUP IS NULL THEN 'GRAND TOTAL' ELSE  ");
             sbSql.AppendLine("       CASE WHEN (SALES.INVOICEACCOUNT IS NULL) AND NOT(SALES.NUMBERSEQUENCEGROUP IS NULL) AND GROUPING(SALES.NAMEALIAS)=0 THEN 'TOTAL' ELSE SALES.INVOICEACCOUNT END  ");
             sbSql.AppendLine("  END INVOICEACCOUNT");
             sbSql.AppendLine("  ,SALES.ECL_GROUPCODE,'',''");

             if(!boolTotal){
                 sbSql.AppendLine(" ,SUM(SALES.[K.SETSALE]) [K.SET]");
                 sbSql.AppendLine("  ,SUM(SALES.[K.PCSSALE]) [K.PCS]");
                 sbSql.AppendLine(" ,SUM(SALES.[K.BAHTSALE]) [K.BAHT]");
                 sbSql.AppendLine(",''");
                 sbSql.AppendLine("  ,SUM(SALESRETURN.[K.SETRETURN]) [K.SET]");
                 sbSql.AppendLine(" ,SUM(SALESRETURN.[K.PCSRETURN]) [K.PCS]");
                 sbSql.AppendLine("  ,SUM(SALESRETURN.[K.BAHTRETURN]) [K.BAHT]");
             }

             sbSql.AppendLine(" FROM (");
             sbSql.AppendLine("  SELECT ");
             sbSql.AppendLine("  NUMBERSEQUENCEGROUP ");
             sbSql.AppendLine("  ,INVOICEACCOUNT");
             sbSql.AppendLine("  ,ECL_GROUPCODE");
             sbSql.AppendLine(" ,NAMEALIAS");
             sbSql.AppendLine(" ,SUM([SET])/1000 [K.SETSALE]");
             sbSql.AppendLine(" ,SUM([PCS])/1000 [K.PCSSALE]");
             sbSql.AppendLine(" ,SUM([BAHT])/1000 [K.BAHTSALE]");

             sbSql.AppendLine(" FROM (");
             sbSql.AppendLine("SELECT SUBSTRING(NUMBERSEQUENCEGROUP,LEN(NUMBERSEQUENCEGROUP)-2,3) NUMBERSEQUENCEGROUP");
             sbSql.AppendLine("  ,CASE WHEN ECL_GROUPCODE='OTH' THEN 'ZZZZZZ' ELSE ECL_GROUPCODE END AS ECL_GROUPCODE");
             sbSql.AppendLine("  ,NAMEALIAS");
             sbSql.AppendLine(" ,INVOICEACCOUNT");
             sbSql.AppendLine("  ,SUM(LineAmountMST) [Baht]");
             sbSql.AppendLine(" ,SUM([SET])[SET]");
             sbSql.AppendLine("  ,SUM(PCS)[PCS]");
             sbSql.AppendLine("  FROM HOYA_vwSalesDetail");

             if (SalesSummaryOBJ.Factory == "GMO")
             {
                 _strFac = "MO";
             }
             else
             {
                 _strFac = SalesSummaryOBJ.Factory;

             }
             sbSql.AppendLine("  WHERE INVENTSITEID='" + SalesSummaryOBJ.Factory + "'");
             sbSql.AppendLine("   AND HOYA_TRADING=0");
             sbSql.AppendLine("   AND ECL_SALESCOMERCIAL=1");


             //sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT')");
             sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' )");


             if (boolTotal)
             {
                 sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year, 4, 1)) + "',103) ");
                 sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year + 1, 3, 31)) + "',103)");

             }
             else
             {
                 sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103) ");
                 sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt.AddMonths(1).AddDays(-1)) + "',103)");

             }


             sbSql.AppendLine(" GROUP BY NUMBERSEQUENCEGROUP,ECL_GROUPCODE,NAMEALIAS,INVOICEACCOUNT)S2");
             sbSql.AppendLine(" GROUP BY NUMBERSEQUENCEGROUP,ECL_GROUPCODE,NAMEALIAS,INVOICEACCOUNT) as SALES");

             sbSql.AppendLine("  LEFT JOIN ");
             sbSql.AppendLine(" (SELECT ");
             sbSql.AppendLine(" NUMBERSEQUENCEGROUP [ReturnNumSequen]");
             sbSql.AppendLine(",INVOICEACCOUNT [ReturnInvoiceAccount]");
             sbSql.AppendLine(" ,ECL_GROUPCODE [ReturnEcl_groupCode]");
             sbSql.AppendLine(",'' NAMEALIAS");
             sbSql.AppendLine(" ,SUM([SET])/1000 [K.SETRETURN]");
             sbSql.AppendLine(" ,SUM([PCS])/1000 [K.PCSRETURN]");
             sbSql.AppendLine(" ,SUM([BAHT])/1000 [K.BAHTRETURN]");
             sbSql.AppendLine("FROM (");
             sbSql.AppendLine(" SELECT SUBSTRING(NUMBERSEQUENCEGROUP,LEN(NUMBERSEQUENCEGROUP)-2,3) NUMBERSEQUENCEGROUP");
             sbSql.AppendLine(",CASE WHEN ECL_GROUPCODE='OTH' THEN 'ZZZZZZ' ELSE ECL_GROUPCODE END AS ECL_GROUPCODE");
             sbSql.AppendLine(",NAMEALIAS");
             sbSql.AppendLine(",INVOICEACCOUNT");
             sbSql.AppendLine(",SUM(LineAmountMST) [Baht]");
             sbSql.AppendLine(",SUM([SET])[SET]");
             sbSql.AppendLine(",SUM(PCS)[PCS]");
             sbSql.AppendLine("FROM HOYA_vwSalesDetail");

             if (SalesSummaryOBJ.Factory == "GMO")
             {
                 _strFac = "MO";
             }
             else
             {
                 _strFac = SalesSummaryOBJ.Factory;

             }
             sbSql.AppendLine("  WHERE INVENTSITEID='" + SalesSummaryOBJ.Factory + "'");

             sbSql.AppendLine("AND HOYA_TRADING=0");
             sbSql.AppendLine(" AND ECL_SALESCOMERCIAL=1");
             sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-REXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-RINT' )");

             if (boolTotal)
             {
                 sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year, 4, 1)) + "',103) ");
                 sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year + 1, 3, 31)) + "',103)");

             }
             else
             {
                 sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103) ");
                 sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt.AddMonths(1).AddDays(-1)) + "',103)");

             }


             sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,ECL_GROUPCODE,NAMEALIAS,INVOICEACCOUNT)S2");
             sbSql.AppendLine(" GROUP BY NUMBERSEQUENCEGROUP,ECL_GROUPCODE,NAMEALIAS,INVOICEACCOUNT)SALESRETURN");
      
             sbSql.AppendLine(" ON SALES.INVOICEACCOUNT=SALESRETURN.ReturnInvoiceAccount");
             sbSql.AppendLine("AND SALES.ECL_GROUPCODE=SALESRETURN.ReturnEcl_groupCode");
             sbSql.AppendLine("GROUP BY SALES.NUMBERSEQUENCEGROUP,SALES.NAMEALIAS, SALES.INVOICEACCOUNT");
             sbSql.AppendLine(",SALES.ECL_GROUPCODE WITH ROLLUP");
             sbSql.AppendLine("HAVING SALES.INVOICEACCOUNT IS NULL OR NOT(SALES.ECL_GROUPCODE IS NULL)");
             sbSql.AppendLine("ORDER BY GROUPING(NUMBERSEQUENCEGROUP),NUMBERSEQUENCEGROUP,GROUPING(SALES.NAMEALIAS),SALES.NAMEALIAS,GROUPING(SALES.INVOICEACCOUNT),INVOICEACCOUNT,ECL_GROUPCODE");
                    */

       ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;

        }

        public ADODB.Recordset getSalesResultNoCOM(DateTime dt, SalesSummaryOBJ SalesSummaryOBJ,bool SalesType)
        {
            StringBuilder sbSql = new StringBuilder();


        sbSql.AppendLine(" SELECT SUM([SET])/1000 [K.SET]");
        sbSql.AppendLine(" ,SUM(PCS)/1000 [K.PCS]");
        sbSql.AppendLine(" ,0 [K.BAHT]");
        sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
        sbSql.AppendLine(" WHERE ECL_SALESCOMERCIAL=2"); //1=COM, 2=NOCOM
        sbSql.AppendLine(" AND HOYA_TRADING=0");

        sbSql.AppendLine("   AND INVENTSITEID='" + SalesSummaryOBJ.Factory + "'");
        if(SalesSummaryOBJ.Factory == "GMO" ){
            _strFac = "MO";
        }else{
            _strFac = SalesSummaryOBJ.Factory;
        }

        if (SalesType)
        {

          //  sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT')");
            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' )");
    
        }
        else
        {
            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + _strFac + "-REXT" + "')");
            sbSql.AppendLine(" OR  NUMBERSEQUENCEGROUP = ('" + _strFac + "-RINT" + "'))");
            sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP != ('" + _strFac + "-RNOC" + "')");
         

        }

        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103) ");
        sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt.AddMonths(1).AddDays(-1)) + "',103)");

      
        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;

        }

        public ADODB.Recordset getSalesResultTrading(DateTime dt,SalesSummaryOBJ SalesSummryOBJ, int intCom,bool SalesType)
        {
            StringBuilder sbSql = new StringBuilder();


        sbSql.AppendLine(" SELECT SUM([SET])/1000 [K.SET]");
        sbSql.AppendLine(" ,SUM(PCS)/1000 [K.PCS]");
        sbSql.AppendLine(" ,SUM(CASE WHEN ECL_SALESCOMERCIAL=0 THEN 0 ELSE LineAmountMST/1000 END) [K.BAHT]");
        sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
        sbSql.AppendLine(" WHERE ECL_SALESCOMERCIAL=" + intCom); //1=COM, 2=NOCOM
 

        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103) ");
        sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt.AddMonths(1).AddDays(-1)) + "',103)");

        sbSql.AppendLine("AND INVENTSITEID = '"+ SalesSummryOBJ.Factory +"'");

        if (SalesSummryOBJ.Factory == "GMO")
        {
            _strFac = "MO";
        }
        else
        {
            _strFac = SalesSummryOBJ.Factory;
        }

        if (SalesType)
        {

            sbSql.AppendLine(" AND (HOYA_TRADING=1)");
        }
        else
        {
            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-RTRD')");
        }

        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103) ");
        sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt.AddMonths(1).AddDays(-1)) + "',103)");


        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;

        }

        public ADODB.Recordset getSalesResultByCustomer(DateTime dt, SalesSummaryOBJ SalesSummaryOBJ, bool boolTrading, string strReport)
        {
            StringBuilder sbSql = new StringBuilder();


        sbSql.AppendLine(" SELECT");
        sbSql.AppendLine("sales.CustName,sales.INVOICEACCOUNT,sales.ECL_REASON,sales.NUMBERSEQUENCEGROUP,sales.CustGroup");

        if( strReport == "ByMonth"){
            sbSql.AppendLine(",sales.[K.SET],sales.[K.PCS],sales.[K.BAHT],'',salesReturn.[K.SET],salesReturn.[K.PCS],salesReturn.[K.BAHT]");
        }

        sbSql.AppendLine("FROM(");
        sbSql.AppendLine("SELECT CustName,INVOICEACCOUNT,ECL_REASON,'' NUMBERSEQUENCEGROUP,'' CustGroup");


        if (strReport == "ByMonth")
        {
            sbSql.AppendLine(" ,SUM([SET])/1000 [K.SET],SUM(PCS)/1000 [K.PCS]");
            sbSql.AppendLine(" ,SUM(CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END)/1000 [K.BAHT]");
        }


        sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
        sbSql.AppendLine(" WHERE InventSiteId='" + SalesSummaryOBJ.Factory + "'");
        sbSql.AppendLine(" AND HOYA_TRADING=0");

        if(SalesSummaryOBJ.Factory == "GMO"){
            _strFac = "MO";
        } else{
            _strFac = SalesSummaryOBJ.Factory;
          }


       // sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" +_strFac+ "-INT')");
        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-EXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-INT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CINT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-CEXT' )");
    

        if (strReport=="Total")
        {
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year, 4, 1)) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year + 1, 3, 31)) + "',103)");

        }
        else
        {
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt.AddMonths(1).AddDays(-1)) + "',103)");

        }

 
        sbSql.AppendLine(" GROUP BY CustName,INVOICEACCOUNT,ECL_REASON)sales");

        sbSql.AppendLine("LEFT JOIN");

        sbSql.AppendLine("(");
        sbSql.AppendLine("SELECT CustName,INVOICEACCOUNT,ECL_REASON,'' NUMBERSEQUENCEGROUP,'' CustGroup");


        if (strReport == "ByMonth")
        {
            sbSql.AppendLine(" ,SUM([SET])/1000 [K.SET],SUM(PCS)/1000 [K.PCS]");
            sbSql.AppendLine(" ,SUM(CASE WHEN ECL_SALESCOMERCIAL=1 THEN LineAmountMST ELSE 0 END)/1000 [K.BAHT]");
        }

        sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
        sbSql.AppendLine(" WHERE InventSiteId='" + SalesSummaryOBJ.Factory + "'");
        sbSql.AppendLine(" AND HOYA_TRADING=0");

        if (SalesSummaryOBJ.Factory == "GMO")
        {
            _strFac = "MO";
        }
        else
        {
            _strFac = SalesSummaryOBJ.Factory;
        }


        sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = '" + _strFac + "-REXT' OR NUMBERSEQUENCEGROUP = '" + _strFac + "-RINT')");

        if (strReport == "Total")
        {
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year, 4, 1)) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", new DateTime(dt.Year + 1, 3, 31)) + "',103)");

        }
        else
        {
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt.AddMonths(1).AddDays(-1)) + "',103)");

        }


        sbSql.AppendLine(" GROUP BY CustName,INVOICEACCOUNT,ECL_REASON)salesReturn");
        sbSql.AppendLine("ON sales.INVOICEACCOUNT = salesReturn.INVOICEACCOUNT");

              if (strReport == "ByMonth")
        {
            sbSql.AppendLine("WHERE sales.[K.BAHT] ! = 0 ");
        }


            
        sbSql.AppendLine(" ORDER BY sales.INVOICEACCOUNT,sales.ECL_REASON");

      ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;
        }


    }//end class
}
