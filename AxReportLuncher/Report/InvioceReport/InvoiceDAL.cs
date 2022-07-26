using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.Report.InvioceReport
{
    class InvoiceDAL
    {
        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();

        public DataTable getNumberSequenceGroup(string strFactory, int intShipmentLocation)
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(" SELECT DISTINCT NUMBERSEQUENCEGROUPID");
            sbSql.AppendLine(" FROM ECL_SalesImportSetup");

           sbSql.AppendLine(" WHERE InventSiteID='" + strFactory + "' AND ShipmentLoc=" + intShipmentLocation);

            
            
            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public DataTable getCustomerGroup()
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(" SELECT NUMBERSEQUENCEGROUP2 [Customer],THB,PCS");
            sbSql.AppendLine(" ORDER BY CustGroup");


            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public ADODB.Recordset getSummaryByItem(InvoiceOBJ InvoiceOBJ,bool external)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }

            if (external)
            {

                sbSql.AppendLine(" SELECT ");
                sbSql.AppendLine("  CASE WHEN ECL_GROUPCODE IS NULL  THEN 'TOTAL GROSS SALES - EXTERNAL' ELSE ECL_GROUPCODE END [CUSTOMER]");
                sbSql.AppendLine(" ,CASE WHEN CURRENCY IS NULL AND NOT(ECL_GROUPCODE IS NULL) THEN 'TOTAL' ELSE CURRENCY END CURRENCY");


                while (dtFrom <= dtTo)
                {

                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Amount Cur]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Baht/SET]", dtFrom.Month));
                    dtFrom = dtFrom.AddMonths(1);
                }

                sbSql.AppendLine(" FROM (");
                sbSql.AppendLine(" SELECT ECL_GROUPCODE,CURRENCYCODEISO [CURRENCY]");
                sbSql.AppendLine(" ,[SET]");
                sbSql.AppendLine(" ,[PCS]");
                sbSql.AppendLine(" ,LINEAMOUNT [Amount Cur]");
                sbSql.AppendLine(" ,SalesPrice");
                sbSql.AppendLine(" ,LineAmountMST [Amount Baht]");
                sbSql.AppendLine(" ,SalesPrice*EXCHRATE [Baht/PC]");
                sbSql.AppendLine(" ,InvoiceDate");
                sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
                sbSql.AppendLine(" ,ECL_SalesComercial");
                sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
                sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
                sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
                sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");


                //sbSql.AppendLine("AND (NUMBERSEQUENCEGROUP LIKE '%-CE%' OR NUMBERSEQUENCEGROUP LIKE '%-EXT%')");   //EXT CEXT
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT')"); //CEXT EXT INT CINT
                  
                sbSql.AppendLine(" AND ECL_SalesComercial=1 AND HOYA_TRADING = 0");

                sbSql.AppendLine(" ) ItemDetail");
                sbSql.AppendLine("GROUP BY [ECL_GROUPCODE],CURRENCY,ECL_SalesComercial WITH ROLLUP");
                sbSql.AppendLine(" HAVING NOT(ECL_SalesComercial IS NULL) OR [CURRENCY] IS NULL");
                sbSql.AppendLine(" ORDER BY GROUPING(ECL_GROUPCODE),ECL_GROUPCODE,GROUPING(CURRENCY),Currency");

            }
            else
            {

                sbSql.AppendLine(" SELECT ");
                sbSql.AppendLine("  CASE WHEN ECL_GROUPCODE IS NULL  THEN 'TOTAL GROSS SALES - INTERNAL' ELSE ECL_GROUPCODE END [CUSTOMER]");
                sbSql.AppendLine(" ,CASE WHEN CURRENCY IS NULL AND NOT(ECL_GROUPCODE IS NULL) THEN 'TOTAL' ELSE CURRENCY END CURRENCY");



                while (dtFrom <= dtTo)
                {

                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Amount Cur]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Baht/SET]", dtFrom.Month));
                    dtFrom = dtFrom.AddMonths(1);
                }


                sbSql.AppendLine(" FROM (");
                sbSql.AppendLine(" SELECT ECL_GROUPCODE,CURRENCYCODEISO [CURRENCY]");
                sbSql.AppendLine(" ,[SET]");
                sbSql.AppendLine(" ,[PCS]");
                sbSql.AppendLine(" ,LINEAMOUNT [Amount Cur]");
                sbSql.AppendLine(" ,SalesPrice");
                sbSql.AppendLine(" ,LineAmountMST [Amount Baht]");
                sbSql.AppendLine(" ,SalesPrice*EXCHRATE [Baht/PC]");
                sbSql.AppendLine(" ,InvoiceDate");
                sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
                sbSql.AppendLine(" ,ECL_SalesComercial");
                sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
                sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
                sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
                sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");


   
              
                  sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CINT')"); // INT CINT


                sbSql.AppendLine(" AND ECL_SalesComercial=1 AND HOYA_TRADING = 0");
                sbSql.AppendLine(" ) ItemDetail");
                sbSql.AppendLine("GROUP BY [ECL_GROUPCODE],CURRENCY,ECL_SalesComercial WITH ROLLUP");
                sbSql.AppendLine(" HAVING NOT(ECL_SalesComercial IS NULL) OR [CURRENCY] IS NULL");
                //sbSql.AppendLine(" HAVING NOT(ECL_SalesComercial IS NULL) ");
                sbSql.AppendLine(" ORDER BY GROUPING(ECL_GROUPCODE),ECL_GROUPCODE,GROUPING(CURRENCY),Currency");

            }//end comercial



            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSummaryByItemNocome(InvoiceOBJ InvoiceOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }



                 sbSql.AppendLine(" SELECT 'SALES NO COME','' ");
                //sbSql.AppendLine("  CASE WHEN ECL_GROUPCODE IS NULL  THEN 'TOTAL GROSS SALES - EXTERNAL' ELSE ECL_GROUPCODE END [CUSTOMER]");
                //sbSql.AppendLine(" ,CASE WHEN CURRENCY IS NULL AND NOT(ECL_GROUPCODE IS NULL) THEN 'TOTAL' ELSE CURRENCY END CURRENCY");


                while (dtFrom <= dtTo)
                {

                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    dtFrom = dtFrom.AddMonths(1);
                }

                sbSql.AppendLine(" FROM (");
                sbSql.AppendLine(" SELECT ECL_GROUPCODE,CURRENCYCODEISO [CURRENCY]");
                sbSql.AppendLine(" ,[SET]");
                sbSql.AppendLine(" ,[PCS]");
                sbSql.AppendLine(" ,LINEAMOUNT [Amount Cur]");
                sbSql.AppendLine(" ,SalesPrice");
                sbSql.AppendLine(" ,LineAmountMST [Amount Baht]");
                sbSql.AppendLine(" ,SalesPrice*EXCHRATE [Baht/PC]");
                sbSql.AppendLine(" ,InvoiceDate");
                sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
                sbSql.AppendLine(" ,ECL_SalesComercial");
                sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
                sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
                sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
                sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");

                //sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-" + "'))");
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT" + "') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-INT" + "')  ");
                sbSql.AppendLine(" OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT" + "') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT" + "'))  ");

                sbSql.AppendLine(" AND ECL_SalesComercial=2 ");

                sbSql.AppendLine(" ) ItemDetail");
                

            
           

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSummaryByItemReturn(InvoiceOBJ InvoiceOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }



            sbSql.AppendLine(" SELECT 'SALE RETURN','' ");
            //sbSql.AppendLine("  CASE WHEN ECL_GROUPCODE IS NULL  THEN 'TOTAL GROSS SALES - EXTERNAL' ELSE ECL_GROUPCODE END [CUSTOMER]");
            //sbSql.AppendLine(" ,CASE WHEN CURRENCY IS NULL AND NOT(ECL_GROUPCODE IS NULL) THEN 'TOTAL' ELSE CURRENCY END CURRENCY");


            while (dtFrom <= dtTo)
            {

            sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
            sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
            sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Amount Cur]", dtFrom.Month));
            sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/PCS]", dtFrom.Month));
            sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/SET]", dtFrom.Month));
            sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
            sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
            sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Baht/SET]", dtFrom.Month));
         
                dtFrom = dtFrom.AddMonths(1);
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT ECL_GROUPCODE,CURRENCYCODEISO [CURRENCY]");
            sbSql.AppendLine(" ,[SET]");
            sbSql.AppendLine(" ,[PCS]");
            sbSql.AppendLine(" ,LINEAMOUNT [Amount Cur]");
            sbSql.AppendLine(" ,SalesPrice");
            sbSql.AppendLine(" ,LineAmountMST [Amount Baht]");
            sbSql.AppendLine(" ,SalesPrice*EXCHRATE [Baht/PC]");
            sbSql.AppendLine(" ,InvoiceDate");
            sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" ,ECL_SalesComercial");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");

            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT" + "') OR NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT" + "'))");
            sbSql.AppendLine(" AND ECL_SalesComercial=1 AND HOYA_TRADING = 0");

            sbSql.AppendLine(" ) ItemDetail");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSummaryByItemTrading(InvoiceOBJ InvoiceOBJ,bool Return)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }



            sbSql.AppendLine(" SELECT '' ");
            //sbSql.AppendLine("  CASE WHEN ECL_GROUPCODE IS NULL  THEN 'TOTAL GROSS SALES - EXTERNAL' ELSE ECL_GROUPCODE END [CUSTOMER]");
            //sbSql.AppendLine(" ,CASE WHEN CURRENCY IS NULL AND NOT(ECL_GROUPCODE IS NULL) THEN 'TOTAL' ELSE CURRENCY END CURRENCY");


            while (dtFrom <= dtTo)
            {

                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Amount Cur]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/PCS]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/SET]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Baht/SET]", dtFrom.Month));

                dtFrom = dtFrom.AddMonths(1);
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT ECL_GROUPCODE,CURRENCYCODEISO [CURRENCY]");
            sbSql.AppendLine(" ,[SET]");
            sbSql.AppendLine(" ,[PCS]");
            sbSql.AppendLine(" ,LINEAMOUNT [Amount Cur]");
            sbSql.AppendLine(" ,SalesPrice");
            sbSql.AppendLine(" ,LineAmountMST [Amount Baht]");
            sbSql.AppendLine(" ,SalesPrice*EXCHRATE [Baht/PC]");
            sbSql.AppendLine(" ,InvoiceDate");
            sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" ,ECL_SalesComercial");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");

       
            sbSql.AppendLine(" AND ECL_SalesComercial=1");

            if (Return)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD" + "'))");
            }
            else
            {
                sbSql.AppendLine(" AND HOYA_TRADING = 1");
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP != ('" + strFac + "-RTRD" + "'))");

            }




            sbSql.AppendLine(" ) ItemDetail");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getDetailbyByGroupCode(InvoiceOBJ InvoiceOBJ,bool Comercial)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

            String Com = "";
            if (Comercial)
            {
                Com = "COMERCIAL";
            }
            else
            {
                Com = "NO COMERCIAL";
            }

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }

                sbSql.AppendLine(" SELECT ");
                sbSql.AppendLine("  CASE WHEN [ECL_GROUPCODE] IS NULL THEN 'GRAND TOTAL-" + Com + "' ELSE [ECL_GROUPCODE] END [GROUPCODE]");
                sbSql.AppendLine(" ,[ECL_LENTYPE]");
                sbSql.AppendLine(",CASE WHEN [ITEMID] IS NULL AND NOT([ECL_GROUPCODE] IS NULL) THEN 'TOTAL' ELSE [ITEMID] END [ITEMID]");
                sbSql.AppendLine(",ECL_APPCODE AS [APPCODE]");

                sbSql.AppendLine(" ,CASE ECL_SalesComercial ");
                sbSql.AppendLine(" 	    WHEN 1 THEN 'Com'");
                sbSql.AppendLine(" 	    WHEN 2 THEN 'NO Com'");
                sbSql.AppendLine("  END COM,CURRENCY");

           
          
            while (dtFrom <= dtTo)
            {

              
                if (Comercial) //come
                {
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Amount Cur]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Baht/SET]", dtFrom.Month));

                }
                else
                {

                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));

                }
                

                dtFrom = dtFrom.AddMonths(1);
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT ECL_GROUPCODE,ECL_LENTYPE,ITEMID,ECL_APPCODE [ECL_APPCODE]");
            sbSql.AppendLine(" ,CURRENCYCODEISO [CURRENCY]");
            sbSql.AppendLine(" ,EXCHRATE");

            sbSql.AppendLine(",[SET]");
            sbSql.AppendLine(",[PCS]");
           


            sbSql.AppendLine(" ,LINEAMOUNT [Amount Cur]");
            sbSql.AppendLine(" ,SalesPrice");
             
             sbSql.AppendLine(" ,LineAmountMST [Amount Baht]");
             sbSql.AppendLine(" ,LineAmountMST [Baht/PC]");
             sbSql.AppendLine(" ,InvoiceDate");
             sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
             sbSql.AppendLine(" ,ECL_SalesComercial");
          
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");

            if (InvoiceOBJ.ShipmentLocation == 6)
            {
            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT%' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT%' OR NUMBERSEQUENCEGROUP LIKE '%-INT%' OR NUMBERSEQUENCEGROUP LIKE '%-CINT%')"); //CEXT EXT INT CINT

            } else if (InvoiceOBJ.ShipmentLocation == 1)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CINT')"); //CEXT EXT INT CINT

            }
            else if (InvoiceOBJ.ShipmentLocation == 2)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT')"); //CEXT EXT INT CINT
            }
            else
            {

                sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP = ('" + InvoiceOBJ.NumberSequenceGroup.Replace("-", "-") + "')");

            }



             if (Comercial)
             {
                 sbSql.AppendLine(" AND ECL_SalesComercial=1");//com

             }
             else
             {
                 sbSql.AppendLine(" AND ECL_SalesComercial=2");//com

             }

           


            sbSql.AppendLine(" AND HOYA_TRADING = 0");
            sbSql.AppendLine(" ) ItemDetail");
            sbSql.AppendLine(" GROUP BY ECL_GROUPCODE,ECL_LENTYPE,[ITEMID],[ECL_APPCODE],CURRENCY,ECL_SalesComercial WITH ROLLUP");
            sbSql.AppendLine(" HAVING NOT(ECL_SalesComercial IS NULL) OR [ECL_LENTYPE] IS NULL");
            sbSql.AppendLine(" ORDER BY GROUPING(ECL_GROUPCODE),ECL_GROUPCODE ,GROUPING(ECL_LENTYPE),ECL_LENTYPE,GROUPING([ITEMID]),[ITEMID],");
            sbSql.AppendLine(" GROUPING([ECL_APPCODE]),[ECL_APPCODE],GROUPING([CURRENCY]),[CURRENCY]");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSummarybyLenType(InvoiceOBJ InvoiceOBJ, bool Comercial)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

            String Com = "";
            if (Comercial)
            {
                Com = "COMERCIAL";
            }
            else
            {
                Com = "NO COMERCIAL";
            }

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }

            sbSql.AppendLine(" SELECT ");

            if (Comercial)
            {
                sbSql.AppendLine("  CASE WHEN [ECL_LENTYPE] IS NULL THEN 'TOTAL GROSS SALE -" + Com + "' ELSE [ECL_LENTYPE] END [ECL_LENTYPE]");
                
            }
            else
            {
                sbSql.AppendLine("'TOTAL GROSS SALE -" + Com + "'");
             
            }



            while (dtFrom <= dtTo)
            {

                if (Comercial)
                {
                   
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                     sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
                }
                else
                {
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));

                }
                dtFrom = dtFrom.AddMonths(1);
            }

            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT ECL_LENTYPE");
            sbSql.AppendLine(" ,EXCHRATE");
            sbSql.AppendLine(" ,[SET]");
            sbSql.AppendLine(" ,LINEAMOUNT [Amount Cur]");
            sbSql.AppendLine(" ,SalesPrice");

            sbSql.AppendLine(" ,LineAmountMST [Amount Baht]");
            sbSql.AppendLine(" ,LineAmountMST [Baht/PC]");
            sbSql.AppendLine(" ,InvoiceDate");
            sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" ,ECL_SalesComercial");

            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");

            if (InvoiceOBJ.ShipmentLocation == 1)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CINT')"); //CEXT EXT INT CINT

            }
            else if (InvoiceOBJ.ShipmentLocation == 2)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT')"); //CEXT EXT INT CINT
            }
            else
            {

                sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP = ('" + InvoiceOBJ.NumberSequenceGroup.Replace("-", "-") + "')");

            }



            if (Comercial)
            {
                sbSql.AppendLine(" AND ECL_SalesComercial=1");//com
                sbSql.AppendLine(" AND HOYA_TRADING = 0");
                sbSql.AppendLine(" ) ItemDetail");
                sbSql.AppendLine(" GROUP BY ECL_LENTYPE WITH ROLLUP");
                sbSql.AppendLine(" ORDER BY GROUPING(ECL_LENTYPE),ECL_LENTYPE");


            }
            else
            {
                sbSql.AppendLine(" AND ECL_SalesComercial=2");//com
                sbSql.AppendLine(" ) ItemDetail");
            }




           

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getDetailbyLenType(InvoiceOBJ InvoiceOBJ, bool Comercial)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

            String Com = "";
            if (Comercial)
            {
                Com = "COMERCIAL";
            }
            else
            {
                Com = "NO COMERCIAL";
            }

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }

            sbSql.AppendLine(" SELECT ");
            sbSql.AppendLine("  CASE WHEN [ECL_LENTYPE] IS NULL THEN 'TOTAL GROSS SALE -" + Com + "' ELSE [ECL_LENTYPE] END [ECL_LENTYPE]");
            sbSql.AppendLine(" ,CASE WHEN [ECL_GROUPCODE]  IS NULL AND NOT([ECL_LENTYPE] IS NULL) THEN 'TOTAL' ELSE [ECL_GROUPCODE] END [GROUPCODE]");
            sbSql.AppendLine(" ,ITEMID AS [ITEMID],ECL_APPCODE AS [APPCODE]");

             sbSql.AppendLine(" ,CASE ECL_SalesComercial ");
             sbSql.AppendLine(" 	    WHEN 1 THEN 'Com'");
             sbSql.AppendLine(" 	    WHEN 2 THEN 'NO Com'");
             sbSql.AppendLine("  END COM,CURRENCY");

            while (dtFrom <= dtTo)
            {
                if(Comercial){
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Amount Cur]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Baht/SET]", dtFrom.Month));
        
                }else{
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/PCS]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/SET]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,''", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
                sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Baht/SET]", dtFrom.Month));
                }

                dtFrom = dtFrom.AddMonths(1);
            }
            sbSql.AppendLine("FROM(");
            sbSql.AppendLine(" SELECT ECL_GROUPCODE,ECL_LENTYPE,ITEMID,ECL_APPCODE [ECL_APPCODE]");
            sbSql.AppendLine(" ,CURRENCYCODEISO [CURRENCY]");
            sbSql.AppendLine(" ,EXCHRATE");
            sbSql.AppendLine(" ,[SET]");
            sbSql.AppendLine(" ,[PCS]");
            sbSql.AppendLine(" ,LINEAMOUNT [Amount Cur]");
            sbSql.AppendLine(" ,SalesPrice");;
             
             sbSql.AppendLine(" ,LineAmountMST [Amount Baht]");
             sbSql.AppendLine(" ,LineAmountMST [Baht/PC]");
             sbSql.AppendLine(" ,InvoiceDate");
             sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
             sbSql.AppendLine(" ,ECL_SalesComercial");
             sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
         
            sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");


            if (InvoiceOBJ.ShipmentLocation == 1)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CINT')"); //CEXT EXT INT CINT

            }
            else if (InvoiceOBJ.ShipmentLocation == 2)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT')"); //CEXT EXT INT CINT
            }
            else
            {

                sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP = ('" + InvoiceOBJ.NumberSequenceGroup.Replace("-", "-") + "')");

            }



            if (Comercial)
            {
                sbSql.AppendLine(" AND ECL_SalesComercial=1");//com
            }
            else
            {
                sbSql.AppendLine(" AND ECL_SalesComercial=2");//com
            }


        sbSql.AppendLine(" AND HOYA_TRADING = 0");
        sbSql.AppendLine(" ) ItemDetail");
        sbSql.AppendLine(" GROUP BY ECL_LENTYPE,ECL_GROUPCODE,[ITEMID],[ECL_APPCODE],CURRENCY,ECL_SalesComercial WITH ROLLUP");
        sbSql.AppendLine(" HAVING NOT(ECL_SalesComercial IS NULL) OR ECL_GROUPCODE IS NULL");
        sbSql.AppendLine(" ORDER BY GROUPING(ECL_LENTYPE),ECL_LENTYPE, GROUPING(ECL_GROUPCODE),ECL_GROUPCODE,GROUPING([ITEMID]),[ITEMID],");
        sbSql.AppendLine(" GROUPING([ECL_APPCODE]),[ECL_APPCODE],GROUPING([CURRENCY]),[CURRENCY]");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSummarybyLenTypeReturnAndTrading(InvoiceOBJ InvoiceOBJ,bool type,bool TradingReturn,string str)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);
            bool check = true;
            String strFac = InvoiceOBJ.strFactory;

            

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }

            sbSql.AppendLine(" SELECT ");

              while (dtFrom <= dtTo)
            {
                if (check)
                {


                    if (str == "DetailbyLentype")
                    {
                        sbSql.AppendLine(String.Format(" SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Amount Cur]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/PCS]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/SET]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Baht/SET]", dtFrom.Month));


                    }
                    else
                    {
                        sbSql.AppendLine(String.Format(" SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Baht/PC] ELSE 0 END) [Baht/PC]", dtFrom.Month));

                    }

                    check = false;
                }
                else
                {

                    if (str == "DetailbyLentype")
                    {
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Amount Cur]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/PCS]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [SalesPrice/SET]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Baht/PCS]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Cur] ELSE 0 END) [Baht/SET]", dtFrom.Month));
                    }
                    else
                    {
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Amount Baht] ELSE 0 END) [Amount Baht]", dtFrom.Month));
                        sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [Baht/PC] ELSE 0 END) [Baht/PC]", dtFrom.Month));


                    }
                }
          
                dtFrom = dtFrom.AddMonths(1);
            }
                
            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine(" SELECT ECL_LENTYPE");
            sbSql.AppendLine(" ,EXCHRATE");
            sbSql.AppendLine(" ,[PCS]");
            sbSql.AppendLine(" ,LineAmount [Amount Cur]");
            sbSql.AppendLine(" ,SalesPrice");
            sbSql.AppendLine(" ,LineAmountMst [Amount Baht]");
            sbSql.AppendLine(" ,SalesPrice*EXCHRATE [Baht/PC]");
            sbSql.AppendLine(" ,InvoiceDate");
            sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" ,[SET],ECL_SalesComercial");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
          

            sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");


            if (type)
            {
                if (TradingReturn)
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD" + "'))");
                }
                else
                {
                    sbSql.AppendLine(" AND HOYA_TRADING = 1");
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP != ('" + strFac + "-RTRD" + "'))");
                   
                }

            }
            else
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT" + "'))");
                sbSql.AppendLine(" AND HOYA_TRADING = 0");
            }

            sbSql.AppendLine(" AND ECL_SalesComercial=1");//com
            sbSql.AppendLine(" ) ItemDetail");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSaleByCustomerAndCurrency(InvoiceOBJ InvoiceOBJ,bool TypeSale ,bool Trading)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

        
            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }

            if (TypeSale)
            {

                sbSql.AppendLine(" SELECT 'SALE' [TOTAL SALE]");
                sbSql.AppendLine(" ,CASE WHEN CURRENCYCODEISO ='CNY' THEN 'RMB' ELSE CURRENCYCODEISO END [CUR]");

                while (dtFrom <= dtTo)
                {
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='THB' THEN [LineAmount] ELSE 0 END)[THB]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='JPY' THEN [LineAmount] ELSE 0 END)[JPY]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='USD' THEN [LineAmount] ELSE 0 END)[USD]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='CNY' THEN [LineAmount] ELSE 0 END) [RMB]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [LineAmountMST] ELSE 0 END)[THB]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",null [/SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",null [/PCS]", dtFrom.Month));
                    dtFrom = dtFrom.AddMonths(1);
                }

                sbSql.AppendLine(" FROM(");
                sbSql.AppendLine("SELECT");
                sbSql.AppendLine(" CURRENCYCODEISO");
                //sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM([SET]) ELSE 0 END [SET]");
                //sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(PCS) ELSE 0 END [PCS]");
                sbSql.AppendLine(",  SUM([SET])  [SET]");
                sbSql.AppendLine(", SUM(PCS) [PCS]");

                sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
                sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
                sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
                sbSql.AppendLine("FROM HOYA_vwSalesDetail");
                sbSql.AppendLine(" WHERE InventSiteId='" + InvoiceOBJ.Factory + "'");

                if (Trading)
                {
                    sbSql.AppendLine(" AND HOYA_TRADING = 1");
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP != ('" + strFac + "-RTRD" + "'))");

                }
                else
                {
                   // sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT" + "') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-INT" + "'))  ");
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT' OR NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CINT')"); //CEXT EXT INT CINT

                    sbSql.AppendLine("AND (HOYA_TRADING = 0)");
                }

                sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
                sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
                sbSql.AppendLine("GROUP BY CURRENCYCODEISO,ECL_SALESCOMERCIAL,INVOICEDATE)salesTotal ");
                sbSql.AppendLine(" GROUP BY CURRENCYCODEISO ");


            }
            else
            {
                //Return
                // dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
                // dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

                sbSql.AppendLine("SELECT 'SALE RETURN'  [TOTAL SALE RETURN]");
                sbSql.AppendLine(" ,CASE WHEN CURRENCYCODEISO ='CNY' THEN 'RMB' ELSE CURRENCYCODEISO END [CUR]");

                while (dtFrom <= dtTo)
                {
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(" ,SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='THB' THEN [LineAmount] ELSE 0 END)[THB]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='JPY' THEN [LineAmount] ELSE 0 END)[JPY]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='USD' THEN [LineAmount] ELSE 0 END)[USD]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='CNY' THEN [LineAmount] ELSE 0 END) [RMB]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [LineAmountMST] ELSE 0 END)[THB]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",null [/SET]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",null [/PCS]", dtFrom.Month));

                    dtFrom = dtFrom.AddMonths(1);
                }

                sbSql.AppendLine(" FROM(");
                sbSql.AppendLine(" SELECT");
                sbSql.AppendLine(" CURRENCYCODEISO");
               // sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM([SET]) ELSE 0 END [SET]");
              //  sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(PCS) ELSE 0 END [PCS]");

                sbSql.AppendLine(",  SUM([SET])  [SET]");
                sbSql.AppendLine(", SUM(PCS) [PCS]");
                


                sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
                sbSql.AppendLine(" ,CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
                sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
                 sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
                 sbSql.AppendLine(" WHERE InventSiteId='" + InvoiceOBJ.Factory + "'");

                 if (Trading)
                 {
                    // sbSql.AppendLine("  AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD" + "')) ");
                    // sbSql.AppendLine(" AND HOYA_TRADING = 1");
                     sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD" + "'))");
                 }
                 else
                 { 
                       sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT" + "') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT" + "'))  ");
                       sbSql.AppendLine("AND (HOYA_TRADING = 0)");

                 }

                 sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
                 sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
                 sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
                 sbSql.AppendLine("GROUP BY CURRENCYCODEISO,ECL_SALESCOMERCIAL,INVOICEDATE)salesTotal ");
                 sbSql.AppendLine(" GROUP BY CURRENCYCODEISO ");

            }

           

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// End getSaleByCustomerAndCurrency

        public ADODB.Recordset getSaleByCustomerAndCustCode(DataTable dt,InvoiceOBJ InvoiceOBJ, string InvoiceAcc, bool Numbersequence)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;


            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }


            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN  sales.CUR IS NULL THEN CASE WHEN sales.[TOTAL SALE] = 'EXT' THEN 'TOTAL SALE' ELSE");
            sbSql.AppendLine(" CASE WHEN sales.[TOTAL SALE] = 'REXT' THEN 'TOTAL SALE RETURN' END END ELSE ");
            sbSql.AppendLine(" CASE WHEN sales.[TOTAL SALE] = 'EXT' THEN 'SALE' ");
            sbSql.AppendLine("WHEN sales.[TOTAL SALE] = 'REXT' THEN 'SALE RETURN'");
            sbSql.AppendLine(" END  END [TOTALSALES]");
            sbSql.AppendLine(" ,sales.CUR");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM(sales.[SET" + String.Format("{0:yyMM}", dr[0]) + "]) [SET" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[PCS" + String.Format("{0:yyMM}", dr[0]) + "]) [PCS" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[THB" + String.Format("{0:yyMM}", dr[0]) + "]) [THB" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[JPY" + String.Format("{0:yyMM}", dr[0]) + "]) [JPY" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[USD" + String.Format("{0:yyMM}", dr[0]) + "]) [USD" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[RMB" + String.Format("{0:yyMM}", dr[0]) + "]) [RMB" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[THBMST" + String.Format("{0:yyMM}", dr[0]) + "]) [THBMST" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,'' [/SET" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,'' [/PCS" + String.Format("{0:yyMM}", dr[0]) + "] ");
              

            }

            sbSql.AppendLine("FROM(");

         sbSql.AppendLine("SELECT");
         sbSql.AppendLine("CASE WHEN  NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT'");
         sbSql.AppendLine("OR NUMBERSEQUENCEGROUP LIKE '%-CINT' OR NUMBERSEQUENCEGROUP LIKE '%-TRD' THEN 'EXT'  ELSE ");
         sbSql.AppendLine("CASE WHEN  NUMBERSEQUENCEGROUP LIKE '%-REXT' OR NUMBERSEQUENCEGROUP LIKE '%-RINT' OR NUMBERSEQUENCEGROUP LIKE '%-RTRD' THEN 'REXT'");
         sbSql.AppendLine("END END [TOTAL SALE]");
         sbSql.AppendLine(",CASE WHEN CURRENCYCODEISO ='CNY' THEN 'RMB' ELSE CURRENCYCODEISO END [CUR]");


              
                foreach (DataRow  dr in dt.Rows)
                {
                    sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) +" THEN [SET] ELSE 0 END) [SET" + String.Format("{0:yyMM}", dr[0]) + "]");
                    sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " THEN [PCS] ELSE 0 END) [PCS" + String.Format("{0:yyMM}", dr[0]) + "]");
                    sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " AND CURRENCYCODEISO ='THB' THEN [LineAmount] ELSE 0 END) [THB" + String.Format("{0:yyMM}", dr[0]) + "]");
                    sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " AND CURRENCYCODEISO ='JPY' THEN [LineAmount]  ELSE 0 END) [JPY" + String.Format("{0:yyMM}", dr[0]) + "]");
                    sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " AND CURRENCYCODEISO ='USD' THEN [LineAmount]  ELSE 0 END) [USD" + String.Format("{0:yyMM}", dr[0]) + "]");
                    sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " AND CURRENCYCODEISO ='CNY' THEN [LineAmount]  ELSE 0 END) [RMB" + String.Format("{0:yyMM}", dr[0]) + "]");
                    sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " THEN [LineAmountMST] ELSE 0 END) [THBMST" + String.Format("{0:yyMM}", dr[0]) + "]");

                    //sbSql.AppendLine(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [LineAmountMST] ELSE 0 END)[THBMST" + String.Format("{0:yyMM}", dr[0]) + "]");
                   // sbSql.AppendLine(String.Format(",null [/SET" + String.Format("{0:yyMM}", dr[0]) + "]", dtFrom.Month));
                   // sbSql.AppendLine(String.Format(",null [/PCS" + String.Format("{0:yyMM}", dr[0]) + "]", dtFrom.Month));
                    //dtFrom = dtFrom.AddMonths(1);
                }

        sbSql.AppendLine(" FROM(");
        sbSql.AppendLine(" SELECT");
        sbSql.AppendLine("  NUMBERSEQUENCEGROUP");
        sbSql.AppendLine(",NAMEALIAS");
        sbSql.AppendLine(",INVOICEACCOUNT");
        sbSql.AppendLine(",CURRENCYCODEISO");
       // sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM([SET]) ELSE 0 END [SET]");
        //sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(PCS) ELSE 0 END [PCS]");
        sbSql.AppendLine(",SUM([SET]) [SET]");
        sbSql.AppendLine(",SUM([PCS]) [PCS]");

        sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
        sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
       // sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
        sbSql.AppendLine(",INVOICEDATE [INVOICEDATE]");
        sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
        sbSql.AppendLine(" WHERE InventSiteId='" + InvoiceOBJ.Factory + "'");

//=====================================================================================//

        if (Numbersequence)
        {
            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT" + "')  OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT" + "'))");
            sbSql.AppendLine("AND (HOYA_TRADING = 0)");

        }
        else
        {

            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-INT" + "') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT" + "'))");
            sbSql.AppendLine("AND (HOYA_TRADING = 0)");

        }

 //=====================================================================================//



                sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
                sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
                  if (InvoiceAcc != "")
            {
                sbSql.AppendLine("AND INVOICEACCOUNT ='" + InvoiceAcc + "'");
            }
              
                sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,INVOICEACCOUNT,CURRENCYCODEISO,ECL_SALESCOMERCIAL,INVOICEDATE");


        sbSql.AppendLine("UNION ALL");
        sbSql.AppendLine(" SELECT");
        sbSql.AppendLine("  NUMBERSEQUENCEGROUP");
        sbSql.AppendLine(",NAMEALIAS");
        sbSql.AppendLine(",INVOICEACCOUNT");
        sbSql.AppendLine(",CURRENCYCODEISO");

        //sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM([SET]) ELSE 0 END [SET]");
        //sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(PCS) ELSE 0 END [PCS]");
        sbSql.AppendLine(",SUM([SET]) [SET]");
        sbSql.AppendLine(",SUM([PCS]) [PCS]");

        sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
        sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
        //sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
        sbSql.AppendLine(",INVOICEDATE [INVOICEDATE]");
        sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
        sbSql.AppendLine(" WHERE InventSiteId='" + InvoiceOBJ.Factory + "'");



     

        //==================================================================================//
        if (Numbersequence)
        {
            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT" + "'))");
            sbSql.AppendLine("AND (HOYA_TRADING = 0)");

        }
        else
        {

            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT" + "'))");
            sbSql.AppendLine("AND (HOYA_TRADING = 0)");

        }

        //==================================================================================//


                sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
                sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
                sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
                
           
            if (InvoiceAcc != "")
            {
                sbSql.AppendLine("AND INVOICEACCOUNT ='" + InvoiceAcc + "'");
            }
                sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,INVOICEACCOUNT,CURRENCYCODEISO,ECL_SALESCOMERCIAL,INVOICEDATE)salesTotal");

                sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,INVOICEACCOUNT,CURRENCYCODEISO --WITH ROLLUP ");
                //sbSql.AppendLine("HAVING NOT NUMBERSEQUENCEGROUP IS NULL AND  NOT(INVOICEACCOUNT IS NULL)");
                sbSql.AppendLine(")as sales");
                //sbSql.AppendLine("ORDER BY GROUPING(INVOICEACCOUNT),INVOICEACCOUNT,GROUPING(NUMBERSEQUENCEGROUP),NUMBERSEQUENCEGROUP,");
                //sbSql.AppendLine("GROUPING(NAMEALIAS),NAMEALIAS,GROUPING(CURRENCYCODEISO),CURRENCYCODEISO");
                sbSql.AppendLine("GROUP BY [TOTAL SALE],CUR WITH ROLLUP HAVING not [TOTAL SALE] is null");
                sbSql.AppendLine("ORDER BY GROUPING([TOTAL SALE]),[TOTAL SALE],GROUPING(CUR),CUR");

            
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// End getSaleByCustomerAndCustcode


        public ADODB.Recordset getSaleByCustomerAndCustCode2(InvoiceOBJ InvoiceOBJ, string InvoiceAcc, bool Numbersequence)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;


            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }


            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN CURRENCYCODEISO IS NULL THEN 'TOTAL'+ ");
            sbSql.AppendLine(" CASE WHEN  NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT' OR NUMBERSEQUENCEGROUP LIKE '%-CINT' OR NUMBERSEQUENCEGROUP LIKE '%-TRD' THEN 'SALE' ELSE");
            sbSql.AppendLine("'SALE RETURN' END ELSE ");
            sbSql.AppendLine("CASE WHEN  NUMBERSEQUENCEGROUP LIKE '%-REXT' OR NUMBERSEQUENCEGROUP LIKE '%-RINT' OR NUMBERSEQUENCEGROUP LIKE '%-RTRD' THEN 'SALE RETURN' ELSE 'SALE'");
            sbSql.AppendLine("END END [TOTAL SALE]");
            sbSql.AppendLine(" ,CASE WHEN CURRENCYCODEISO ='CNY' THEN 'RMB' ELSE CURRENCYCODEISO END [CUR]");


            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [SET] ELSE 0 END) [SET]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [PCS] ELSE 0 END) [PCS]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='THB' THEN [LineAmount] ELSE 0 END)[THB]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='JPY' THEN [LineAmount] ELSE 0 END)[JPY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='USD' THEN [LineAmount] ELSE 0 END)[USD]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} AND CURRENCYCODEISO ='CNY' THEN [LineAmount] ELSE 0 END) [RMB]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [LineAmountMST] ELSE 0 END)[THB]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",null [/SET]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",null [/PCS]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }

            sbSql.AppendLine(" FROM(");
            sbSql.AppendLine(" SELECT");
            sbSql.AppendLine("  NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",CURRENCYCODEISO");
 
            sbSql.AppendLine(",SUM([SET]) [SET]");
            sbSql.AppendLine(",SUM([PCS]) [PCS]");

            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
            sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + InvoiceOBJ.Factory + "'");

            //=====================================================================================//

            if (Numbersequence)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT" + "')  OR NUMBERSEQUENCEGROUP = ('" + strFac + "-CEXT" + "'))");
                sbSql.AppendLine("AND (HOYA_TRADING = 0)");

            }
            else
            {

                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-INT" + "') OR  NUMBERSEQUENCEGROUP = ('" + strFac + "-CINT" + "'))");
                sbSql.AppendLine("AND (HOYA_TRADING = 0)");

            }

            //=====================================================================================//



            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
            if (InvoiceAcc != "")
            {
                sbSql.AppendLine("AND INVOICEACCOUNT IN ('" + InvoiceAcc.Replace(",", "','") + "')");
            }

            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,CURRENCYCODEISO,INVOICEDATE,ECL_SALESCOMERCIAL");


            sbSql.AppendLine("UNION ALL");
            sbSql.AppendLine(" SELECT");
            sbSql.AppendLine("  NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",CURRENCYCODEISO");

            sbSql.AppendLine(",SUM([SET]) [SET]");
            sbSql.AppendLine(",SUM([PCS]) [PCS]");

            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
            sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),INVOICEDATE,12)+'01',12) InvoiceMonth");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + InvoiceOBJ.Factory + "'");





            //==================================================================================//
            if (Numbersequence)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-REXT" + "'))");
                sbSql.AppendLine("AND (HOYA_TRADING = 0)");

            }
            else
            {

                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-RINT" + "'))");
                sbSql.AppendLine("AND (HOYA_TRADING = 0)");

            }

            //==================================================================================//


            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");


            if (InvoiceAcc != "")
            {
                sbSql.AppendLine("AND INVOICEACCOUNT IN ('" + InvoiceAcc.Replace(",", "','") + "')");
            }
            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,CURRENCYCODEISO,INVOICEDATE,ECL_SALESCOMERCIAL)salesTotal");

            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,CURRENCYCODEISO WITH ROLLUP  ");
            sbSql.AppendLine("HAVING NOT NUMBERSEQUENCEGROUP IS NULL");
            //sbSql.AppendLine("ORDER BY GROUPING(INVOICEACCOUNT),INVOICEACCOUNT,GROUPING(NUMBERSEQUENCEGROUP),NUMBERSEQUENCEGROUP,");
            //sbSql.AppendLine("GROUPING(NAMEALIAS),NAMEALIAS,GROUPING(CURRENCYCODEISO),CURRENCYCODEISO");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// End 



        public ADODB.Recordset getCustomer(InvoiceOBJ InvoiceOBJ, bool Trading, bool Numbersequence)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;


            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }


            sbSql.AppendLine("SELeCT * FROM(");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" CASE WHEN NUMBERSEQUENCEGROUP='" + strFac + "-EXT" + "' OR NUMBERSEQUENCEGROUP='" + strFac + "-REXT" + "' THEN 'EXTERNAL SALE' + CASE WHEN NOT(NAMEALIAS IS NULL) THEN ' ('+NAMEALIAS+')' ELSE '' END ELSE ");
            sbSql.AppendLine(" 	CASE WHEN NUMBERSEQUENCEGROUP='" + strFac + "-INT" + "' OR NUMBERSEQUENCEGROUP='" + strFac + "-RINT" + "' THEN 'INTERNAL SALE' + CASE WHEN NOT(NAMEALIAS IS NULL) THEN ' ('+NAMEALIAS+')' ELSE '' END ELSE ");
            sbSql.AppendLine("		CASE WHEN NUMBERSEQUENCEGROUP='" + strFac + "-TRD" + "' OR NUMBERSEQUENCEGROUP='" + strFac + "-RTRD" + "' OR NUMBERSEQUENCEGROUP='" + strFac + "-CTRD" +  "'THEN 'EXTERNAL SALE' + CASE WHEN NOT(NAMEALIAS IS NULL) THEN ' ('+NAMEALIAS+')' ELSE '' END ELSE ");
            sbSql.AppendLine("NUMBERSEQUENCEGROUP END  END END NUMBERSEQUENCEGROUP2");
            sbSql.AppendLine(",INVOICEACCOUNT");
          
            sbSql.AppendLine(" FROM(");
            sbSql.AppendLine(" SELECT");
            sbSql.AppendLine("  NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",NAMEALIAS");
            sbSql.AppendLine(",INVOICEACCOUNT");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + InvoiceOBJ.Factory + "'");

            if (Numbersequence)
            {
                if (Trading)
                {
                    sbSql.AppendLine("AND (HOYA_TRADING = 1 OR NUMBERSEQUENCEGROUP = 'MO-CTRD')");
                   // sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT" + "'))  ");
                }
                else
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-EXT" + "'))  ");
                    sbSql.AppendLine("AND (HOYA_TRADING = 0)");

                }

            }
            else
            {

                if (Trading)
                {
                    //sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD" + "'))");
                    sbSql.AppendLine("AND (HOYA_TRADING = 1)");
                    //sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-INT" + "'))  ");
                }
                else
                {
                    sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-INT" + "'))  ");
                    sbSql.AppendLine("AND (HOYA_TRADING = 0)");

                }

            }

          

            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
            sbSql.AppendLine("AND INVOICEACCOUNT !='ARAF001'");
            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,INVOICEACCOUNT)salesTotal");

            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,INVOICEACCOUNT)as sales ");
            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP2,INVOICEACCOUNT");
            sbSql.AppendLine("ORDER BY INVOICEACCOUNT ");
           


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// End getSaleByCustomerAndCustcode

        public ADODB.Recordset getSaleByCustomerAndCustCodeTrading(DataTable dt,InvoiceOBJ InvoiceOBJ, bool Trading,string InvoiceAcc)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;


            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }




            sbSql.AppendLine("SELECT");
           // sbSql.AppendLine("CASE WHEN  sales.CUR IS NULL THEN CASE WHEN sales.[TOTAL SALE] = 'TRD' THEN 'TOTAL SALE' ELSE");
           // sbSql.AppendLine(" CASE WHEN sales.[TOTAL SALE] = 'RTRD' THEN 'TOTAL SALE RETURN' END END ELSE ");
          //  sbSql.AppendLine(" CASE WHEN sales.[TOTAL SALE] = 'TRD' THEN 'SALE' ");
           // sbSql.AppendLine("WHEN sales.[TOTAL SALE] = 'RTRD' THEN 'SALE RETURN'");


            //10/5/2018

             sbSql.AppendLine("CASE WHEN  sales.CUR IS NULL THEN CASE WHEN sales.[TOTAL SALE] = 'EXT' THEN 'TOTAL SALE' ELSE");
             sbSql.AppendLine(" CASE WHEN sales.[TOTAL SALE] = 'REXT' THEN 'TOTAL SALE RETURN' END END ELSE ");
              sbSql.AppendLine(" CASE WHEN sales.[TOTAL SALE] = 'EXT' THEN 'SALE' ");
             sbSql.AppendLine("WHEN sales.[TOTAL SALE] = 'REXT' THEN 'SALE RETURN'");



            sbSql.AppendLine(" END  END [TOTALSALES]");
            sbSql.AppendLine(" ,sales.CUR");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM(sales.[SET" + String.Format("{0:yyMM}", dr[0]) + "]) [SET" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[PCS" + String.Format("{0:yyMM}", dr[0]) + "]) [PCS" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[THB" + String.Format("{0:yyMM}", dr[0]) + "]) [THB" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[JPY" + String.Format("{0:yyMM}", dr[0]) + "]) [JPY" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[USD" + String.Format("{0:yyMM}", dr[0]) + "]) [USD" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[RMB" + String.Format("{0:yyMM}", dr[0]) + "]) [RMB" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM(sales.[THBMST" + String.Format("{0:yyMM}", dr[0]) + "]) [THBMST" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,'' [/SET" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,'' [/PCS" + String.Format("{0:yyMM}", dr[0]) + "] ");


            }

            sbSql.AppendLine("FROM(");

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN  NUMBERSEQUENCEGROUP LIKE '%-TRD' OR NUMBERSEQUENCEGROUP LIKE '%-CTRD' THEN 'EXT'  ELSE ");
            sbSql.AppendLine("CASE WHEN  NUMBERSEQUENCEGROUP LIKE '%-RTRD' THEN 'REXT' ");
            sbSql.AppendLine("END END [TOTAL SALE]");

            sbSql.AppendLine(",CASE WHEN CURRENCYCODEISO ='CNY' THEN 'RMB' ELSE CURRENCYCODEISO END [CUR]");


            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " THEN [SET] ELSE 0 END) [SET" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " THEN [PCS] ELSE 0 END) [PCS" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " AND CURRENCYCODEISO ='THB' THEN [LineAmount] ELSE 0 END) [THB" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " AND CURRENCYCODEISO ='JPY' THEN [LineAmount]  ELSE 0 END) [JPY" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " AND CURRENCYCODEISO ='USD' THEN [LineAmount]  ELSE 0 END) [USD" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " AND CURRENCYCODEISO ='CNY' THEN [LineAmount]  ELSE 0 END) [RMB" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(",SUM(CASE WHEN CONVERT(CHAR(4),INVOICEDATE,12)=" + String.Format("{0:yyMM}", dr[0]) + " THEN [LineAmountMST] ELSE 0 END) [THBMST" + String.Format("{0:yyMM}", dr[0]) + "]");

                //sbSql.AppendLine(",SUM(CASE WHEN MONTH(InvoiceMonth)={0} THEN [LineAmountMST] ELSE 0 END)[THBMST" + String.Format("{0:yyMM}", dr[0]) + "]");
                // sbSql.AppendLine(String.Format(",null [/SET" + String.Format("{0:yyMM}", dr[0]) + "]", dtFrom.Month));
                // sbSql.AppendLine(String.Format(",null [/PCS" + String.Format("{0:yyMM}", dr[0]) + "]", dtFrom.Month));
                //dtFrom = dtFrom.AddMonths(1);
            }

            sbSql.AppendLine(" FROM(");


            sbSql.AppendLine(" SELECT");
            sbSql.AppendLine("  NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",NAMEALIAS");
            sbSql.AppendLine(",INVOICEACCOUNT");
            sbSql.AppendLine(",CURRENCYCODEISO");
            // sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM([SET]) ELSE 0 END [SET]");
            //sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(PCS) ELSE 0 END [PCS]");
            sbSql.AppendLine(",SUM([SET]) [SET]");
            sbSql.AppendLine(",SUM([PCS]) [PCS]");

            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
            sbSql.AppendLine(",INVOICEDATE [INVOICEDATE]");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + InvoiceOBJ.Factory + "'");

            //==================================================================================//
            if (Trading)
            {
                sbSql.AppendLine("AND (HOYA_TRADING = 1)");
            }

            //==================================================================================//


            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
            if (InvoiceAcc != "")
            {
                sbSql.AppendLine("AND INVOICEACCOUNT ='" + InvoiceAcc + "'");
            }

            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,INVOICEACCOUNT,CURRENCYCODEISO,ECL_SALESCOMERCIAL,INVOICEDATE");


            sbSql.AppendLine("UNION ALL");
            sbSql.AppendLine(" SELECT");
            sbSql.AppendLine("  NUMBERSEQUENCEGROUP");
            sbSql.AppendLine(",NAMEALIAS");
            sbSql.AppendLine(",INVOICEACCOUNT");
            sbSql.AppendLine(",CURRENCYCODEISO");

            //sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM([SET]) ELSE 0 END [SET]");
            //sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(PCS) ELSE 0 END [PCS]");
            sbSql.AppendLine(",SUM([SET]) [SET]");
            sbSql.AppendLine(",SUM([PCS]) [PCS]");

            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmount) ELSE 0 END [LineAmount]");
            sbSql.AppendLine(",CASE WHEN ECL_SALESCOMERCIAL=1 THEN SUM(LineAmountMST) ELSE 0 END [LineAmountMST]");
            sbSql.AppendLine(",INVOICEDATE [INVOICEDATE]");
            sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
            sbSql.AppendLine(" WHERE InventSiteId='" + InvoiceOBJ.Factory + "'");



     

            //==================================================================================//
            if (Trading)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP = ('" + strFac + "-RTRD" + "'))  ");
            }

            //==================================================================================//


            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");


            if (InvoiceAcc != "")
            {
                sbSql.AppendLine("AND INVOICEACCOUNT ='" + InvoiceAcc + "'");
            }
            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,INVOICEACCOUNT,CURRENCYCODEISO,ECL_SALESCOMERCIAL,INVOICEDATE)salesTotal");

            sbSql.AppendLine("GROUP BY NUMBERSEQUENCEGROUP,NAMEALIAS,INVOICEACCOUNT,CURRENCYCODEISO --WITH ROLLUP ");
            //sbSql.AppendLine("HAVING NOT NUMBERSEQUENCEGROUP IS NULL AND  NOT(INVOICEACCOUNT IS NULL)");
           // sbSql.AppendLine("ORDER BY GROUPING(INVOICEACCOUNT),INVOICEACCOUNT,GROUPING(NUMBERSEQUENCEGROUP),NUMBERSEQUENCEGROUP,");
            //sbSql.AppendLine("GROUPING(NAMEALIAS),NAMEALIAS,GROUPING(CURRENCYCODEISO),CURRENCYCODEISO");
            sbSql.AppendLine(")as sales");
            sbSql.AppendLine("GROUP BY [TOTAL SALE],CUR WITH ROLLUP HAVING not [TOTAL SALE] is null");
            sbSql.AppendLine("ORDER BY GROUPING([TOTAL SALE]),[TOTAL SALE],GROUPING(CUR),CUR");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// End getSaleByCustomerAndCustcode

        public ADODB.Recordset getInvoiceDetail(InvoiceOBJ InvoiceOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }


                sbSql.AppendLine(" SELECT ");
                sbSql.AppendLine("  INVOICEDATE [DATE]");
                sbSql.AppendLine("  ,INVOICEID");
                sbSql.AppendLine(" ,ItemID");
                sbSql.AppendLine(",EXTERNALITEMID [Part No]");
                sbSql.AppendLine(" ,ECL_REASON[TradPart]");
                sbSql.AppendLine(" ,CURRENCY Curr");
                sbSql.AppendLine(",ECL_GLASSTYPE");
                sbSql.AppendLine(",CASE ECL_SalesComercial WHEN 1 THEN 'Com'WHEN 2 THEN 'NO Com'END COM");
            

             sbSql.AppendLine(",SUM([SET]) [SET]");
             sbSql.AppendLine(",SUM([PCS]) [PCS]");
             sbSql.AppendLine(",SUM(NW) [KGS]");
             sbSql.AppendLine(",SUM([Amount Cur]) [Amount Cur]");
             sbSql.AppendLine(",SUM([Amount Baht]) [Amount Baht]");
             //sbSql.AppendLine(",SUM([Baht/PC]) [Baht/PC]");

             sbSql.AppendLine(" FROM (");
             sbSql.AppendLine(" SELECT ECL_REASON,INVOICEID,ItemID");
             sbSql.AppendLine(",EXTERNALITEMID");
             sbSql.AppendLine("  ,CURRENCYCODEISO [CURRENCY]");
             sbSql.AppendLine(",EXCHRATE");
             sbSql.AppendLine(",[PCS]");
             sbSql.AppendLine(" ,LINEAMOUNT [Amount Cur]");
             sbSql.AppendLine(" ,LineAmountMST [Amount Baht]");
             //sbSql.AppendLine(",SalesPrice*EXCHRATE [Baht/PC]");
             sbSql.AppendLine(",InvoiceDate");
             sbSql.AppendLine(",ECL_GLASSTYPE");
             sbSql.AppendLine(",[SET],ECL_SalesComercial,NW");
    
             sbSql.AppendLine(" FROM HOYA_vwSalesDetail");
             sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
             sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
             sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
             sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");

             if (InvoiceOBJ.ShipmentLocation == 6)
             {
                 sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT%' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT%' OR NUMBERSEQUENCEGROUP LIKE '%-INT%' OR NUMBERSEQUENCEGROUP LIKE '%-CINT%')"); //CEXT EXT INT CINT

             }
             else if (InvoiceOBJ.ShipmentLocation == 1)
             {
                 sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CINT')"); //CEXT EXT INT CINT

             }
             else if (InvoiceOBJ.ShipmentLocation == 2)
             {
                 sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT')"); //CEXT EXT INT CINT
             }
             else
             {

                 sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP = ('" + InvoiceOBJ.NumberSequenceGroup.Replace("-", "-") + "')");

             }

            // sbSql.AppendLine("AND INVOICEAMOUNTMST >0");
             sbSql.AppendLine(" ) ItemDetail");
             sbSql.AppendLine(" GROUP BY INVOICEDATE,INVOICEID,ItemID,EXTERNALITEMID,ECL_REASON,CURRENCY,ECL_GLASSTYPE,ECL_SalesComercial ");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getInvoiceByDate(InvoiceOBJ InvoiceOBJ,DateTime dtInvoiceDate,string locationID)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;

            if (InvoiceOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = InvoiceOBJ.strFactory;
            }


         sbSql.AppendLine(" SELECT LEDGERVOUCHER,INVOICEDATE,INVOICEID,CustName,SALESNAME,CURRENCYCODEISO,EXCHRATE");
        sbSql.AppendLine(",CASE WHEN ECL_SalesComercial=1 THEN SUM([SET]) ELSE 0 END [SET COM]");
        sbSql.AppendLine(",CASE WHEN ECL_SalesComercial=1 THEN SUM(PCS) ELSE 0 END [PCS COM]");
        sbSql.AppendLine(",CASE WHEN ECL_SalesComercial=1 THEN SUM(NW) ELSE 0 END [KG COM]");
        sbSql.AppendLine(",CASE WHEN ECL_SalesComercial=1 THEN SUM(LineAmount) ELSE 0 END [Inv.Amt. COM]");
        sbSql.AppendLine(",CASE WHEN ECL_SalesComercial=1 THEN SUM(LineAmountMST) ELSE 0 END [Inv.Amt.Bht. COM]");

        sbSql.AppendLine(",CASE WHEN ECL_SalesComercial=2 THEN SUM([SET]) ELSE 0 END [SET NOCOM]");
        sbSql.AppendLine(",CASE WHEN ECL_SalesComercial=2 THEN SUM(PCS) ELSE 0 END [PCS NOCOM]");
        sbSql.AppendLine(",CASE WHEN ECL_SalesComercial=2 THEN SUM(NW) ELSE 0 END [KG NOCOM]");
        sbSql.AppendLine(",0 [Inv.Amt. NOCOM]");
        sbSql.AppendLine(",0 [Inv.Amt.Bht. NOCOM]");
        sbSql.AppendLine(" FROM hoya_vwSalesDetail");


            sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");


            sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");

            sbSql.AppendLine(" AND INVOICEDATE = CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dtInvoiceDate) + "',103) ");
            //sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");

            if (InvoiceOBJ.ShipmentLocation == 6)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT' OR NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CINT')"); //CEXT EXT INT CINT

            }
            else if (InvoiceOBJ.ShipmentLocation == 1)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-INT' OR NUMBERSEQUENCEGROUP LIKE '%-CINT')"); //CEXT EXT INT CINT

            }
            else if (InvoiceOBJ.ShipmentLocation == 2)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-CEXT')"); //CEXT EXT INT CINT
            }

            else if (InvoiceOBJ.ShipmentLocation == 5)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE ('%-TRD') OR  NUMBERSEQUENCEGROUP LIKE ('%-CTRD'))"); //CEXT EXT INT CINT
            }
            else
            {

                sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP LIKE ('" + InvoiceOBJ.NumberSequenceGroup.Replace("-", "-") + "')");

            }

            if (locationID != "")
            {
                sbSql.AppendLine(" AND INVENTLOCATIONID IN ('"+locationID+"')");

            }

            sbSql.AppendLine(" GROUP BY LEDGERVOUCHER,INVOICEDATE,INVOICEID,CustName,SALESNAME,CURRENCYCODEISO,EXCHRATE,ECL_SalesComercial");

           sbSql.AppendLine(" ORDER BY LEDGERVOUCHER");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        } //End invoice by Date


        public ADODB.Recordset getInvoiceByCustomer(InvoiceOBJ InvoiceOBJ,string locationID)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;


        sbSql.AppendLine(" SELECT LEDGERVOUCHER,INVOICEDATE,INVOICEID");
        sbSql.AppendLine(" ,CASE WHEN CustName IS NULL THEN 'GRAND TOTAL' ELSE CustName END CustName");
        sbSql.AppendLine(" ,SALESNAME,CURRENCYCODEISO,EXCHRATE");
        sbSql.AppendLine(" ");
        sbSql.AppendLine(" ,SUM([SET COM]),SUM([PCS COM]),SUM([KG COM])");
        sbSql.AppendLine(" ,CASE WHEN (CURRENCYCODEISO IS NULL) THEN NULL ELSE SUM([Inv.Amt. COM]) END [Inv.Amt. COM]");
        sbSql.AppendLine(" ,SUM([Inv.Amt.Bht. COM])");
        sbSql.AppendLine(" ,SUM([SET NOCOM]),SUM([PCS NOCOM]),SUM([KG NOCOM])");
        sbSql.AppendLine(" ,CASE WHEN (CURRENCYCODEISO IS NULL) THEN NULL ELSE SUM([Inv.Amt. NOCOM]) END [Inv.Amt. NOCOM]");
        sbSql.AppendLine(" ,SUM([Inv.Amt.Bht. NOCOM])");
        sbSql.AppendLine(" FROM(");
        sbSql.AppendLine(" SELECT LEDGERVOUCHER,INVOICEID,INVOICEDATE,CustName,SALESNAME,CURRENCYCODEISO,EXCHRATE");
        sbSql.AppendLine(" ,CASE WHEN ECL_SalesComercial=1 THEN SUM([SET]) ELSE 0 END [SET COM]");
        sbSql.AppendLine(" ,CASE WHEN ECL_SalesComercial=1 THEN SUM(PCS) ELSE 0 END [PCS COM]");
        sbSql.AppendLine(" ,CASE WHEN ECL_SalesComercial=1 THEN SUM(NW) ELSE 0 END [KG COM]");
        sbSql.AppendLine(" ,CASE WHEN ECL_SalesComercial=1 THEN SUM(LineAmount) ELSE 0 END [Inv.Amt. COM]");
        sbSql.AppendLine(" ,CASE WHEN ECL_SalesComercial=1 THEN SUM(LineAmountMST) ELSE 0 END [Inv.Amt.Bht. COM]");
        sbSql.AppendLine(" ,CASE WHEN ECL_SalesComercial=2 THEN SUM([SET]) ELSE 0 END [SET NOCOM]");
        sbSql.AppendLine(" ,CASE WHEN ECL_SalesComercial=2 THEN SUM(PCS) ELSE 0 END [PCS NOCOM]");
        sbSql.AppendLine(" ,CASE WHEN ECL_SalesComercial=2 THEN SUM(NW) ELSE 0 END [KG NOCOM]");
        sbSql.AppendLine(" ,0 [Inv.Amt. NOCOM]");
        sbSql.AppendLine(" ,0 [Inv.Amt.Bht. NOCOM]");
        sbSql.AppendLine("  FROM hoya_vwSalesDetail");

        sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
        sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");



        if (InvoiceOBJ.ShipmentLocation == 6)
        {
            sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-INT'')"); 

        }
        else
        {
            sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP LIKE ('" + InvoiceOBJ.NumberSequenceGroup.Replace("-", "-") + "')");
        }
  

        if (locationID != "")
        {
            sbSql.AppendLine(" AND INVENTLOCATIONID = '"+locationID+"'");

        }
          
        sbSql.AppendLine(" GROUP BY LEDGERVOUCHER,INVOICEID,INVOICEDATE,CustName,SALESNAME,CURRENCYCODEISO,EXCHRATE,ECL_SalesComercial ");
        sbSql.AppendLine(" ) INVOICE_DATA");
        sbSql.AppendLine(" GROUP BY CustName,CURRENCYCODEISO,EXCHRATE,SALESNAME,LEDGERVOUCHER,INVOICEDATE,INVOICEID WITH ROLLUP");
        sbSql.AppendLine(" HAVING (EXCHRATE IS NULL) OR NOT(INVOICEID IS NULL) ");
        sbSql.AppendLine(" ORDER BY GROUPING(CustName),CustName,GROUPING(CURRENCYCODEISO),CURRENCYCODEISO");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        } //End invoice by Customer

        public ADODB.Recordset getInvoiceByItem(InvoiceOBJ InvoiceOBJ, bool comercial ,string locationID)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;


           
        sbSql.AppendLine(" SELECT [INVOICEDATE], ");
        sbSql.AppendLine("CASE WHEN (INVOICEID IS NULL)  THEN 'GRAND TOTAL'  ELSE [ECL_POSTALADDRESSID] END  [ECL_POSTALADDRESSID]");
        sbSql.AppendLine(",[INVOICEID]");
        sbSql.AppendLine(" ,SALESNAME");
        sbSql.AppendLine(",CASE WHEN ECL_SALESTOLOCATION IS NULL THEN 'TOTAL' ELSE ECL_SALESTOLOCATION END [ECL_SALESTOLOCATION] ");
        sbSql.AppendLine(",DLVTERM + ' ' + DLVREASON 'TRADETERM'");
        sbSql.AppendLine(" ,CURRENCYCODEISO");
        sbSql.AppendLine(" ,ECL_HOYAPONUMBER,ITEMID,ECL_GAIKEI,  SUM([SET]) [SET], SUM(PCS) QTY");
        sbSql.AppendLine(" ,SALESPRICE");
        sbSql.AppendLine(" ,SUM(LINEAMOUNT) LINEAMOUNT");
        sbSql.AppendLine(" ,EXCHRATE");
        sbSql.AppendLine(" ,SUM(LINEAMOUNTMST) LINEAMOUNTMST");
        sbSql.AppendLine(" ,SUM(NW) NW");
        sbSql.AppendLine(" ,SUM(GW) GW");
        sbSql.AppendLine(" ,CASE ECL_SalesComercial ");
        sbSql.AppendLine(" 	    WHEN 1 THEN 'Com'");
        sbSql.AppendLine(" 	    WHEN 2 THEN 'NO Com'");
        sbSql.AppendLine("  END COM");
        sbSql.AppendLine(" FROM HOYA_vwSalesDetail");


        sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
        sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");

            if (InvoiceOBJ.ShipmentLocation == 6)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-INT'')");
            }
            else
            {
                sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP LIKE ('" + InvoiceOBJ.NumberSequenceGroup.Replace("-", "-") + "')");
            }

            if (locationID != "")
            {
                sbSql.AppendLine(" AND INVENTLOCATIONID = '" + locationID + "'");

            }


            if (comercial)
            {
                sbSql.AppendLine(" AND ECL_SalesComercial=1"); //'Com
            }
            else
            {
                sbSql.AppendLine(" AND ECL_SalesComercial=2"); //'No Com
            }

        sbSql.AppendLine(" GROUP BY INVOICEDATE,ECL_POSTALADDRESSID,INVOICEID");
        sbSql.AppendLine(" ,SALESNAME,ECL_SALESTOLOCATION,DLVTERM,DLVREASON");
        sbSql.AppendLine(" ,CURRENCYCODEISO,ECL_HOYAPONUMBER,ITEMID,ECL_GAIKEI");
        sbSql.AppendLine(" ,SALESPRICE,EXCHRATE,ECL_SalesComercial WITH ROLLUP");
        sbSql.AppendLine("HAVING  (NOT [ECL_SALESCOMERCIAL] IS NULL) OR (INVOICEDATE IS NULL)OR (SALESNAME IS NULL) AND NOT INVOICEID IS NULL");

        sbSql.AppendLine("ORDER BY GROUPING([INVOICEDATE]),[INVOICEDATE],GROUPING([INVOICEID]),[INVOICEID]");
        sbSql.AppendLine(" ,GROUPING(SALESNAME),SALESNAME,GROUPING(ECL_SALESTOLOCATION),ECL_SALESTOLOCATION,GROUPING(DLVTERM),DLVTERM,GROUPING(DLVREASON),DLVREASON,");
        sbSql.AppendLine("GROUPING(CURRENCYCODEISO),CURRENCYCODEISO,GROUPING(ECL_HOYAPONUMBER),ECL_HOYAPONUMBER,GROUPING(ITEMID),[ITEMID],GROUPING(ECL_GAIKEI),[ECL_GAIKEI]");
        sbSql.AppendLine(",GROUPING(SALESPRICE),[SALESPRICE],GROUPING(EXCHRATE),[EXCHRATE],GROUPING(ECL_SALESCOMERCIAL),[ECL_SALESCOMERCIAL]");



            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        } //End invoice by Item

        public ADODB.Recordset getInvoiceByInvoice(InvoiceOBJ InvoiceOBJ, bool comercial, string locationID)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;


         sbSql.AppendLine("SELECT Sales.[INVOICEDATE] ");
        sbSql.AppendLine(",[ECL_POSTALADDRESSID]");
        sbSql.AppendLine(",Sales.INVOICEID");
        sbSql.AppendLine(",Sales.SALESNAME,Sales.ECL_SALESTOLOCATION,Sales.DLVTERM + ' ' + Sales.DLVREASON 'TRADETERM'");
        sbSql.AppendLine(",Sales.CURRENCYCODEISO");
        sbSql.AppendLine(",Sales.[SET]");
        sbSql.AppendLine(",Sales.QTY");
        sbSql.AppendLine(",Sales.LINEAMOUNT");
        sbSql.AppendLine(",EXCHRATE");
        sbSql.AppendLine(",Sales.LINEAMOUNTMST");
        sbSql.AppendLine(",CASE ECL_SALESIMPORTPACKING.PACKTYPE WHEN 1 THEN 'PLS' WHEN 2 THEN 'CTN' WHEN 3 THEN 'CASE' WHEN 4 THEN 'PACK' WHEN 0 THEN 'NONE' END");
        sbSql.AppendLine(",SUM(CONVERT(INT,ECL_SALESIMPORTPACKING.PACKNUM)) PACKNUM");
        sbSql.AppendLine(",Sales.NW");
        sbSql.AppendLine(",Sales.GW");
        sbSql.AppendLine(",Sales.COM");
        sbSql.AppendLine("FROM(");
        sbSql.AppendLine("SELECT ");
        sbSql.AppendLine("SALESID");
        sbSql.AppendLine(",INVOICEDATE");
        sbSql.AppendLine(",ECL_POSTALADDRESSID");
        sbSql.AppendLine(",INVOICEID");
        sbSql.AppendLine(",SALESNAME");
        sbSql.AppendLine(",ECL_SALESTOLOCATION");
        sbSql.AppendLine(",DLVTERM");
        sbSql.AppendLine(",DLVREASON");
        sbSql.AppendLine(",CURRENCYCODEISO");
        sbSql.AppendLine(",SUM([SET])[SET],SUM(PCS)QTY,SUM(LineAmount) LINEAMOUNT,EXCHRATE,SUM(LINEAMOUNTMST)LINEAMOUNTMST");
        sbSql.AppendLine(",SUM(NW) NW,SUM(GW)GW");
        sbSql.AppendLine(" ,CASE ECL_SalesComercial ");
        sbSql.AppendLine("WHEN 1 THEN 'Com'");
        sbSql.AppendLine("WHEN 2 THEN 'NO Com'");
        sbSql.AppendLine("  END COM");
        sbSql.AppendLine("FROM HOYA_vwSalesDetail ");

        sbSql.AppendLine(" WHERE INVENTSITEID='" + InvoiceOBJ.Factory + "'");
        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
        sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");

            if (InvoiceOBJ.ShipmentLocation == 6)
            {
                sbSql.AppendLine(" AND (NUMBERSEQUENCEGROUP LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-INT'')");
            }
            else
            {
                sbSql.AppendLine(" AND NUMBERSEQUENCEGROUP LIKE ('" + InvoiceOBJ.NumberSequenceGroup.Replace("-", "-") + "')");

            }

            if (locationID != "")
            {
                sbSql.AppendLine(" AND INVENTLOCATIONID = '" + locationID + "'");

            }


            if (comercial)
            {
                sbSql.AppendLine(" AND ECL_SalesComercial=1"); //'Com
            }
            else
            {
                sbSql.AppendLine(" AND ECL_SalesComercial=2"); //'No Com
            }

         sbSql.AppendLine("GROUP BY HOYA_vwSalesDetail.SALESID,INVOICEDATE,ECL_POSTALADDRESSID,INVOICEID,SALESNAME");
        sbSql.AppendLine(",ECL_SALESTOLOCATION,DLVTERM,DLVREASON,CURRENCYCODEISO,EXCHRATE,ECL_SALESCOMERCIAL");
        sbSql.AppendLine(")as Sales");
        sbSql.AppendLine("LEFT JOIN ECL_SALESIMPORTPACKING  ON Sales.SALESID= ECL_SALESIMPORTPACKING.SALESID");
        sbSql.AppendLine("GROUP BY Sales.INVOICEDATE,Sales.ECL_POSTALADDRESSID,Sales.INVOICEID,Sales.SALESNAME,Sales.ECL_SALESTOLOCATION");
        sbSql.AppendLine(" ,Sales.DLVTERM,Sales.DLVREASON,Sales.CURRENCYCODEISO,Sales.[SET],Sales.QTY,Sales.LINEAMOUNT");
        sbSql.AppendLine(" ,Sales.EXCHRATE,Sales.LINEAMOUNTMST,ECL_SALESIMPORTPACKING.PACKTYPE,Sales.NW,Sales.GW, Sales.COM WITH ROLLUP");
        sbSql.AppendLine(" HAVING  NOT COM IS NULL ");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        } //End invoice by Item



        public ADODB.Recordset getSalesByCustomer(InvoiceOBJ InvoiceOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(InvoiceOBJ.DateFrom.Year, InvoiceOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(InvoiceOBJ.DateTo.Year, InvoiceOBJ.DateTo.Month, 1);

            String strFac = InvoiceOBJ.strFactory;


            sbSql.AppendLine("SELECT salesdata.[Invoice Date],salesData.[Cust code],salesData.[Invoice No],salesData.[Address],salesData.[Des],salesData.[Trade Term]");
            sbSql.AppendLine(",SUM(salesData.[Amt])[AMT],salesData.[CURR],SUM(SalesDAta.[Ex Rate])[ExRate],SUM(SalesDAta.[Amt(Bht)])[AMTBHT]");
            sbSql.AppendLine(",CASE packing.PackType ");
            sbSql.AppendLine("WHEN 0 THEN 'None'");
            sbSql.AppendLine("WHEN 1 THEN 'PLS'");
            sbSql.AppendLine("WHEN 2 THEN 'CTN'");
            sbSql.AppendLine("WHEn 3 THEN 'CASE'");
            sbSql.AppendLine("WHEn 4 THEn 'PACK'");
            sbSql.AppendLine("END [PACKTYPE]");
            sbSql.AppendLine(",SUM(packing.packnum)[Packnum],SUM([PCS]) [PCS],SUM(packing.[NW])[NW],SUM(packing.[GW])[GW]");

            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("Select ");
            sbSql.AppendLine("Salestable.ECL_CustInvoiceDate [Invoice Date]");
            sbSql.AppendLine(",SalesTAble.CustAccount [Cust code]"); 
            sbSql.AppendLine(",custinvoicejour.invoiceid [Invoice No]");
            sbSql.AppendLine(",salestable.salesname [Address]");
            sbSql.AppendLine(",SalesTable.ECL_SalesToLocation [Des]");
            //sbSql.AppendLine(",Custinvoicejour.DLVTERM+' '+SalesTable.ECL_SalesToLocation [Trade Term]");
            sbSql.AppendLine(",SalesTable.DlvTerm +' '+SalesTable.ECL_FreightTo [Trade Term]");
            sbSql.AppendLine(",SUM (SALESLINE.LINEAMOUNT)[Amt]");
            sbSql.AppendLine(",CURRENCY.CURRENCYCODEISO [CURR]");
            sbSql.AppendLine(",CUSTINVOICEJOUR.[EXCHRATE]/100 [Ex Rate]");
            sbSql.AppendLine(",SUM (CASE WHEN CURRENCY.CURRENCYCODEISO = 'THB' THEN SALESLINE.LineAmount ELSE (SALESLINE.LineAmount) * CUSTINVOICEJOUR.[EXCHRATE] / 100 END)[Amt(Bht)]");
            sbSql.AppendLine(",SUM(SalesLine.SalesQty) [PCS]");
            sbSql.AppendLine("from Salestable inner join salesline ");
            sbSql.AppendLine("on salestable.salesid = salesline.salesid");
            sbSql.AppendLine("inner join custinvoicejour On custinvoicejour.salesid = salestable.salesid");
            sbSql.AppendLine("AND SALESTABLE.DATAAREAID = CUSTINVOICEJOUR.DATAAREAID ");
            sbSql.AppendLine("INNER JOIN CURRENCY ON CURRENCY.CURRENCYCODE = CUSTINVOICEJOUR.CURRENCYCODE");
            sbSql.AppendLine("WHERE ");
            sbSql.AppendLine("SalesTable.InventSiteId='" + InvoiceOBJ .Factory+"'");
            sbSql.AppendLine(" AND Salestable.ECL_CustInvoiceDate BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
           
            if (InvoiceOBJ.ShipmentLocation == 6)
            {
                sbSql.AppendLine(" AND ( SALESTABLE.NUMBERSEQUENCEGROUP   LIKE '%-EXT' OR NUMBERSEQUENCEGROUP LIKE '%-INT'')");
            }
            else
            {
                sbSql.AppendLine("   AND SALESTABLE.NUMBERSEQUENCEGROUP   LIKE ('" + InvoiceOBJ.NumberSequenceGroup.Replace("-", "-") + "')");

            }


            sbSql.AppendLine("GROUP BY Salestable.ECL_CustInvoiceDate,SalesTAble.CustAccount,custinvoicejour.invoiceid,salestable.salesname");
            sbSql.AppendLine(",SalesTable.ECL_SalesToLocation,SalesTable.ECL_SalesToLocation,SalesTable.DlvTerm,SalesTable.ECL_FreightTo");
            sbSql.AppendLine(",CURRENCY.CURRENCYCODEISO,CUSTINVOICEJOUR.[EXCHRATE] ) As salesdata");

            sbSql.AppendLine("LEFT OUTER JOIN(SELECT ");
            sbSql.AppendLine("Sumpack.ID,Sumpack.SalesID,SUM(Sumpack.NW)NW,SUM(Sumpack.GW)GW");
            sbSql.AppendLine(",Sumpack.PackTYPE,SUM(Sumpack.PackNUm)[PackNum]");
            sbSql.AppendLine(" FROM(");
            sbSql.AppendLine("SELECT CUSTINVOICEJOUR.INVOICEID AS ID, CUSTINVOICEJOUR.salesid");
            sbSql.AppendLine(",(ECL_SALESIMPORTPACKING.NW)[NW],(ECL_SALESIMPORTPACKING.GW)[GW]");
            sbSql.AppendLine(",ECL_SALESIMPORTPACKING.DATAAREAID,ECL_SALESIMPORTPACKING.PACKTYPE");
            sbSql.AppendLine(",CONVERT(INT,ECL_SALESIMPORTPACKING.PackNUM)[PackNum]");
            sbSql.AppendLine("FROM ECL_SALESIMPORTPACKING ");
            sbSql.AppendLine("INNER JOIN CUSTINVOICEJOUR ON ECL_SALESIMPORTPACKING.SALESID = CUSTINVOICEJOUR.SALESID ");
            sbSql.AppendLine("AND ECL_SALESIMPORTPACKING.DATAAREAID = CUSTINVOICEJOUR.DATAAREAID)as Sumpack");
            sbSql.AppendLine("GROUP BY Sumpack.ID,Sumpack.salesid,Sumpack.DATAAREAID,Sumpack.PackTYPE) AS packing ON packing.ID = salesdata.[Invoice No]");
            sbSql.AppendLine("GROUP BY salesdata.[Invoice Date],salesData.[Cust code],salesData.[Invoice No],salesData.[Address]");
            sbSql.AppendLine(",salesData.[Des],salesData.[Trade Term],salesData.[CURR],packing.PackType");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        } //End invoice by Customer





    }//end class
}
