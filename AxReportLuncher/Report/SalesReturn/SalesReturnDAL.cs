using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.Report.SalesReturn
{
    class SalesReturnDAL
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

        public DataTable getCustomerGroup()
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(" SELECT CustGroup,Name FROM CUSTGROUP");
            sbSql.AppendLine(" ORDER BY CustGroup");


            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }


        public ADODB.Recordset getSalesReturnBook(SalesReturnOBJ SalesReturnOBJ, bool boolComercial, string locationID)
        {
            
        StringBuilder sbSql = new StringBuilder();

       

     sbSql.AppendLine("SELECT CustInvoiceJour.LedgerVoucher as Voucher  ");
        sbSql.AppendLine(",CustInvoiceJour.InvoiceDate as DocDate ");
        sbSql.AppendLine(",SALESTABLE.ECL_CUSTINVOICENUMBER  as Invoice  ");
        sbSql.AppendLine(",SALESTABLE.ECL_CUSTINVOICEDATE as InvDate");
        sbSql.AppendLine(",SALESTABLE.INVOICEACCOUNT as CustCode");
        sbSql.AppendLine(",SALESTABLE.SALESID as SalesOrder");
            
        sbSql.AppendLine(",DIRPARTYTABLE.NAME AS [Customer name]");
        sbSql.AppendLine(",SALESTABLE.PAYMENT as Term");
        sbSql.AppendLine(",CURRENCYCODEISO as Curr");
        sbSql.AppendLine(",SALESTABLE.FIXEDEXCHRATE/100 [EX.Rate]");
        sbSql.AppendLine(",SUM(SALESLINE.LineAmount) *-1[Inv.Amt.Curr]");
        sbSql.AppendLine(",'' [VAT] ,SUM(SALESLINE.LINEAMOUNT)*(SALESTABLE.FIXEDEXCHRATE/100) *-1 [Inv.Amt.Bht]");
        sbSql.AppendLine(",''[Grand Total]");
        sbSql.AppendLine(" ,CUSTINVOICEJOUR.DUEDATE [Due Date]");
        sbSql.AppendLine(",SALESTABLE.ShippingDateRequested [AWB.Date]");
        sbSql.AppendLine(" ,tb_INVENTDIM.INVENTLOCATIONID [Location]");
        sbSql.AppendLine(" ,DimFinTag_sec.ECL_SHORTNAME [Crete by]");
        sbSql.AppendLine(" FROM SALESTABLE ");
        sbSql.AppendLine("INNER JOIN SALESLINE ON SALESLINE.SALESID = SALESTABLE.SALESID ");
        sbSql.AppendLine("AND SALESLINE.DATAAREAID = SALESTABLE.DATAAREAID ");
        sbSql.AppendLine("INNER JOIN INVENTTABLE ON SALESLINE.ITEMID = INVENTTABLE.ITEMID ");
        sbSql.AppendLine("AND SALESTABLE.DATAAREAID = INVENTTABLE.DATAAREAID");
        sbSql.AppendLine("LEFT JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID = CUSTINVOICEJOUR.SALESID	");
        sbSql.AppendLine("INNER JOIN CUSTTABLE CUSTDEST ON CUSTDEST.ACCOUNTNUM = SALESTABLE.CUSTACCOUNT ");
        sbSql.AppendLine("AND CUSTDEST.DATAAREAID = SALESTABLE.DATAAREAID");
        sbSql.AppendLine("INNER JOIN CUSTTABLE CUSTSOURCE ON CUSTSOURCE.ACCOUNTNUM = SALESTABLE.INVOICEACCOUNT ");
        sbSql.AppendLine("AND CUSTSOURCE.DATAAREAID = SALESTABLE.DATAAREAID ");
        sbSql.AppendLine("INNER JOIN DIRPARTYTABLE ON CUSTSOURCE.PARTY = DIRPARTYTABLE.RECID ");
        sbSql.AppendLine("LEFT JOIN CURRENCY ON CURRENCY.CURRENCYCODE = SALESTABLE.CURRENCYCODE ");
        sbSql.AppendLine("INNER JOIN HCMWORKER ON SALESTABLE.WORKERSALESTAKER = HCMWORKER.RECID");
        sbSql.AppendLine("INNER JOIN HcmWorkerTitle ON HcmWorkerTitle.WORKER=HCMWORKER.RECID");
        sbSql.AppendLine("INNER JOIN DirPersonName ON DirPersonName.PERSON=HCMWORKER.PERSON");
        sbSql.AppendLine("INNER JOIN HCMemployment ON HCMemployment.WORKER=hcmworker.RECID");
        sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM DimValSet_fac ON DimValSet_fac.DIMENSIONATTRIBUTEVALUESET=HCMemployment.DefaultDimension");
        sbSql.AppendLine("INNER JOIN DimensionAttributeValue DimVal_fac ON DimVal_fac.RECID=DimValSet_fac.DIMENSIONATTRIBUTEVALUE");
        sbSql.AppendLine("INNER JOIN DimensionFinancialTag DimFinTag_fac ON DimFinTag_fac.recid=DimVal_fac.ENTITYINSTANCE");
        sbSql.AppendLine("INNER JOIN DimensionAttribute Dim_fac ON Dim_fac.recid=DimVal_fac.DIMENSIONATTRIBUTE");
        sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM DimValSet_sec ON DimValSet_sec.DIMENSIONATTRIBUTEVALUESET=HCMemployment.DefaultDimension");
        sbSql.AppendLine("INNER JOIN DimensionAttributeValue DimVal_sec ON DimVal_sec.RECID=DimValSet_sec.DIMENSIONATTRIBUTEVALUE");
        sbSql.AppendLine("INNER JOIN DimensionFinancialTag DimFinTag_sec ON DimFinTag_sec.recid=DimVal_sec.ENTITYINSTANCE");
        sbSql.AppendLine("INNER JOIN DimensionAttribute Dim_sec ON Dim_sec.recid=DimVal_sec.DIMENSIONATTRIBUTE");
        sbSql.AppendLine("INNER JOIN HCMWORKERTASKASSIGNMENT ON HCMWORKERTASKASSIGNMENT.WORKER=HCMWORKER.RECID");
        sbSql.AppendLine("INNER JOIN HcmWorkerTask ON HcmWorkerTask.RECID=HCMWORKERTASKASSIGNMENT.WORKERTASK");

        // 11/9/2018
        sbSql.AppendLine("LEFT OUTER JOIN (SELECT inventdimid,INVENTLOCATIONID,DATAAREAID from INVENTDIM GROUP BY INVENTDIMID,INVENTLOCATIONID,DATAAREAID ) as tb_INVENTDIM");
        sbSql.AppendLine("ON SALESLINE.INVENTDIMID=tb_INVENTDIM.INVENTDIMID AND SALESLINE.DATAAREAID=tb_INVENTDIM.DATAAREAID");


        sbSql.AppendLine(" WHERE SALESTABLE.INVENTSITEID='" + SalesReturnOBJ.Factory + "'");
        sbSql.AppendLine(" AND SALESTABLE.createdDateTime BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", SalesReturnOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("      AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", SalesReturnOBJ.DateTo) + "',103)");
        sbSql.AppendLine(" AND SALESTABLE.CUSTGROUP IN ('" + SalesReturnOBJ.CustomerGroup + "')");

        if (SalesReturnOBJ.Factory == "GMO")
        {
            sbSql.AppendLine(" AND SALESTABLE.NUMBERSEQUENCEGROUP LIKE ('MO-R%" + "')");
        }
        else
        {
            sbSql.AppendLine(" AND SALESTABLE.NUMBERSEQUENCEGROUP LIKE ('" + SalesReturnOBJ.Factory + "-R%" + "')");
        }

        sbSql.AppendLine("AND SALESTABLE.SALESSTATUS!=4");
        sbSql.AppendLine("AND DirPersonName.VALIDTO>GETDATE() AND Dim_sec.NAME='D2_Section' AND Dim_fac.NAME='D1_Factory' ");
        sbSql.AppendLine("AND HcmWorkerTitle.VALIDTO>GETDATE()");
        sbSql.AppendLine("AND HCMemployment.VALIDTO>GETDATE()");
        sbSql.AppendLine("AND HcmWorkerTask.WORKERTASKID LIKE 'SalesReturnOnline%'");

        if (locationID != "")
        {
            sbSql.AppendLine(" AND tb_INVENTDIM.INVENTLOCATIONID = '" + locationID + "'");

        }

        sbSql.AppendLine("GROUP BY  CustInvoiceJour.LedgerVoucher ,CustInvoiceJour.InvoiceDate,SALESTABLE.ECL_CUSTINVOICENUMBER,SALESTABLE.ECL_CUSTINVOICEDATE");
        sbSql.AppendLine(",SALESTABLE.INVOICEACCOUNT ,SALESTABLE.SALESID,DIRPARTYTABLE.NAME,SALESTABLE.PAYMENT,CURRENCYCODEISO,SALESTABLE.FIXEDEXCHRATE ");
        sbSql.AppendLine(",CUSTINVOICEJOUR.DUEDATE ,SALESTABLE.ShippingDateRequested,tb_INVENTDIM.INVENTLOCATIONID,DimFinTag_sec.ECL_SHORTNAME");         
         ADODB.Recordset rs = new ADODB.Recordset();
         ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
         ADODBConnection.Open();
         object ret = null;
         rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

         return rs;
        }

        public ADODB.Recordset getSalesReturnByItem(SalesReturnOBJ SalesReturnOBJ, bool boolComercial, string locationID)
        {

            StringBuilder sbSql = new StringBuilder();

            
        sbSql.AppendLine("SELECT");
        sbSql.AppendLine("CASE WHEN [Customer] IS NULL THEN 'Grand Total' ELSE [Customer] END [Customer]");
        sbSql.AppendLine(" ,CASE WHEN [Item sale]IS NULL AND NOT [Customer] IS NULL THEN 'Total' ELSE [Item sale] END [Item sale]");
        sbSql.AppendLine(",summary.[Group Code]");
        sbSql.AppendLine(",SUM(summary.[Pcs.]) *-1 [PCS]");
        sbSql.AppendLine(",Sum(summary.[Amount Baht]*summary.[ExRate.]) [AmtBht] ");
        sbSql.AppendLine(",'COM' [COM]");

        sbSql.AppendLine(" FROM(");

        sbSql.AppendLine("SELECT");
        sbSql.AppendLine("SALESNAME  [Customer]");
        sbSql.AppendLine(",[RELATESALESITEM]  [Item sale]");
        sbSql.AppendLine(",ECL_GROUPCODE  [Group Code]");
        sbSql.AppendLine(",SUM([PCS]) as [Pcs.]");
        sbSql.AppendLine(",[Exc.rate] [ExRate.]");
        sbSql.AppendLine(",[Currency]");
        sbSql.AppendLine(",SUM([InvACurr]*-1) [Amt.CURR]");
        sbSql.AppendLine(",Sum([InvABht]*-1) [Amount Baht]");

        sbSql.AppendLine("FROM(");
        sbSql.AppendLine("SELECT");
        sbSql.AppendLine("SALESNAME [SALESNAME]");
        sbSql.AppendLine(",INVENTTABLE.HOYA_RELATESALESITEM [RELATESALESITEM]");
        sbSql.AppendLine(",ECL_GROUPCODE");
        sbSql.AppendLine(",CURRENCYCODEISO  [Currency]");
        sbSql.AppendLine(",SALESTABLE.FIXEDEXCHRATE /100[Exc.rate]");
        sbSql.AppendLine(",SUM(SALESLINE.SalesQty) [PCS]");
        sbSql.AppendLine(",SUM(SALESLINE.LineAmount) [InvACurr]");
        sbSql.AppendLine(",SUM(SALESLINE.LINEAMOUNT) [InvABht]");
        sbSql.AppendLine("FROM SALESLINE");
        sbSql.AppendLine("INNER JOIN INVENTTABLE ON SALESLINE.ITEMID = INVENTTABLE.ITEMID AND SALESLINE.DATAAREAID = INVENTTABLE.DATAAREAID");
        sbSql.AppendLine("INNER JOIN CUSTINVOICEJOUR ON SALESLINE.SALESID = CUSTINVOICEJOUR.SALESID");
        sbSql.AppendLine("INNER JOIN CURRENCY ON CURRENCY.CURRENCYCODE = SALESLINE.CURRENCYCODE");
        sbSql.AppendLine("INNER JOIN SALESTABLE ON SALESTABLE.SALESID = SALESLINE.SALESID AND  SALESTABLE.DATAAREAID = SALESLINE.DATAAREAID");

        // 11/9/2018
        sbSql.AppendLine("LEFT OUTER JOIN (SELECT inventdimid,INVENTLOCATIONID,DATAAREAID from INVENTDIM GROUP BY INVENTDIMID,INVENTLOCATIONID,DATAAREAID ) as tb_INVENTDIM");
        sbSql.AppendLine("ON SALESLINE.INVENTDIMID=tb_INVENTDIM.INVENTDIMID AND SALESLINE.DATAAREAID=tb_INVENTDIM.DATAAREAID");

        sbSql.AppendLine(" WHERE  SALESTABLE.INVENTSITEID='" + SalesReturnOBJ.Factory + "'");
        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesReturnOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesReturnOBJ.DateTo) + "',103)");
        sbSql.AppendLine(" AND SALESTABLE.CUSTGROUP IN ('" + SalesReturnOBJ.CustomerGroup + "')");

            if (boolComercial)
            {
                sbSql.AppendLine("AND ECL_SalesComercial=1");
            }
            else
            {
                sbSql.AppendLine("AND ECL_SalesComercial=2");
            }

            if (SalesReturnOBJ.Factory == "GMO")
            {
                _strFac = "MO";
            }
            else
            { 
                _strFac = SalesReturnOBJ.Factory;
            }

            ShipLoc = SalesReturnOBJ.ShipmentLocation;

            if (ShipLoc == 1)
            {
                sbSql.AppendLine(" AND  SALESTABLE.NUMBERSEQUENCEGROUP = ('" + _strFac + "-RTD%" + "')");
                sbSql.AppendLine(" AND  SALESTABLE.NUMBERSEQUENCEGROUP != ('" + _strFac + "-RNOC" + "')");
            }
            else
            {
                sbSql.AppendLine(" AND ( SALESTABLE.NUMBERSEQUENCEGROUP = ('" + _strFac + "-REXT" + "')");
                sbSql.AppendLine(" OR   SALESTABLE.NUMBERSEQUENCEGROUP = ('" + _strFac + "-RINT" + "'))");
                sbSql.AppendLine(" AND  SALESTABLE.NUMBERSEQUENCEGROUP != ('" + _strFac + "-RNOC" + "')");
            }

     if (locationID != "")
        {
            sbSql.AppendLine(" AND tb_INVENTDIM.INVENTLOCATIONID = '" + locationID + "'");

        }



      sbSql.AppendLine("   GROUP BY SALESNAME,INVENTTABLE.HOYA_RELATESALESITEM,ECL_GROUPCODE,CURRENCYCODEISO,FIXEDEXCHRATE");
      sbSql.AppendLine(" ) INVOICE_DATA");
      sbSql.AppendLine("  GROUP BY SALESNAME,[RELATESALESITEM],ECL_GROUPCODE,[Currency],[Exc.rate]");
      sbSql.AppendLine(")summary ");
      sbSql.AppendLine("   GROUP BY summary.Customer,summary.[Item sale],summary.[Group Code],summary.Currency");
      sbSql.AppendLine("  WITH ROLLUP HAVING (NOT [Currency]  IS NULL) OR ( [Item sale] IS NULL)");

 
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();
            object ret;
            rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

            return rs;


        }

        public ADODB.Recordset getSaleReturnByCustomer(SalesReturnOBJ SalesReturnOBJ, bool boolComercial,string locationID)
        {

            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine("SELECT");
        sbSql.AppendLine("CASE WHEN [Customer] IS NULL THEN 'Grand Total' ELSE [Customer] END [Customer]");
        sbSql.AppendLine(" ,CASE WHEN [Item sale]IS NULL AND NOT [Customer] IS NULL THEN 'Total' ELSE [Item sale] END [Item sale]");
        sbSql.AppendLine(",summary.[Group Code]");
        sbSql.AppendLine(",SUM(summary.[Pcs.]) *-1 [PCS]");

       sbSql.AppendLine(",CASE WHEN  [Item sale] IS NULL THEN NULL ELSE  SUM(summary.[ExRate.]) END [EXRATE]");
        sbSql.AppendLine(",summary.Currency");
        sbSql.AppendLine(",SUM(summary.[Amt.CURR])");
        sbSql.AppendLine(",Sum(summary.[Amount Baht]*summary.[ExRate.]) [AmtBht] ");
        sbSql.AppendLine(" FROM(");

        sbSql.AppendLine("SELECT");
        sbSql.AppendLine("SALESNAME  [Customer]");
        sbSql.AppendLine(",[RELATESALESITEM]  [Item sale]");
        sbSql.AppendLine(",ECL_GROUPCODE  [Group Code]");
        sbSql.AppendLine(",SUM([PCS]) as [Pcs.]");
        sbSql.AppendLine(",[Exc.rate] [ExRate.]");
        sbSql.AppendLine(",[Currency]");
        sbSql.AppendLine(",SUM([InvACurr]*-1) [Amt.CURR]");
        sbSql.AppendLine(",Sum([InvABht]*-1) [Amount Baht]");

        sbSql.AppendLine("FROM(");
        sbSql.AppendLine("SELECT");
        sbSql.AppendLine("SALESNAME [SALESNAME]");
        sbSql.AppendLine(",INVENTTABLE.HOYA_RELATESALESITEM [RELATESALESITEM]");
        sbSql.AppendLine(",ECL_GROUPCODE");
        sbSql.AppendLine(",CURRENCYCODEISO  [Currency]");
        sbSql.AppendLine(",SALESTABLE.FIXEDEXCHRATE /100[Exc.rate]");
        sbSql.AppendLine(",SUM(SALESLINE.SalesQty) [PCS]");
        sbSql.AppendLine(",SUM(SALESLINE.LineAmount) [InvACurr]");
        sbSql.AppendLine(",SUM(SALESLINE.LINEAMOUNT) [InvABht]");
        sbSql.AppendLine("FROM SALESLINE");
        sbSql.AppendLine("INNER JOIN INVENTTABLE ON SALESLINE.ITEMID = INVENTTABLE.ITEMID AND SALESLINE.DATAAREAID = INVENTTABLE.DATAAREAID");
        sbSql.AppendLine("INNER JOIN CUSTINVOICEJOUR ON SALESLINE.SALESID = CUSTINVOICEJOUR.SALESID");
        sbSql.AppendLine("INNER JOIN CURRENCY ON CURRENCY.CURRENCYCODE = SALESLINE.CURRENCYCODE");
        sbSql.AppendLine("INNER JOIN SALESTABLE ON SALESTABLE.SALESID = SALESLINE.SALESID AND  SALESTABLE.DATAAREAID = SALESLINE.DATAAREAID");

        // 11/9/2018
        sbSql.AppendLine("LEFT OUTER JOIN (SELECT inventdimid,INVENTLOCATIONID,DATAAREAID from INVENTDIM GROUP BY INVENTDIMID,INVENTLOCATIONID,DATAAREAID ) as tb_INVENTDIM");
        sbSql.AppendLine("ON SALESLINE.INVENTDIMID=tb_INVENTDIM.INVENTDIMID AND SALESLINE.DATAAREAID=tb_INVENTDIM.DATAAREAID");



        sbSql.AppendLine(" WHERE SALESTABLE.INVENTSITEID='" + SalesReturnOBJ.Factory + "'");
        sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesReturnOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("      AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", SalesReturnOBJ.DateTo) + "',103)");

        sbSql.AppendLine(" AND  SALESTABLE.CUSTGROUP IN ('" + SalesReturnOBJ.CustomerGroup + "')");
        sbSql.AppendLine("AND ECL_SalesComercial=1");

            if (SalesReturnOBJ.Factory == "GMO")
            {

                _strFac = "MO";
            }
            else
            {

                _strFac = SalesReturnOBJ.Factory;
            }

            ShipLoc = SalesReturnOBJ.ShipmentLocation;

            if (ShipLoc == 1)
            {
                sbSql.AppendLine(" AND SALESTABLE.NUMBERSEQUENCEGROUP = ('" + _strFac + "-RTD%" + "')");
                sbSql.AppendLine(" AND SALESTABLE.NUMBERSEQUENCEGROUP != ('" + _strFac + "-RNOC" + "')");
            }
            else
            {
                sbSql.AppendLine(" AND (SALESTABLE.NUMBERSEQUENCEGROUP = ('" + _strFac + "-REXT" + "')");
                sbSql.AppendLine(" OR  SALESTABLE.NUMBERSEQUENCEGROUP = ('" + _strFac + "-RINT" + "'))");
                sbSql.AppendLine(" AND SALESTABLE.NUMBERSEQUENCEGROUP != ('" + _strFac + "-RNOC" + "')");
            }


            if (locationID != "")
            {
                sbSql.AppendLine(" AND tb_INVENTDIM.INVENTLOCATIONID = '" + locationID + "'");

            }


        sbSql.AppendLine("   GROUP BY SALESNAME,INVENTTABLE.HOYA_RELATESALESITEM,ECL_GROUPCODE,CURRENCYCODEISO,FIXEDEXCHRATE");
        sbSql.AppendLine(" ) INVOICE_DATA");
        sbSql.AppendLine("  GROUP BY SALESNAME,[RELATESALESITEM],ECL_GROUPCODE,[Currency],[Exc.rate]");
        sbSql.AppendLine(")summary ");
        sbSql.AppendLine("   GROUP BY summary.Customer,summary.[Item sale],summary.[Group Code],summary.Currency");
        sbSql.AppendLine("  WITH ROLLUP HAVING (NOT [Currency]  IS NULL) OR ( [Item sale] IS NULL)");


           
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();
            object ret;
          
            rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

            return rs;


        }

        public ADODB.Recordset getSalesReturnRemainReport(SalesReturnOBJ SalesReturnOBJ, bool boolComercial, string locationID)
        {

            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine("SELECT");
        sbSql.AppendLine("CUSTINVOICEJOUR.INVOICEDATE [Receive Date]");
        sbSql.AppendLine(",CustInvoiceJour.LedgerVoucher[Receive NO]");
        sbSql.AppendLine(",SALESTABLE.CREATEDDATETIME [SO Date]");
        sbSql.AppendLine(",SALESTABLE.SALESID [SO NO.]");
        sbSql.AppendLine(",SALESTABLE.INVOICEACCOUNT [Cust.Code]");
        sbSql.AppendLine(",SALESTABLE.SalesName [Customer name]");
        sbSql.AppendLine(",DimFinTag_sec.VALUE [Sec.]");
        sbSql.AppendLine(",SALESLINE.ITEMID [Item Id]");
        sbSql.AppendLine(",SALESLINE.NAME [Name]");
       // sbSql.AppendLine(",CASE WHEN SalesLine.SalesQty < 0 THEN  SalesLine.SalesQty *-1  ELSE  SalesLine.SalesQty END   [Qty]");
        sbSql.AppendLine(", SalesLine.SalesQty *-1  [Qty]");
        sbSql.AppendLine(",SalesLine.SalesUnit [Unit]");
        sbSql.AppendLine(",SalesTable.CurrencyCode [Cur]");
        sbSql.AppendLine(",SalesLine.SalesPrice [Unit Price]");

        sbSql.AppendLine(", CustInvoiceTrans.Qty * -1  [Qty Rcpt]");
        sbSql.AppendLine(",CustInvoiceTrans.LINEAMOUNTTAX *-1 [Total Rcpt]");
        //sbSql.AppendLine(",CASE WHEN CustInvoiceTrans.Qty <0 THEN CustInvoiceTrans.Qty * -1 ELSE CustInvoiceTrans.Qty END  [Qty Rcpt]");
        //sbSql.AppendLine(",CASE WHEN CustInvoiceTrans.LINEAMOUNTTAX <0 THEN CustInvoiceTrans.LINEAMOUNTTAX *-1 ELSE CustInvoiceTrans.LINEAMOUNTTAX END  [Total Rcpt]");


        sbSql.AppendLine(",''[Qty Remain]");
        sbSql.AppendLine(",''[Total Remain]");
        sbSql.AppendLine(",SalesTable.DlvTerm [Inco.Tm]");
        sbSql.AppendLine(",SalesTable.Payment [Paym.Tm]");


        sbSql.AppendLine(",DimFinTag_fac.ECL_SHORTNAME [Crete by]");
        sbSql.AppendLine(",CASE WHEN SALESTABLE.SalesStatus = '1' THEN 'Open'");
        sbSql.AppendLine(" WHEN SALESTABLE.SalesStatus = '3' THEN 'Invoiced'");
        sbSql.AppendLine("WHEN SALESTABLE.SalesStatus = '4' THEN 'Canceled'");
        sbSql.AppendLine("END [STATUS] 	");

        sbSql.AppendLine("FROM SALESTABLE");
        sbSql.AppendLine("LEFT OUTER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID = CUSTINVOICEJOUR.SALESID ");
        sbSql.AppendLine("AND SALESTABLE.DATAAREAID = CUSTINVOICEJOUR.DATAAREAID  ");
        sbSql.AppendLine("INNER JOIN CUSTTABLE CUSTDEST ON CUSTDEST.ACCOUNTNUM = SALESTABLE.CUSTACCOUNT");
        sbSql.AppendLine("AND CUSTDEST.DATAAREAID = SALESTABLE.DATAAREAID ");
        sbSql.AppendLine("INNER JOIN CUSTTABLE CUSTSOURCE ON CUSTSOURCE.ACCOUNTNUM = SALESTABLE.INVOICEACCOUNT");
        sbSql.AppendLine("AND CUSTSOURCE.DATAAREAID = SALESTABLE.DATAAREAID ");
        sbSql.AppendLine("INNER JOIN DIRPARTYTABLE ON CUSTSOURCE.PARTY = DIRPARTYTABLE.RECID ");
        sbSql.AppendLine("LEFT OUTER JOIN CURRENCY ON CURRENCY.CURRENCYCODE = CUSTINVOICEJOUR.CURRENCYCODE");
        sbSql.AppendLine("INNER JOIN SALESLINE ON SALESLINE.SALESID = SALESTABLE.SALESID AND SALESLINE.DATAAREAID = SALESTABLE.DATAAREAID ");
        sbSql.AppendLine("LEFT OUTER JOIN INVENTLOCATION ON SALESTABLE.INVENTLOCATIONID = INVENTLOCATION.INVENTLOCATIONID ");
        sbSql.AppendLine("AND SALESTABLE.DATAAREAID = INVENTLOCATION.DATAAREAID");

        sbSql.AppendLine("INNER JOIN HCMWORKER ON SALESTABLE.WORKERSALESTAKER = HCMWORKER.RECID");
        sbSql.AppendLine("INNER JOIN HcmWorkerTitle ON HcmWorkerTitle.WORKER=HCMWORKER.RECID");
        sbSql.AppendLine("INNER JOIN DirPersonName ON DirPersonName.PERSON=HCMWORKER.PERSON");
        sbSql.AppendLine("INNER JOIN HCMemployment ON HCMemployment.WORKER=hcmworker.RECID");



        sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM DimValSet_fac ON DimValSet_fac.DIMENSIONATTRIBUTEVALUESET=HCMemployment.DefaultDimension");
        sbSql.AppendLine("INNER JOIN DimensionAttributeValue DimVal_fac ON DimVal_fac.RECID=DimValSet_fac.DIMENSIONATTRIBUTEVALUE");
        sbSql.AppendLine("INNER JOIN DimensionFinancialTag DimFinTag_fac ON DimFinTag_fac.recid=DimVal_fac.ENTITYINSTANCE");
        sbSql.AppendLine("INNER JOIN DimensionAttribute Dim_fac ON Dim_fac.recid=DimVal_fac.DIMENSIONATTRIBUTE");

        sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM DimValSet_sec ON DimValSet_sec.DIMENSIONATTRIBUTEVALUESET=SALESLINE.DefaultDimension");
        sbSql.AppendLine("INNER JOIN DimensionAttributeValue DimVal_sec ON DimVal_sec.RECID=DimValSet_sec.DIMENSIONATTRIBUTEVALUE");
        sbSql.AppendLine("INNER JOIN DimensionFinancialTag DimFinTag_sec ON DimFinTag_sec.recid=DimVal_sec.ENTITYINSTANCE");
        sbSql.AppendLine("INNER JOIN DimensionAttribute Dim_sec ON Dim_sec.recid=DimVal_sec.DIMENSIONATTRIBUTE");

            
            
            sbSql.AppendLine("LEFT OUTER JOIN CustInvoiceTrans ON SalesLine.SalesId=CustInvoiceTrans.ORIGSALESID");
        sbSql.AppendLine("AND SalesLine.LineNum=CustInvoiceTrans.LineNum");
        sbSql.AppendLine("AND SalesLine.DataAreaId=CustInvoiceTrans.DataAreaId");
        sbSql.AppendLine("LEFT OUTER JOIN (SELECT TransRecId,Value Value,DATAAREAID,MarkupTrans.RECID,MarkupTrans.ORIGRECID FROM MarkupTrans  WHERE TransTableId=64 /*CustInvoiceTrans*/) MarkupTrans ON CustInvoiceTrans.RecId=MarkupTrans.TRANSRECID");
        sbSql.AppendLine("AND CustInvoiceTrans.DATAAREAID=MarkupTrans.DATAAREAID");

        // 11/9/2018
        sbSql.AppendLine("LEFT OUTER JOIN (SELECT inventdimid,INVENTLOCATIONID,DATAAREAID from INVENTDIM GROUP BY INVENTDIMID,INVENTLOCATIONID,DATAAREAID ) as tb_INVENTDIM");
        sbSql.AppendLine("ON SALESLINE.INVENTDIMID=tb_INVENTDIM.INVENTDIMID AND SALESLINE.DATAAREAID=tb_INVENTDIM.DATAAREAID");


         sbSql.AppendLine(" WHERE SALESTABLE.INVENTSITEID='" + SalesReturnOBJ.Factory + "'");


         sbSql.AppendLine(" AND SALESTABLE.createdDateTime BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", SalesReturnOBJ.DateFrom) + "',103) ");
         sbSql.AppendLine("      AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", SalesReturnOBJ.DateTo) + "',103)");

         sbSql.AppendLine(" AND  SALESTABLE.CUSTGROUP IN ('" + SalesReturnOBJ.CustomerGroup + "')");

         sbSql.AppendLine("AND DirPersonName.VALIDTO>GETDATE() ");
         sbSql.AppendLine("AND Dim_sec.NAME='D3_Subsection'  AND Dim_fac.NAME='D2_Section'  ");
         sbSql.AppendLine("AND HcmWorkerTitle.VALIDTO>GETDATE()");
         sbSql.AppendLine("AND HCMemployment.VALIDTO>GETDATE()");


         //sbSql.AppendLine("AND CustInvoiceTrans.Qty != 0");
         //sbSql.AppendLine("AND DimFinTag_sec.ECL_SHORTNAME !='CONT'");


            if (SalesReturnOBJ.Factory == "GMO")
            {
                sbSql.AppendLine(" AND SALESTABLE.NUMBERSEQUENCEGROUP LIKE ('MO-R%" + "')");
            }
            else
            {
                sbSql.AppendLine(" AND SALESTABLE.NUMBERSEQUENCEGROUP LIKE ('" + SalesReturnOBJ.Factory + "-R%" + "')");
            }



            if (locationID != "")
            {
                sbSql.AppendLine(" AND tb_INVENTDIM.INVENTLOCATIONID = '" + locationID + "'");

            }


            sbSql.AppendLine("ORDER BY CUSTINVOICEJOUR.INVOICEDATE,CustInvoiceJour.LedgerVoucher,SALESLINE.SALESID");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();
            object ret;

            rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

            return rs;


        }





    }
}
