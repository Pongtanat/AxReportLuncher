using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.Report.PaymentGeneralReport
{
    class PaymentGeneralDAL
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

        public ADODB.Recordset getPaymentGeneral(PaymentGeneralOBJ PaymentGeneralOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(PaymentGeneralOBJ.DateFrom.Year, PaymentGeneralOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(PaymentGeneralOBJ.DateTo.Year, PaymentGeneralOBJ.DateTo.Month, 1);

            String strFac = PaymentGeneralOBJ.strFactory;

            if (PaymentGeneralOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = PaymentGeneralOBJ.strFactory;
            }


            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" VENDINVOICEJOUR.LEDGERVOUCHER  [Voucher]");
            sbSql.AppendLine(" ,VENDINVOICETRANS.INVOICEID [InvoiceID]");
            sbSql.AppendLine(",VENDINVOICETRANS.INVOICEDATE  [Voucher Date]");
            sbSql.AppendLine(",VENDINVOICETRANS.PURCHID [Po No.]");
            

            //sbSql.AppendLine(",MAX(VENDTRANS.DOCUMENTDATE) [PO Date]");
            sbSql.AppendLine(",MAX(VENDPURCHN.PURCHORDERDATE) [PO Date]");

            sbSql.AppendLine(" ,VENDINVOICEJOUR.ORDERACCOUNT [Vend Code]");
            sbSql.AppendLine(",VENDDIRPARTYTABLEVIEW.NAME [Purch name]");
            
            sbSql.AppendLine(",VENDINVOICEJOUR.CURRENCYCODE [Curr]");

            //sbSql.AppendLine(" ,ISNULL(VENDINVOICEJOUR.ExchRate,0)/100 EXCHRATE");
            sbSql.AppendLine(",CONVERT(decimal(10,5),ISNULL(VENDINVOICEJOUR.ExchRate,0)/100 ) EXCHRATE");
            sbSql.AppendLine(",MAX(PURCHTABLE.PAYMENT) TERM");
            //11/26/2018
            //sbSql.AppendLine(",SUM(ISNULL(VENDINVOICETRANS.LINEAMOUNT,0)) AmtCurr");
            sbSql.AppendLine(",SUM(ISNULL(VENDINVOICETRANS.LINEAMOUNT,0)+ISNULL(DISCOUNT.VALUE,0)) AmtCurr");

            sbSql.AppendLine(",SUM(ISNULL(VENDINVOICETRANS.TAXAMOUNT,0) * ISNULL(VENDINVOICEJOUR.ExchRate,0)) / 100 VAT");

            //11/26/2018
            //sbSql.AppendLine(" ,(SUM(ISNULL(VENDINVOICETRANS.LINEAMOUNT,0))+ISNULL(VENDINVOICEJOUR.SUMTAX,0))*(ISNULL(VENDINVOICEJOUR.EXCHRATE,0)/100) INVOICEAMOUNTMST");
            //sbSql.AppendLine(",SUM(ISNULL(VENDINVOICETRANS.LINEAMOUNT,0)+ISNULL(DISCOUNT.VALUE,0))+(SUM(ISNULL(VENDINVOICETRANS.TAXAMOUNT,0)) * (ISNULL(VENDINVOICEJOUR.ExchRate,0) / 100)) INVOICEAMOUNTMST");
            sbSql.AppendLine(",SUM(ISNULL(VENDINVOICETRANS.LINEAMOUNTMST,0)+ISNULL(DISCOUNT.VALUE,0))+(SUM(ISNULL(VENDINVOICETRANS.TAXAMOUNT,0)) * (ISNULL(VENDINVOICEJOUR.ExchRate,0) / 100)) INVOICEAMOUNTMST");

            sbSql.AppendLine(",SUM(ISNULL(FREIGHT.VALUE,0)) FREIGHT");
            sbSql.AppendLine(",SUM(ISNULL(INNSURANCE.VALUE,0))INNSURANCE");

            //11/26/2018
           // sbSql.AppendLine(" ,(SUM(ISNULL(VENDINVOICETRANS.LINEAMOUNT,0))+ISNULL(VENDINVOICEJOUR.SUMTAX,0))*(ISNULL(VENDINVOICEJOUR.EXCHRATE,0)/100)+SUM(ISNULL(FREIGHT.VALUE,0))+SUM(ISNULL(INNSURANCE.VALUE,0)) GRANDTOTAL");
            sbSql.AppendLine(",SUM(ISNULL(VENDINVOICETRANS.LINEAMOUNT,0)+ISNULL(DISCOUNT.VALUE,0))+(SUM(ISNULL(VENDINVOICETRANS.TAXAMOUNT,0) * ISNULL(VENDINVOICEJOUR.ExchRate,0)) / 100) +SUM(ISNULL(FREIGHT.VALUE,0))+SUM(ISNULL(INNSURANCE.VALUE,0)) GRANDTOTAL");

            sbSql.AppendLine(" ,VENDTRANSOPEN.ECL_WHTAXAMOUNT WHT");
            sbSql.AppendLine(" ,SUM(VENDINVOICETRANS.LINEAMOUNT)+(SUM(ISNULL(VENDINVOICETRANS.TAXAMOUNT,0) * ISNULL(VENDINVOICEJOUR.ExchRate,0)) / 100) +SUM(ISNULL(FREIGHT.VALUE,0))+SUM(ISNULL(INNSURANCE.VALUE,0))+ISNULL(VENDTRANSOPEN.ECL_WHTAXAMOUNT,0) NetPay");

            sbSql.AppendLine(",LASTSETTLEVOUCHER  [Last Settlement]");
            sbSql.AppendLine(",LASTSETTLEDATE [PayMent Date]");
            sbSql.AppendLine(",CASE PURCHTABLE.PURCHSTATUS");
            sbSql.AppendLine("WHEN 0 THEN 'None'");
            sbSql.AppendLine("WHEN 1 THEN 'Open order' ");
            sbSql.AppendLine("WHEN 2 THEN 'Received'");
            sbSql.AppendLine(" WHEN 3 THEN CASE WHEN VendTrans.Closed != CONVERT(datetime,'1900-01-01',103)  THEN 'Paid' ELSE '' END ");
            sbSql.AppendLine(" WHEN 4 THEN 'Canceled'");
            sbSql.AppendLine("END [PO_LINE_STATUS]");
  
            
            
            sbSql.AppendLine("FROM VENDINVOICETRANS ");
            sbSql.AppendLine(" INNER JOIN INVENTDIM on VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine("INNER JOIN VENDINVOICEJOUR ON ");
            sbSql.AppendLine("VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE=VENDINVOICEJOUR.INVOICEDATE AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP=VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID=VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=VENDINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN VENDTRANS ON VENDINVOICEJOUR.LEDGERVOUCHER=VENDTRANS.VOUCHER");
            sbSql.AppendLine("AND VENDINVOICEJOUR.INVOICEACCOUNT=VENDTRANS.ACCOUNTNUM");
            sbSql.AppendLine("AND VENDINVOICEJOUR.INVOICEDATE=VENDTRANS.TRANSDATE");
            sbSql.AppendLine("AND VENDINVOICEJOUR.DATAAREAID=VENDTRANS.DATAAREAID");
            sbSql.AppendLine(" LEFT OUTER JOIN VENDTRANSOPEN ON VENDTRANS.RECID=VENDTRANSOPEN.REFRECID");
            sbSql.AppendLine("AND VENDTRANS.DATAAREAID=VENDTRANSOPEN.DATAAREAID");
            sbSql.AppendLine(" LEFT OUTER JOIN (SELECT NUM,INVOICEACCOUNT,VENDINVOICESAVESTATUS,MAX(DOCUMENTDATE) DOCUMENTDATE,MAX(HOYA_AWBDATE) AWBDATE,DATAAREAID ");
            sbSql.AppendLine("FROM VENDINVOICEINFOTABLE GROUP BY NUM,INVOICEACCOUNT,VENDINVOICESAVESTATUS,DATAAREAID) VENDINVOICEINFOTABLE");
            sbSql.AppendLine("ON VENDINVOICEJOUR.INVOICEID=VENDINVOICEINFOTABLE.NUM");
            sbSql.AppendLine("AND VENDINVOICEJOUR.INVOICEACCOUNT=VENDINVOICEINFOTABLE.INVOICEACCOUNT");
            sbSql.AppendLine("AND VENDINVOICEJOUR.DATAAREAID=VENDINVOICEINFOTABLE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN PURCHTABLE ON PURCHTABLE.PURCHID=VENDINVOICETRANS.ORIGPURCHID");
            sbSql.AppendLine(" AND PURCHTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN PURCHLINE ON PURCHTABLE.PURCHID=PURCHLINE.PURCHID AND VENDINVOICETRANS.INVENTTRANSID=PURCHLINE.INVENTTRANSID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=PURCHLINE.DATAAREAID");

            sbSql.AppendLine("INNER JOIN NUMBERSEQUENCEGROUP ON PURCHTABLE.NUMBERSEQUENCEGROUP=NUMBERSEQUENCEGROUP.NUMBERSEQUENCEGROUPID");
            sbSql.AppendLine("AND NUMBERSEQUENCEGROUP.NUMGROUPDATAAREAID = PURCHTABLE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN VENDTABLE on vendtable.ACCOUNTNUM=vendInvoiceJour.InvoiceAccount");
            sbSql.AppendLine("AND vendtable.DATAAREAID=vendInvoiceJour.DATAAREAID");
            sbSql.AppendLine("INNER JOIN VENDDIRPARTYTABLEVIEW on  VENDDIRPARTYTABLEVIEW.PARTY = VENDTABLE.PARTY");

            //12-06-2018
            sbSql.AppendLine("INNER JOIN (SELECT PURCHID,DATAAREAID, MAX(PURCHORDERDATE) PURCHORDERDATE  FROM VENDPURCHORDERJOUR GROUP BY PURCHID,DATAAREAID )VENDPURCHN  ON PURCHLINE.PURCHID = VENDPURCHN.PURCHID");
            sbSql.AppendLine("AND VENDPURCHN.DATAAREAID= PURCHLINE.DATAAREAID");
    
         
            sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(MARKUPCODE) MARKUPCODE,MAX(TRANSRECID) TRANSRECID,MAX(VALUE) VALUE,DATAAREAID FROM MARKUPTRANS WHERE TRANSTABLEID='492' AND MARKUPCODE='FREIGHT' GROUP BY TRANSRECID,DATAAREAID) FREIGHT ON");
            sbSql.AppendLine("VENDINVOICETRANS.RECID=FREIGHT.TRANSRECID AND VENDINVOICETRANS.DATAAREAID=FREIGHT.DATAAREAID");
            sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(MARKUPCODE) MARKUPCODE,MAX(TRANSRECID) TRANSRECID,MAX(VALUE) VALUE,DATAAREAID FROM MARKUPTRANS WHERE TRANSTABLEID='492' AND MARKUPCODE='INSURANCE' GROUP BY TRANSRECID,DATAAREAID) INNSURANCE ON");                           
            sbSql.AppendLine("VENDINVOICETRANS.RECID=INNSURANCE.TRANSRECID AND VENDINVOICETRANS.DATAAREAID=INNSURANCE.DATAAREAID");
            //11-26-2018
            sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(MARKUPCODE) MARKUPCODE,MAX(TRANSRECID) TRANSRECID,MAX(VALUE) VALUE,DATAAREAID FROM MARKUPTRANS WHERE TRANSTABLEID='492' AND MARKUPCODE='DISCOUNT' GROUP BY TRANSRECID,DATAAREAID) DISCOUNT ON");
            sbSql.AppendLine("VENDINVOICETRANS.RECID=DISCOUNT.TRANSRECID AND VENDINVOICETRANS.DATAAREAID=DISCOUNT.DATAAREAID");



            sbSql.AppendLine(" WHERE  INVENTDIM.INVENTSITEID ='" + PaymentGeneralOBJ.Factory + "'");
            sbSql.AppendLine(" AND VENDINVOICETRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", PaymentGeneralOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", PaymentGeneralOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" AND (VENDINVOICEINFOTABLE.VENDINVOICESAVESTATUS=0 OR VENDINVOICEINFOTABLE.VENDINVOICESAVESTATUS IS NULL)");



            if (PaymentGeneralOBJ.GroupVoucher == "Domestic")
            {
                sbSql.AppendLine(" AND VendInvoiceJour.numberSequenceGroup IN ('" + strFac + "-DM','" + strFac + "-CND')");
            }
            else if (PaymentGeneralOBJ.GroupVoucher == "Import")
            {

                sbSql.AppendLine(" AND VendInvoiceJour.numberSequenceGroup IN ('" + strFac + "-IM','" + strFac + "-CNI')");
            }
            else if (PaymentGeneralOBJ.GroupVoucher == "Material")
            {
                sbSql.AppendLine(" AND VendInvoiceJour.numberSequenceGroup IN ('" + strFac + "-MT','" + strFac + "-CNM')");
            }
            else if (PaymentGeneralOBJ.GroupVoucher == "Payment")
            {
                sbSql.AppendLine(" AND VendInvoiceJour.numberSequenceGroup IN ('" + strFac + "-HO','" + strFac + "-HO2','" + strFac + "-CHO','" + strFac + "-CHO2')");
            }

            if (PaymentGeneralOBJ.StartVoucher != "")
            {
                sbSql.AppendLine(" AND VENDINVOICEJOUR.LEDGERVOUCHER between '" + PaymentGeneralOBJ.StartVoucher + "' and '" + PaymentGeneralOBJ.EndVoucher + "'");
            }




            sbSql.AppendLine("GROUP BY VENDINVOICEJOUR.LEDGERVOUCHER,VENDINVOICETRANS.INVOICEID,VENDINVOICETRANS.INVOICEDATE ,VENDINVOICETRANS.PURCHID,VENDINVOICETRANS.INVOICEDATE,VENDINVOICEJOUR.ORDERACCOUNT");
            sbSql.AppendLine(" ,VENDDIRPARTYTABLEVIEW.NAME,VENDINVOICEJOUR.CURRENCYCODE,VENDINVOICEJOUR.EXCHRATE,VENDINVOICETRANS.TAXAMOUNT,VENDTRANSOPEN.ECL_WHTAXAMOUNT");
            sbSql.AppendLine("  ,LASTSETTLEVOUCHER,LASTSETTLEDATE,PURCHTABLE.PURCHSTATUS,VendTrans.Closed ");
            sbSql.AppendLine(" ORDER BY VENDINVOICEJOUR.LEDGERVOUCHER");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }



    }
}
