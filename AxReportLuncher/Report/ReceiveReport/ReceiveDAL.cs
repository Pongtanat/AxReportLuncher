using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NewVersion.Report.ReceiveReport
{
    class ReceiveDAL
    {
        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();


        public ADODB.Recordset getReceiveDetail(ReceiveOBJ ReceiveOBJ)
        {
            StringBuilder sbSql = new StringBuilder();


            String strFac = ReceiveOBJ.strFactory;

            if (ReceiveOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = ReceiveOBJ.strFactory;
            }

            /*

            sbSql.AppendLine("SELECT INVENTDIM.INVENTSITEID [FACTORY] ");
            sbSql.AppendLine(",VendInvoiceJour.InvoiceDate  [Receivedate]");
            sbSql.AppendLine(",VendInvoiceJour.LEDGERVOUCHER ReceivingNo");
            sbSql.AppendLine(",CASE NUMBERSEQUENCEGROUP.ECL_PURCHASETYPE WHEN 0 THEN 'RETURN' WHEN 1 THEN 'DOMESTICS'");
            sbSql.AppendLine("WHEN 2 THEN 'IMPORT'");
            sbSql.AppendLine("WHEN 3 THEN 'MATERIAL'");
            sbSql.AppendLine("END AS RRType");
            sbSql.AppendLine(",VENDINVOICETRANS.OrigPurchId PONo");
            sbSql.AppendLine(",VendInvoiceJour.INVOICEID [InvoiceNO]");
            sbSql.AppendLine(" ,VENDINVOICETRANS.itemID ITEMNO");
            sbSql.AppendLine(",VENDINVOICETRANS.NAME ItemName");
            sbSql.AppendLine(",CASE PURCHTABLE.PURCHASETYPE WHEN 3 THEN 'Receipt'");
            sbSql.AppendLine("WHEN 4 THEN 'Return'");
            sbSql.AppendLine("END AS TranType");
            sbSql.AppendLine(" ,VendInvoiceJour.InvoiceAccount VenderCode");
            sbSql.AppendLine(",VendDirPartyTableView.NAME VenderName");
            sbSql.AppendLine(",VENDINVOICETRANS.PurchUnit Unit");
            sbSql.AppendLine(",VENDINVOICETRANS.CURRENCYCODE CURR");
            sbSql.AppendLine(" ,CASE WHEN ISNULL(MARKUPTRANS.MarkupCode,'')<>'NOCOM' THEN 'Goods' ELSE 'NOCOM' END [COM]");
            sbSql.AppendLine(",VendInvoiceTrans.PurchPrice [VALUE]");
            sbSql.AppendLine(",DIMENSIONFINANCIALTAG.ECL_SHORTNAME Section");
            sbSql.AppendLine(",D3_DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE SubSection");
            sbSql.AppendLine(" ,VENDINVOICETRANS.QTY");
            sbSql.AppendLine(",VendInvoiceTrans.PurchPrice Price");
            sbSql.AppendLine(" ,(VENDINVOICETRANS.QTY*VendInvoiceTrans.PurchPrice) TotalReceipt");
            sbSql.AppendLine(",FREIGHT.VALUE [FREIGHT]");
            sbSql.AppendLine(",INNSURANCE.VALUE [INSURANCE]");
            sbSql.AppendLine(",VENDINVOICETRANS.LINEAMOUNTMST");
            sbSql.AppendLine(",VENDINVOICETRANS.TAXAMOUNT VAT");
            sbSql.AppendLine(",CASE WHEN ISNULL(MARKUPTRANS.MarkupCode,'')<>'NOCOM' THEN 'COM' ELSE 'NOCOM' END [COM]");
           
            sbSql.AppendLine("FROM VendInvoiceJour");
            sbSql.AppendLine("INNER JOIN VENDTABLE on vendtable.ACCOUNTNUM=vendInvoiceJour.InvoiceAccount");
            sbSql.AppendLine("AND vendtable.DATAAREAID=vendInvoiceJour.DATAAREAID");
            sbSql.AppendLine("INNER JOIN VendDirPartyTableView on VENDTABLE.PARTY=VendDirPartyTableView.PARTY");
            sbSql.AppendLine("INNER JOIN vendTrans ON vendTrans.Voucher = vendInvoiceJour.LedgerVoucher");
            sbSql.AppendLine("AND vendTrans.AccountNum = vendInvoiceJour.InvoiceAccount");
            sbSql.AppendLine("AND vendTrans.TransDate = vendInvoiceJour.InvoiceDate");
            sbSql.AppendLine("AND vendTrans.DATAAREAID = vendInvoiceJour.DATAAREAID");
            sbSql.AppendLine("INNER JOIN VENDINVOICETRANS ON VendInvoiceJour.PURCHID=VENDINVOICETRANS.PURCHID");
            sbSql.AppendLine("AND VendInvoiceJour.INVOICEID=VENDINVOICETRANS.INVOICEID");
            sbSql.AppendLine("AND VendInvoiceJour.INVOICEDATE=VENDINVOICETRANS.INVOICEDATE");
            sbSql.AppendLine("AND VendInvoiceJour.NUMBERSEQUENCEGROUP=VENDINVOICETRANS.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VendInvoiceJour.INTERNALINVOICEID=VENDINVOICETRANS.INTERNALINVOICEID");
            sbSql.AppendLine("AND VendInvoiceJour.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET=VENDINVOICETRANS.DefaultDimension");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.RECID=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("INNER JOIN DimensionAttribute ON DimensionAttribute.RECID=DimensionAttributeValue.DIMENSIONATTRIBUTE");
            sbSql.AppendLine("INNER JOIN DIMENSIONFINANCIALTAG ON DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE=DIMENSIONFINANCIALTAG.VALUE");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM D3_DIMENSIONATTRIBUTEVALUESETITEM ON D3_DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET=VENDINVOICETRANS.DefaultDimension");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValue D3_DimensionAttributeValue ON D3_DimensionAttributeValue.RECID=D3_DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("INNER JOIN DimensionAttribute D3_DimensionAttribute ON D3_DimensionAttribute.RECID=D3_DimensionAttributeValue.DIMENSIONATTRIBUTE");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine("INNER JOIN PURCHTABLE ON PURCHTABLE.PURCHID=VENDINVOICETRANS.ORIGPURCHID");
            sbSql.AppendLine("AND PURCHTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN PURCHLINE ON PURCHTABLE.PURCHID=PURCHLINE.PURCHID");
            sbSql.AppendLine("AND PURCHTABLE.DATAAREAID=PURCHLINE.DATAAREAID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVENTTRANSID=PURCHLINE.INVENTTRANSID");
            sbSql.AppendLine("INNER JOIN NUMBERSEQUENCEGROUP ON PURCHTABLE.NUMBERSEQUENCEGROUP=NUMBERSEQUENCEGROUP.NUMBERSEQUENCEGROUPID");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            sbSql.AppendLine("");
            
            
            

            sbSql.AppendLine(" WHERE REQFactory='" + ReceiveOBJ.Factory + "'");
            //Confirm
            sbSql.AppendLine("AND REQ_STATUS = '2' AND  EmpOrgName != ''");

            sbSql.AppendLine(" AND CONVERT(date,HOYA_IRTable.REQADMRECEIPTDATE,103) BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", RequistionOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", ReceiveOBJ.DateTo) + "',103)");
            //sbSql.AppendLine("  GROUP BY EmpOrgName,inventtable.ITEMID,REQUSER,ECORESPRODUCTTRANSLATIONS.PRODUCTNAME WITH ROLLUP");
            //sbSql.AppendLine(" HAVING  NOT ECORESPRODUCTTRANSLATIONS.PRODUCTNAME IS NULL OR inventtable.ITEMID IS NULL");

            sbSql.AppendLine("  GROUP BY EmpOrgName,inventtable.ITEMID,REQUSER,ECORESPRODUCTTRANSLATIONS.PRODUCTNAME, HOYA_IRTABLE.REQREQUISITIONTYPE WITH ROLLUP");
            sbSql.AppendLine(" HAVING  NOT HOYA_IRTABLE.REQREQUISITIONTYPE IS NULL OR inventtable.ITEMID IS NULL");


            */

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;
              

        }
    }
}
