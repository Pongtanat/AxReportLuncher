using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.Report.SalesReturn.SummaryTransaction
{
    class ReturnTransactionDAL
    {
        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();

        public DataTable getCategoryByType()
        {
            StringBuilder sbSql = new StringBuilder();

        sbSql.Remove(0, sbSql.Length);
        sbSql.AppendLine(" SELECT DISTINCT ITEMGROUPID AS ITEMGROUPID");
        sbSql.AppendLine(" FROM INVENTITEMGROUPITEM");
        sbSql.AppendLine(" WHERE NOT(ITEMGROUPID LIKE 'S%' AND LEN(ITEMGROUPID)=2)");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public DataTable getAllSectionByFactory(string strFactory)
        {
            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine(" SELECT SSec.ECL_SHORTNAME Section FROM DIMENSIONATTRIBUTE");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEDIRCATEGORY ON DIMENSIONATTRIBUTE.RECID=DIMENSIONATTRIBUTEDIRCATEGORY.DIMENSIONATTRIBUTE");
        sbSql.AppendLine(" INNER JOIN DIMENSIONFINANCIALTAG Fac ON DIMENSIONATTRIBUTEDIRCATEGORY.DIRCATEGORY=Fac.FINANCIALTAGCATEGORY");
        sbSql.AppendLine(" INNER JOIN DIMENSIONFINANCIALTAG SSec ON Fac.ECL_COMCODE=SSec.ECL_COMCODE");
        sbSql.AppendLine(" WHERE NAME='D1_Factory' AND NOT(SSec.VALUE LIKE Fac.ECL_COMCODE + '%') ");
        sbSql.AppendLine(" AND SSec.VALUE <> Fac.VALUE");

        sbSql.AppendLine("AND Fac.ECL_SHORTNAME = '" + strFactory  + "'");
        sbSql.AppendLine(" GROUP BY SSec.ECL_SHORTNAME");
        sbSql.AppendLine(" ORDER BY SSec.ECL_SHORTNAME");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public DataTable getAllSubSectionByFactory(string strFactory)
        {
            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine(" SELECT SSec.VALUE SubSection FROM DIMENSIONATTRIBUTE");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEDIRCATEGORY ON DIMENSIONATTRIBUTE.RECID=DIMENSIONATTRIBUTEDIRCATEGORY.DIMENSIONATTRIBUTE");
        sbSql.AppendLine(" INNER JOIN DIMENSIONFINANCIALTAG Fac ON DIMENSIONATTRIBUTEDIRCATEGORY.DIRCATEGORY=Fac.FINANCIALTAGCATEGORY");
        sbSql.AppendLine(" INNER JOIN DIMENSIONFINANCIALTAG SSec ON Fac.ECL_COMCODE=SSec.ECL_COMCODE");
        sbSql.AppendLine(" WHERE NAME='D1_Factory' AND NOT(SSec.VALUE LIKE Fac.ECL_COMCODE + '%') ");
        sbSql.AppendLine(" AND SSec.VALUE <> Fac.VALUE");
     
        sbSql.AppendLine("AND Fac.ECL_SHORTNAME = '"+strFactory+"'");
        sbSql.AppendLine(" GROUP BY SSec.VALUE");
        sbSql.AppendLine(" ORDER BY SSec.VALUE");
        DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
         return dt;
        }

        public ADODB.Recordset getSummary2(ReturnTransactionOBJ sumObj, string strCategory, string locationID)
        {
            StringBuilder sbSql = new StringBuilder();

            if (sumObj.TransType == 3)
            {

           
            sbSql.AppendLine("SELECT  ");
            sbSql.AppendLine("DIMENSIONFINANCIALTAG.[VALUE] as SubSection");
            sbSql.AppendLine(",CustInvoiceJour.LedgerVoucher as Voucher  ");
            sbSql.AppendLine(",CustInvoiceJour.INVOICEID  as Invoice");
            sbSql.AppendLine(",CustInvoiceJour.INVOICEDATE as InvDate");
            sbSql.AppendLine(",SALESLINE.ITEMID [Item No]");
            sbSql.AppendLine(",SalesLine.NAME [Item Name] ");
            sbSql.AppendLine(",INVENTTABLE.ITEMID  [Item Sales]");
            sbSql.AppendLine(",SALESLINE.SALESUNIT [Unit]");
            sbSql.AppendLine(",SalesLine.SalesQty*-1 [Qty]");
            sbSql.AppendLine(",(SalesLine.LINEAMOUNT*(SALESTABLE.FIXEDEXCHRATE/100)) *-1 [Cost]");
            sbSql.AppendLine(",'' [Cost/PCS]");

            sbSql.AppendLine("FROM  SALESLINE ");
            sbSql.AppendLine("INNER JOIN SALESTABLE ON SALESTABLE.SALESID = SALESLINE.SALESID");
          
            sbSql.AppendLine("INNER JOIN INVENTTABLE ON SALESLINE.ITEMID = INVENTTABLE.ITEMID ");
            sbSql.AppendLine("AND SALESLINE.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine("INNER  JOIN CUSTINVOICEJOUR ON SALESLINE.SALESID = CUSTINVOICEJOUR.SALESID");
           
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTITEMGROUPITEM.ITEMID = SALESLINE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMGROUPITEM.ITEMDATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET=SalesTable.DefaultDimension");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.RECID=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("INNER JOIN DimensionFinancialTag ON DimensionAttributeValue.ENTITYINSTANCE=DimensionFinancialTag.recid");
            sbSql.AppendLine("INNER JOIN DimensionAttribute ON DimensionAttributeValue.DIMENSIONATTRIBUTE=DimensionAttribute.recid");

            // 11/9/2018
           // sbSql.AppendLine("LEFT OUTER JOIN (SELECT inventdimid,INVENTLOCATIONID,DATAAREAID from INVENTDIM GROUP BY INVENTDIMID,INVENTLOCATIONID,DATAAREAID ) as tb_INVENTDIM");
           // sbSql.AppendLine("ON SALESLINE.INVENTDIMID=tb_INVENTDIM.INVENTDIMID AND SALESLINE.DATAAREAID=tb_INVENTDIM.DATAAREAID");




               sbSql.AppendLine(" WHERE  SALESTABLE.INVENTSITEID='" + sumObj.Factory + "'");
                sbSql.AppendLine("AND DIMENSIONATTRIBUTE.NAME='D3_Subsection'");
              //  sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY='" + sumObj.TransType + "'/*PURCHASE ORDER(Received)=3, TRANSACTION(Issued,Shipment)=4*/");
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + strCategory + "'");
              //  sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + sumObj.Section.Replace(",", "','") + "')");


                if (sumObj.Factory == "GMO")
                {
                    sbSql.AppendLine("AND  SALESTABLE.NUMBERSEQUENCEGROUP != 'MO-RNOC'");
                }
                else
                {
                    sbSql.AppendLine("AND  SALESTABLE.NUMBERSEQUENCEGROUP != '" + sumObj.Factory + "-RNOC'");
                }

                sbSql.AppendLine("  AND DIMENSIONFINANCIALTAG.VALUE IN ('" + sumObj.Section.Replace(",", "','") + "')");




                if (sumObj.ItemFrom != "")
                {
                    if (sumObj.ItemTo != "")
                    {
                        sbSql.AppendLine(" AND INVENTTABLE.ITEMID  BETWEEN '" + sumObj.ItemFrom + "' AND '" + sumObj.ItemTo + "'");
                    }
                    else
                    {

                        sbSql.AppendLine(" AND INVENTTABLE.ITEMID  LIKE '" + sumObj.ItemFrom.Replace("*", "%") + "'");
                    }
                }

                if (sumObj.VoucherFrom != "")
                {
                    if (sumObj.VoucherTo != "")
                    {
                        sbSql.AppendLine(" AND CustInvoiceJour.LedgerVoucher BETWEEN '" + sumObj.VoucherFrom + "' AND '" + sumObj.VoucherTo + "'");
                    }
                    else
                    {
                        sbSql.AppendLine(" AND CustInvoiceJour.LedgerVoucher LIKE '" + sumObj.VoucherFrom.Replace("*", "%") + "'");

                    }
                }//


                if (locationID != "")
                {
                    if (locationID == "F1")
                    {
                       // sbSql.AppendLine(" AND tb_INVENTDIM.INVENTLOCATIONID = '" + locationID + "'");
                        sbSql.AppendLine("AND SALESLINE.INVENTDIMID ='DIM1800001'");
                    }
                    else
                    {
                        sbSql.AppendLine("AND SALESLINE.INVENTDIMID ='DIM1800002'");
                    }

                }

                sbSql.AppendLine(" AND CustInvoiceJour.INVOICEDATE BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateFrom) + "',103) ");
                sbSql.AppendLine("      AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateTo) + "',103)");
          
            }
            else if (sumObj.TransType == 4)
            {

             sbSql.AppendLine(" SELECT * FROM (");
            sbSql.AppendLine(" SELECT CASE WHEN INVENTTRANS.VOUCHER IS NULL THEN NULL ELSE DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE END DISPLAYVALUE");
            sbSql.AppendLine(" ,INVENTTRANS.VOUCHER,INVENTJOURNALTABLE.DESCRIPTION REFERENCE,INVENTTRANS.DATEFINANCIAL,INVENTTRANS.ITEMID");
            sbSql.AppendLine(" ,CASE WHEN DISPLAYVALUE IS NULL THEN 'GRAND TOTAL' ELSE CASE WHEN INVENTTRANS.VOUCHER IS NULL THEN 'TOTAL' ELSE ECORESPRODUCTTRANSLATION.NAME END END NAME");
            sbSql.AppendLine(" ,INVENTTABLEMODULE.UNITID");
            sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) QTY, SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
            sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
            sbSql.AppendLine(" FROM INVENTTRANS ");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("	    AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine("	    AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
         

            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
          
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'") ;
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");

            sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + sumObj.Factory + "'");
            sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY=4");
           sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + strCategory + "'");
           sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + sumObj.Section.Replace(",", "','") + "')");



            if (locationID != "")
            {
                sbSql.AppendLine(" AND INVENTDIM.INVENTLOCATIONID = '" + locationID + "'");

            }

            if (sumObj.ItemFrom != "")
            {
                if (sumObj.ItemTo != "")
                {
                    sbSql.AppendLine(" AND INVENTTRANS.ITEMID  BETWEEN '" + sumObj.ItemFrom + "' AND '" + sumObj.ItemTo + "'");
                }
                else
                {

                    sbSql.AppendLine(" AND INVENTTRANS.ITEMID  LIKE '" + sumObj.ItemFrom.Replace("*", "%") + "'");
                }
            }

            if (sumObj.VoucherFrom != "")
            {
                if (sumObj.VoucherTo != "")
                {
                    sbSql.AppendLine(" AND INVENTTRANS.LedgerVoucher BETWEEN '" + sumObj.VoucherFrom + "' AND '" + sumObj.VoucherTo + "'");
                }
                else
                {
                    sbSql.AppendLine(" AND INVENTTRANS.LedgerVoucher LIKE '" + sumObj.VoucherFrom.Replace("*", "%") + "'");

                }
            }//


            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateTo) + "',103)");

            sbSql.AppendLine(" GROUP BY DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE,INVENTTRANS.VOUCHER,INVENTJOURNALTABLE.DESCRIPTION,INVENTTRANS.DATEFINANCIAL");
            sbSql.AppendLine(" ,INVENTTRANS.ITEMID,INVENTTABLEMODULE.UNITID,ECORESPRODUCTTRANSLATION.NAME WITH ROLLUP");

            sbSql.AppendLine(" ) SUMMARY WHERE NOT(NAME IS NULL)");

            }

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();
            object ret = null;
            rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

            return rs;
        } /// end function
          /// 
        public ADODB.Recordset getSummaryByVoucher(ReturnTransactionOBJ sumObj, string strCategory, string locationID)
        {
            StringBuilder sbSql = new StringBuilder();

            if (sumObj.TransType == 3)
            {
            sbSql.AppendLine("SELECT  ");
            sbSql.AppendLine("DIMENSIONFINANCIALTAG.[VALUE] as SubSection");
            sbSql.AppendLine(",CustInvoiceJour.LedgerVoucher as Voucher  ");
            sbSql.AppendLine(",CustInvoiceJour.INVOICEID  as Invoice");
            sbSql.AppendLine(",CustInvoiceJour.INVOICEDATE as InvDate");
            sbSql.AppendLine(",SALESLINE.ITEMID [Item No]");
            sbSql.AppendLine(",SalesLine.NAME [Item Name] ");
            sbSql.AppendLine(",INVENTTABLE.ITEMID  [Item Sales]");
            sbSql.AppendLine(",SALESLINE.SALESUNIT [Unit]");
            sbSql.AppendLine(",SalesLine.SalesQty*-1 [Qty]");
            sbSql.AppendLine(",(SalesLine.LINEAMOUNT*(SALESTABLE.FIXEDEXCHRATE/100)) *-1 [Cost]");
            sbSql.AppendLine(",''[Cost/PCS]");

            sbSql.AppendLine("FROM  SALESLINE ");
            sbSql.AppendLine("INNER JOIN SALESTABLE ON SALESTABLE.SALESID = SALESLINE.SALESID");
            sbSql.AppendLine("INNER JOIN INVENTTABLE ON SALESLINE.ITEMID = INVENTTABLE.ITEMID ");
            sbSql.AppendLine("AND SALESLINE.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine("INNER  JOIN CUSTINVOICEJOUR ON SALESLINE.SALESID = CUSTINVOICEJOUR.SALESID");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTITEMGROUPITEM.ITEMID = SALESLINE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMGROUPITEM.ITEMDATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET=SalesTable.DefaultDimension");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.RECID=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("INNER JOIN DimensionFinancialTag ON DimensionAttributeValue.ENTITYINSTANCE=DimensionFinancialTag.recid");
            sbSql.AppendLine("INNER JOIN DimensionAttribute ON DimensionAttributeValue.DIMENSIONATTRIBUTE=DimensionAttribute.recid");


            // 11/9/2018
            //sbSql.AppendLine("LEFT OUTER JOIN (SELECT inventdimid,INVENTLOCATIONID,DATAAREAID from INVENTDIM GROUP BY INVENTDIMID,INVENTLOCATIONID,DATAAREAID ) as tb_INVENTDIM");
           // sbSql.AppendLine("ON SALESLINE.INVENTDIMID=tb_INVENTDIM.INVENTDIMID AND SALESLINE.DATAAREAID=tb_INVENTDIM.DATAAREAID");


            sbSql.AppendLine(" WHERE SALESTABLE.INVENTSITEID='" + sumObj.Factory + "'");
            sbSql.AppendLine("AND DIMENSIONATTRIBUTE.NAME='D3_Subsection'");
            sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + strCategory + "'");

                if (sumObj.Factory == "GMO")
                {
                    sbSql.AppendLine("AND  SALESTABLE.NUMBERSEQUENCEGROUP != 'MO-RNOC'");

                }
                else
                {
                    sbSql.AppendLine("AND  SALESTABLE.NUMBERSEQUENCEGROUP != '" + sumObj.Factory + "-RNOC'");
                }

                sbSql.AppendLine("  AND DIMENSIONFINANCIALTAG.VALUE IN ('" + sumObj.Section.Replace(",", "','") + "')");

                //sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + sumObj.Section.Replace(",", "','") + "')");


                if (sumObj.ItemFrom != "")
                {
                    if (sumObj.ItemTo != "")
                    {
                        sbSql.AppendLine(" AND INVENTTABLE.ITEMID  BETWEEN '" + sumObj.ItemFrom + "' AND '" + sumObj.ItemTo + "'");
                    }
                    else
                    {

                        sbSql.AppendLine(" AND INVENTTABLE.ITEMID  LIKE '" + sumObj.ItemFrom.Replace("*", "%") + "'");
                    }
                }

                if (sumObj.VoucherFrom != "")
                {
                    if (sumObj.VoucherTo != "")
                    {
                        sbSql.AppendLine(" AND CustInvoiceJour.LedgerVoucher BETWEEN '" + sumObj.VoucherFrom + "' AND '" + sumObj.VoucherTo + "'");
                    }
                    else
                    {
                        sbSql.AppendLine(" AND CustInvoiceJour.LedgerVoucher LIKE '" + sumObj.VoucherFrom.Replace("*", "%") + "'");

                    }
                }//

                if (locationID != "")
                {
                    if (locationID == "F1")
                    {
                        // sbSql.AppendLine(" AND tb_INVENTDIM.INVENTLOCATIONID = '" + locationID + "'");
                        sbSql.AppendLine("AND SALESLINE.INVENTDIMID ='DIM1800001'");
                    }
                    else
                    {
                        sbSql.AppendLine("AND SALESLINE.INVENTDIMID ='DIM1800002'");
                    }

                }

                sbSql.AppendLine(" AND CustInvoiceJour.INVOICEDATE BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateFrom) + "',103) ");
                sbSql.AppendLine("      AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateTo) + "',103)");

                  //sbSql.AppendLine(" GROUP BY VENDINVOICEJOUR.LEDGERVOUCHER,DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE,VENDINVOICEJOUR.INVOICEDATE");
           // sbSql.AppendLine(" ,INVENTTRANS.ITEMID,ECORESPRODUCTTRANSLATION.NAME,INVENTTABLEMODULE.UNITID WITH ROLLUP");

          //  sbSql.AppendLine(" HAVING DISPLAYVALUE IS NULL OR NOT(UNITID IS NULL)");

            }
            else if (sumObj.TransType == 4)
            {
            sbSql.AppendLine(" SELECT CASE WHEN DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IS NULL THEN '' ELSE INVENTTRANS.VOUCHER END VOUCHER");
            sbSql.AppendLine(" ,DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE");
            sbSql.AppendLine(" ,INVENTJOURNALTABLE.DESCRIPTION REFERENCE,INVENTTRANS.DATEFINANCIAL,INVENTTRANS.ITEMID");
            sbSql.AppendLine(" ,ECORESPRODUCTTRANSLATION.NAME");
            sbSql.AppendLine(" ,CASE WHEN INVENTTRANS.VOUCHER IS NULL THEN 'GRAND TOTAL' ELSE");
            sbSql.AppendLine("  CASE WHEN DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IS NULL THEN 'TOTAL' ELSE");
            sbSql.AppendLine("  INVENTTABLEMODULE.UNITID END END 'UNITID'");
            sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) QTY, SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
            sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
            sbSql.AppendLine(" FROM INVENTTRANS ");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("	    AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine("	    AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND                 INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
 
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");

                sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + sumObj.Factory + "'");
                sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY=4");
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + strCategory + "'");
                sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + sumObj.Section.Replace(",", "','") + "')");



                if (sumObj.ItemFrom != "")
                {
                    if (sumObj.ItemTo != "")
                    {
                        sbSql.AppendLine(" AND INVENTTABLE.ITEMID  BETWEEN '" + sumObj.ItemFrom + "' AND '" + sumObj.ItemTo + "'");
                    }
                    else
                    {

                        sbSql.AppendLine(" AND INVENTTABLE.ITEMID  LIKE '" + sumObj.ItemFrom.Replace("*", "%") + "'");
                    }
                }///

                if (sumObj.VoucherFrom != "")
                {
                    if (sumObj.VoucherTo != "")
                    {
                        sbSql.AppendLine(" AND CustInvoiceJour.LedgerVoucher BETWEEN '" + sumObj.VoucherFrom + "' AND '" + sumObj.VoucherTo + "'");
                    }
                    else
                    {
                        sbSql.AppendLine(" AND CustInvoiceJour.LedgerVoucher LIKE '" + sumObj.VoucherFrom.Replace("*", "%") + "'");

                    }
                }//

              if (locationID != "")
            {
                sbSql.AppendLine(" AND INVENTDIM.INVENTLOCATIONID = '" + locationID + "'");

            }


              sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateFrom) + "',103) ");
                sbSql.AppendLine("      AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateTo) + "',103)");
               
                sbSql.AppendLine(" GROUP BY INVENTTRANS.VOUCHER,DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE,INVENTJOURNALTABLE.DESCRIPTION,INVENTTRANS.DATEFINANCIAL");
            sbSql.AppendLine(" ,INVENTTRANS.ITEMID,INVENTTABLEMODULE.UNITID,ECORESPRODUCTTRANSLATION.NAME WITH ROLLUP");


            sbSql.AppendLine(" UNION ALL ");
            sbSql.AppendLine(" SELECT CASE WHEN DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IS NULL THEN '' ELSE VENDINVOICEJOUR.LEDGERVOUCHER END VOUCHER");
            sbSql.AppendLine(" ,DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE");
            sbSql.AppendLine(" ,'' REFERENCE,VENDINVOICEJOUR.INVOICEDATE,INVENTTRANS.ITEMID");
            sbSql.AppendLine(" ,ECORESPRODUCTTRANSLATION.NAME");
            sbSql.AppendLine(" ,CASE WHEN VENDINVOICEJOUR.LEDGERVOUCHER IS NULL THEN 'GRAND TOTAL' ELSE");
            sbSql.AppendLine("  CASE WHEN DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IS NULL THEN 'TOTAL' ELSE");
            sbSql.AppendLine("  INVENTTABLEMODULE.UNITID END END 'UNITID' ");
            sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) QTY, SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
            sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
            sbSql.AppendLine(" FROM INVENTTRANS ");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("	    AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine("	    AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN (SELECT ITEMID, PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID");
            sbSql.AppendLine("              FROM VENDINVOICETRANS");
            sbSql.AppendLine("		        GROUP BY ITEMID,PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID) VENDINVOICETRANS ");
            sbSql.AppendLine("		ON INVENTTRANS.ITEMID=VENDINVOICETRANS.ITEMID AND INVENTTRANS.INVOICEID=VENDINVOICETRANS.INVOICEID");
            sbSql.AppendLine("		AND INVENTTRANS.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine(" ");
            sbSql.AppendLine(" INNER JOIN VENDINVOICEJOUR ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("      AND VENDINVOICETRANS.INVOICEDATE=VENDINVOICEJOUR.INVOICEDATE AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP=VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("      AND VENDINVOICETRANS.INTERNALINVOICEID=VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("      AND VENDINVOICETRANS.DATAAREAID=VENDINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN ECORESPRODUCT ON INVENTTABLE.PRODUCT=ECORESPRODUCT.RECID");
            sbSql.AppendLine(" INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine(" INNER JOIN ECORESPRODUCTCATEGORY ON INVENTTABLE.PRODUCT=ECORESPRODUCTCATEGORY.PRODUCT");
            sbSql.AppendLine(" INNER JOIN EcoResCategory ON ECORESPRODUCTCATEGORY.CATEGORY =EcoResCategory.RECID");
            sbSql.AppendLine(" INNER JOIN EcoResCategoryTranslation ON EcoResCategory.RECID =EcoResCategoryTranslation.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");

            sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + sumObj.Factory + "'");
            sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY=3");

            sbSql.AppendLine(" AND EcoResCategoryTranslation.SearchText='" + strCategory + "'");
            sbSql.AppendLine(" AND ECORESPRODUCT.PRODUCTTYPE=2");
            sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + sumObj.Section.Replace(",", "','") + "')");


            if (sumObj.ItemFrom != "")
            {
                if (sumObj.ItemTo != "")
                {
                    sbSql.AppendLine(" AND INVENTTRANS.ITEMID  BETWEEN '" + sumObj.ItemFrom + "' AND '" + sumObj.ItemTo + "'");
                }
                else
                {

                    sbSql.AppendLine(" AND INVENTTRANS.ITEMID  LIKE '" + sumObj.ItemFrom.Replace("*", "%") + "'");
                }
            }

            if (sumObj.VoucherFrom != "")
            {
                if (sumObj.VoucherTo != "")
                {
                    sbSql.AppendLine(" AND VENDINVOICEJOUR.LEDGERVOUCHER BETWEEN '" + sumObj.VoucherFrom + "' AND '" + sumObj.VoucherTo + "'");
                }
                else
                {
                    sbSql.AppendLine(" AND VENDINVOICEJOUR.LEDGERVOUCHER LIKE '" + sumObj.VoucherFrom.Replace("*", "%") + "'");

                }
            }//

              if (locationID != "")
            {
                sbSql.AppendLine(" AND INVENTDIM.INVENTLOCATIONID = '" + locationID + "'");

            }

            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE BETWEEN ");
            sbSql.AppendLine(" AND SALESTABLE.DATEFINANCIAL BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", sumObj.DateTo) + "',103)");
            sbSql.AppendLine(" GROUP BY VENDINVOICEJOUR.LEDGERVOUCHER,DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE,VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine(" ,INVENTTRANS.ITEMID,INVENTTABLEMODULE.UNITID,ECORESPRODUCTTRANSLATION.NAME WITH ROLLUP");

            sbSql.AppendLine(" ) SUMMARY WHERE DISPLAYVALUE IS NULL OR NOT(UNITID IS NULL)");
            }

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();
            object ret = null;
            rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

            return rs;
        }


    }//end class
}
