using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.Report.MaterialReport
{
    class MaterialDAL
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
            sbSql.AppendLine(" SELECT NUMBERSEQUENCEGROUP2 [Customer],THB,PCS");
            sbSql.AppendLine(" ORDER BY CustGroup");


            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public ADODB.Recordset getMaterialReceive(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            
        sbSql.AppendLine("SELECT * FROM (");
        sbSql.AppendLine("SELECT CASE WHEN LEDGERVOUCHER IS NULL THEN NULL ELSE DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE END SubSection");
        sbSql.AppendLine(",ECL_SUBGROUP,VENDINVOICEJOUR.LEDGERVOUCHER,VENDINVOICEJOUR.INVOICEDATE,INVENTTRANS.ITEMID");

        if (MaterialOBJ.Factory == "RP")
        {
            sbSql.AppendLine(",CASE WHEN LEDGERVOUCHER IS NULL  THEN ECL_SUBGROUP+'-TOTAL' ELSE ECORESPRODUCTTRANSLATION.NAME  END NAME");
        }
        else
        {

            sbSql.AppendLine(",CASE WHEN LEDGERVOUCHER IS NULL  THEN DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE+'-TOTAL' ELSE ECORESPRODUCTTRANSLATION.NAME  END NAME");
        }

        sbSql.AppendLine(",INVENTTABLE.HOYA_PRODUCTIONITEM");
        sbSql.AppendLine(" ,HOYA_GLASSTYPE");
        sbSql.AppendLine(" ,HOYA_SOZAIDIV");
        sbSql.AppendLine(" ,INVENTTABLEMODULE.UNITID");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) QTY, SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
        sbSql.AppendLine(",CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
        sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine("	    AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine("	    AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
        sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID");
        sbSql.AppendLine("	    AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN (SELECT ITEMID, PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID ");
        sbSql.AppendLine("              FROM VENDINVOICETRANS");
        sbSql.AppendLine("		        GROUP BY ITEMID,PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID) VENDINVOICETRANS ");
        sbSql.AppendLine("		ON INVENTTRANS.ITEMID=VENDINVOICETRANS.ITEMID AND INVENTTRANS.INVOICEID=VENDINVOICETRANS.INVOICEID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN VENDINVOICEJOUR ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
        sbSql.AppendLine("      AND VENDINVOICETRANS.INVOICEDATE=VENDINVOICEJOUR.INVOICEDATE AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP=VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
        sbSql.AppendLine("      AND VENDINVOICETRANS.INTERNALINVOICEID=VENDINVOICEJOUR.INTERNALINVOICEID");
        sbSql.AppendLine("      AND VENDINVOICETRANS.DATAAREAID=VENDINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
        sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");

        sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'--Physical=0, Financial=1");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'--Invent=0, Purch=1, Sales=2");
        sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
        sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY=3");

        if(MaterialOBJ.Category=="All"){
             sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");


           // sbSql.AppendLine("AND (DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('IPCMO','OPCMO','WPCMO','Z1ST','Z1PC')");
         //   sbSql.AppendLine("OR INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y'))");
        }else{
             sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
        }

         foreach(DataRow  dr in getAllSubSectionByFactory(MaterialOBJ.Factory).Rows){
                MaterialOBJ.Section += dr["SubSection"]+",";
            }

         sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + MaterialOBJ.Section.Replace(",", "','") + "')");
        
            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE   BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}",MaterialOBJ.DateTo) + "',103)");

        sbSql.AppendLine(" GROUP BY DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE,ECL_SUBGROUP,VENDINVOICEJOUR.LEDGERVOUCHER,VENDINVOICEJOUR.INVOICEDATE");
        sbSql.AppendLine(" ,INVENTTRANS.ITEMID,INVENTTABLE.HOYA_PRODUCTIONITEM,HOYA_GLASSTYPE,HOYA_SOZAIDIV,INVENTTABLEMODULE.UNITID,ECORESPRODUCTTRANSLATION.NAME WITH ROLLUP");
        sbSql.AppendLine(") SUMMARY WHERE NOT(NAME IS NULL OR ECL_SUBGROUP IS NULL)");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterialShipment(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine("SELECT * FROM (");
        sbSql.AppendLine(" SELECT CASE WHEN INVENTTRANS.VOUCHER IS NULL THEN NULL ELSE DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE END DISPLAYVALUE");
        sbSql.AppendLine(" ,ECL_SUBGROUP");

        sbSql.AppendLine(",CASE when INVENTJOURNALTABLE.HOYA_DIFTYPE = 0 THEN 'SALE'");
        sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 1 THEN 'USED'");
        sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 2 THEN 'NG'");
        sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 3 THEN 'RETURN'");
        sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 4 THEN 'DEAD'");
        sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 5 THEN 'GD'");
        sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 6 THEN 'USED RT'");
        sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 7 THEN 'NG RT'");
        sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 9 THEN 'RD'");
        sbSql.AppendLine(" END [TYPE]");


        sbSql.AppendLine(" ,INVENTTRANS.VOUCHER,INVENTJOURNALTABLE.DESCRIPTION REFERENCE,INVENTTRANS.DATEFINANCIAL,INVENTTRANS.ITEMID");

        if (MaterialOBJ.Factory == "RP")
        {
            sbSql.AppendLine(",CASE WHEN INVENTTRANS.VOUCHER IS NULL  THEN CASE WHEN  DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE = 'Z1PR' THEN  'SALE-TOTAL' ELSE");
            sbSql.AppendLine("CASE WHEN  INVENTJOURNALTABLE.HOYA_DIFTYPE =9 THEN ");
            sbSql.AppendLine("'RD-TOTAL' ELSE  ECL_SUBGROUP+'-TOTAL'  END END ELSE ECORESPRODUCTTRANSLATION.NAME  END NAME");
       
     
        }
        else
        {

            sbSql.AppendLine(",CASE WHEN INVENTTRANS.VOUCHER IS NULL  THEN");

            sbSql.AppendLine("CASE when INVENTJOURNALTABLE.HOYA_DIFTYPE = 0 THEN 'SALE'");
            sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 1 THEN 'USED'");
            sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 2 THEN 'NG'");
            sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 3 THEN 'RETURN'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 4 THEN 'DEAD'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 5 THEN 'GD'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 6 THEN 'USED RT'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 7 THEN 'NG RT'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 9 THEN 'RD'");
            sbSql.AppendLine(" END ");
            
            
            sbSql.AppendLine("+'-TOTAL'  ELSE ECORESPRODUCTTRANSLATION.NAME  END NAME");
     
        }
            
         sbSql.AppendLine(" ,INVENTTABLE.HOYA_PRODUCTIONITEM");
        sbSql.AppendLine(" ,HOYA_GLASSTYPE");
        sbSql.AppendLine(" ,HOYA_SOZAIDIV");
        sbSql.AppendLine(" ,INVENTTABLEMODULE.UNITID");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) QTY, SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
        sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
        sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
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
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");


            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }

            sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
            foreach (DataRow dr in getAllSubSectionByFactory(MaterialOBJ.Factory).Rows)
            {
                MaterialOBJ.Section += dr["SubSection"] + ",";
            }
            sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + MaterialOBJ.Section.Replace(",", "','") + "')");

            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

            sbSql.AppendLine(" GROUP BY DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE,ECL_SUBGROUP,INVENTJOURNALTABLE.HOYA_DIFTYPE,INVENTTRANS.VOUCHER,INVENTJOURNALTABLE.DESCRIPTION,INVENTTRANS.DATEFINANCIAL");
            sbSql.AppendLine(",INVENTTRANS.ITEMID,INVENTTABLEMODULE.UNITID,INVENTTABLE.HOYA_PRODUCTIONITEM,HOYA_GLASSTYPE,HOYA_SOZAIDIV,ECORESPRODUCTTRANSLATION.NAME WITH ROLLUP");
            sbSql.AppendLine(") SUMMARY WHERE NOT(NAME IS NULL OR TYPE IS NULL)");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterialShipmentMO(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            sbSql.AppendLine("SELECT * FROM (");
            sbSql.AppendLine(" SELECT  DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE  DISPLAYVALUE");
            sbSql.AppendLine(",CASE when INVENTJOURNALTABLE.HOYA_DIFTYPE = 0 THEN 'SALE'");
            sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 1 THEN 'USED'");
            sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 2 THEN 'NG'");
            sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 3 THEN 'RETURN'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 4 THEN 'DEAD'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 5 THEN 'GD'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 6 THEN 'USED RT'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 7 THEN 'NG RT'");
            sbSql.AppendLine(" END [TYPE]");
            sbSql.AppendLine(",INVENTTRANS.ITEMID");

            sbSql.AppendLine(",CASE WHEN INVENTTRANS.ITEMID IS NULL  THEN");
            sbSql.AppendLine("CASE when INVENTJOURNALTABLE.HOYA_DIFTYPE = 0 THEN 'SALE'");
            sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 1 THEN 'USED'");
            sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 2 THEN 'NG'");
            sbSql.AppendLine("when INVENTJOURNALTABLE.HOYA_DIFTYPE = 3 THEN 'RETURN'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 4 THEN 'DEAD'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 5 THEN 'GD'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 6 THEN 'USED RT'");
            sbSql.AppendLine(" when INVENTJOURNALTABLE.HOYA_DIFTYPE = 7 THEN 'NG RT'");
            sbSql.AppendLine(" END ");
            sbSql.AppendLine("+'-TOTAL'  ELSE ECORESPRODUCTTRANSLATION.NAME  END NAME");
            sbSql.AppendLine(" ,INVENTTABLE.HOYA_PRODUCTIONITEM");
            sbSql.AppendLine(" ,INVENTTABLEMODULE.UNITID");
            sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) QTY, SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
            sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
           
            sbSql.AppendLine(" FROM INVENTTRANS ");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
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
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");


            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }

            sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
            foreach (DataRow dr in getAllSubSectionByFactory(MaterialOBJ.Factory).Rows)
            {
                MaterialOBJ.Section += dr["SubSection"] + ",";
            }
            sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + MaterialOBJ.Section.Replace(",", "','") + "')");

            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

            sbSql.AppendLine(" GROUP BY DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE,INVENTJOURNALTABLE.HOYA_DIFTYPE");
            sbSql.AppendLine(",INVENTTRANS.ITEMID,INVENTTABLEMODULE.UNITID,INVENTTABLE.HOYA_PRODUCTIONITEM,ECORESPRODUCTTRANSLATION.NAME WITH ROLLUP");
            sbSql.AppendLine(") SUMMARY WHERE NOT(NAME IS NULL OR TYPE IS NULL)");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

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
        sbSql.AppendLine("AND Fac.ECL_SHORTNAME = '" +strFactory+"'");
        sbSql.AppendLine(" GROUP BY SSec.VALUE");

         DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
         return dt;

        }

        public ADODB.Recordset getDetailMaterialReport(MaterialOBJ MaterialOBJ,String SubGroup,string hoya_diftype)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

              sbSql.AppendLine("SELECT");

              if (hoya_diftype == "GD")
              {
                  sbSql.AppendLine(" CASE WHEN HOYA_GLASSTYPE IS NULL AND  HOYA_SOZAIDIV IS NULL THEN 'TOTAL " + SubGroup + "-GD' ELSE HOYA_GLASSTYPE END [GLASSTYPE]");

              }
              else if (hoya_diftype == "NG")
              {
                  sbSql.AppendLine(" CASE WHEN HOYA_GLASSTYPE IS NULL AND  HOYA_SOZAIDIV IS NULL THEN 'TOTAL " + SubGroup + "-NG' ELSE HOYA_GLASSTYPE END [GLASSTYPE]");

              }

              else if (hoya_diftype == "RD")
              {
                  sbSql.AppendLine(" CASE WHEN HOYA_GLASSTYPE IS NULL AND  HOYA_SOZAIDIV IS NULL THEN 'TOTAL " + SubGroup + "-RD' ELSE HOYA_GLASSTYPE END [GLASSTYPE]");

              }
              else
              {
                  sbSql.AppendLine(" CASE WHEN HOYA_GLASSTYPE IS NULL AND  HOYA_SOZAIDIV IS NULL THEN 'TOTAL " + SubGroup + "' ELSE HOYA_GLASSTYPE END [GLASSTYPE]");


              }
            
            sbSql.AppendLine(" ,HOYA_SOZAIDIV [SOZAIDIV]");
            sbSql.AppendLine(",CASE WHEN NOT  HOYA_GLASSTYPE  IS NULL  THEN 'KGS'  END [Unit]");
              sbSql.AppendLine(" ,ECL_SUBGROUP [SubGroup]");
              sbSql.AppendLine("");

              while (dtFrom <= dtTo)
              {
                  sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END)*-1[QTY]", dtFrom.Month));
                  sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END)*-1[COST]", dtFrom.Month));
                  sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                  dtFrom = dtFrom.AddMonths(1);
              }


        sbSql.AppendLine(" FROM INVENTTRANS");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
        sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
        sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
        sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
        sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
        sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
        sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST','Z1RD')");
        sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");

        sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");

       
        if (SubGroup != "OTHER")
        {
            if (SubGroup != "EB")
            {
                sbSql.AppendLine("AND ECL_SUBGROUP='" + SubGroup + "'");

            }
            else
            {
                if (hoya_diftype == "GD")
                {
                    sbSql.AppendLine("AND ECL_SUBGROUP='" + SubGroup + "'");
                    sbSql.AppendLine("AND HOYA_Diftype=5");
                }
                else if (hoya_diftype == "NG")
                {
                    sbSql.AppendLine("AND ECL_SUBGROUP='" + SubGroup + "'");
                    sbSql.AppendLine("AND HOYA_Diftype=2");
                }
                else if (hoya_diftype == "RD")
                {
                    sbSql.AppendLine("AND ECL_SUBGROUP='" + SubGroup + "'");
                    sbSql.AppendLine("AND HOYA_Diftype=9");
                }
                else 
                {
                    sbSql.AppendLine("AND ECL_SUBGROUP='" + SubGroup + "'");
                    sbSql.AppendLine("AND HOYA_Diftype NOT IN(2,5,9)");
                }

            }

        }
        else
        {
            sbSql.AppendLine("AND NOT ECL_SUBGROUP IN ('EB','GB','FC','HS' ) ");

        }

        sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
        sbSql.AppendLine(" GROUP BY HOYA_GLASSTYPE,HOYA_SOZAIDIV,ECL_SUBGROUP WITH ROLLUP");
        sbSql.AppendLine("  HAVING NOT ECL_SUBGROUP  IS NULL  OR  HOYA_GLASSTYPE IS NULL");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSummaryMaterialUSED(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" ECL_SUBGROUP   [SUBGROP]  ");            
            sbSql.AppendLine(",'KGS' [Unit]");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END)*-1[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END)*-1[COST]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


        sbSql.AppendLine(" FROM INVENTTRANS");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
        sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
        sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
        sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
        sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
        sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
        sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
        sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");
        sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
        sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
        sbSql.AppendLine("AND ECL_SUBGROUP !=''");
        sbSql.AppendLine("AND HOYA_DIFTYPE NOT IN(2,5)");
        sbSql.AppendLine("GROUP BY ECL_SUBGROUP ");
        


            sbSql.AppendLine("UNION ALL");

             dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
             dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" 'GD'   [SUBGROP]  ");
            sbSql.AppendLine(",'KGS' [Unit]");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END)*-1[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END)*-1[COST]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");
            sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");

            sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            //AND HOYA_DIFTYPE = 5
            sbSql.AppendLine("AND HOYA_DIFTYPE = 5");
            sbSql.AppendLine("AND ECL_SUBGROUP ='EB'");
            sbSql.AppendLine("GROUP BY ECL_SUBGROUP ");



            //----------------------------- NG -------------------------------------
            sbSql.AppendLine("UNION ALL");

            dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" 'NG'   [SUBGROP]  ");
            sbSql.AppendLine(",'KGS' [Unit]");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END)*-1[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END)*-1[COST]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");
            sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");

            sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            //AND HOYA_DIFTYPE = 5
            sbSql.AppendLine("AND HOYA_DIFTYPE = 2");
            sbSql.AppendLine("AND ECL_SUBGROUP ='EB'");
            sbSql.AppendLine("GROUP BY ECL_SUBGROUP ");




            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterialBalanceByItem(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            var firstDayBeforeMonth = new DateTime(MaterialOBJ.DateFrom.AddMonths(-1).Year, MaterialOBJ.DateFrom.AddMonths(-1).Month, 1);
            var lastDayOfBeforeMonth = firstDayBeforeMonth.AddMonths(1).AddDays(-1);

            sbSql.AppendLine("SELECT SUBGROUP [SUBGROUP] ");
            sbSql.AppendLine(",CASE WHEN SUMMARY.[ItemID] IS NULL AND SUBGROUP IS NULL THEN 'GRAND TOTAL' ELSE CASE WHEN SUMMARY.[ItemID] IS NULL THEN 'TOTAL' ELSE SUMMARY.[ItemID] END END ");
            sbSql.AppendLine(",NAME [NAME],HOYA_PRODUCTIONITEM ,HOYA_GLASSTYPE"); 
            sbSql.AppendLine(",HOYA_SOZAIDIV,SUM(QTY) [QTY],SUM(COST) [COST],''[Unit/Cost]");
            sbSql.AppendLine("from(");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("ECL_SUBGROUP [SUBGROUP]");
            sbSql.AppendLine(",SUMMARY.[ItemID] ");
            sbSql.AppendLine(",NAME [NAME]");
            sbSql.AppendLine(",HOYA_PRODUCTIONITEM ");
            sbSql.AppendLine(",HOYA_GLASSTYPE");
            sbSql.AppendLine(",HOYA_SOZAIDIV");
            sbSql.AppendLine(",SUM(QTY) [QTY]");
            sbSql.AppendLine(",SUM(COST) [COST]");
            sbSql.AppendLine(",''[Unit/Cost]");
            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT ");
            sbSql.AppendLine("INVENTTABLE.ECL_SUBGROUP");
            sbSql.AppendLine(",InventTrans.ITEMID [ITEMID]");
            sbSql.AppendLine(",ECORESPRODUCTTRANSLATION.NAME [NAME]");
            sbSql.AppendLine(",INVENTTABLE.HOYA_PRODUCTIONITEM ");
            sbSql.AppendLine(",INVENTTABLE.HOYA_GLASSTYPE");
            sbSql.AppendLine(",INVENTTABLE.HOYA_SOZAIDIV");
            sbSql.AppendLine(",SUM(INVENTTRANS.QTY) QTY");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
            sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine("INNER JOIN  ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
            sbSql.AppendLine(" WHERE INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");
            sbSql.AppendLine(" AND ECL_SUBGROUP IN('EB','FC','GB')");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }
            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", "1/01/2001") + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", lastDayOfBeforeMonth) + "',103)");

            sbSql.AppendLine("GROUP BY INVENTTABLE.ECL_SUBGROUP,INVENTTRANS.ITEMID,ECORESPRODUCTTRANSLATION.NAME,INVENTTABLE.HOYA_PRODUCTIONITEM,INVENTTABLE.HOYA_GLASSTYPE,INVENTTABLE.HOYA_SOZAIDIV     --) begining LEFT OUTER join");

            sbSql.AppendLine("UNION ALL");

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("INVENTTABLE.ECL_SUBGROUP [SUBGROUP]");
            sbSql.AppendLine(",InventTrans.ITEMID [ITEMID]");
            sbSql.AppendLine(",ECORESPRODUCTTRANSLATION.NAME [NAME]");
            sbSql.AppendLine(",INVENTTABLE.HOYA_PRODUCTIONITEM");
            sbSql.AppendLine(",INVENTTABLE.HOYA_GLASSTYPE");
            sbSql.AppendLine(",INVENTTABLE.HOYA_SOZAIDIV");
            sbSql.AppendLine(",SUM(INVENTTRANS.QTY) QTY");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
            sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine("INNER JOIN  ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");

            sbSql.AppendLine(" WHERE INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");
            sbSql.AppendLine(" AND ECL_SUBGROUP IN('EB','FC','GB')");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }
            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

            sbSql.AppendLine("GROUP BY INVENTTABLE.ECL_SUBGROUP,INVENTTRANS.ITEMID,ECORESPRODUCTTRANSLATION.NAME,INVENTTABLE.HOYA_PRODUCTIONITEM,INVENTTABLE.HOYA_GLASSTYPE,INVENTTABLE.HOYA_SOZAIDIV     --) begining LEFT OUTER join");

   
            sbSql.AppendLine(") summary");
            sbSql.AppendLine("GROUP BY ECL_SUBGROUP,ITEMID,NAME,HOYA_PRODUCTIONITEM,HOYA_GLASSTYPE,HOYA_SOZAIDIV ");
            sbSql.AppendLine(") as Summary");
            sbSql.AppendLine("where QTY >0");
            sbSql.AppendLine("GROUP BY [SUBGROUP],ITEMID,NAME,HOYA_PRODUCTIONITEM,HOYA_GLASSTYPE,HOYA_SOZAIDIV with rollup");
            sbSql.AppendLine("HAVING (NOT  HOYA_SOZAIDIV IS NULL) OR SUMMARY.ITEMID IS NULL ");

         ADODB.Recordset rs = new ADODB.Recordset();
         ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
         ADODBConnection.Open();

         rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

          return rs;

        }

        public ADODB.Recordset getMaterailMoveMentByItem(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            var firstDayBeforeMonth = new DateTime(MaterialOBJ.DateFrom.AddMonths(-1).Year, MaterialOBJ.DateFrom.AddMonths(-1).Month, 1);
            var lastDayOfBeforeMonth = firstDayBeforeMonth.AddMonths(1).AddDays(-1);

        sbSql.AppendLine("SELECT * FROM(");
        sbSql.AppendLine("SELECT ");
        sbSql.AppendLine("SUMMARY.[ItemID] ");
        sbSql.AppendLine(",SUMMARY.HOYA_PRODUCTIONITEM");
        if (MaterialOBJ.Factory == "GMO")
        {
            sbSql.AppendLine(" ,CASE WHEN  SUMMARY.ITEMID LIKE 'O-%'  THEN  'OPTICAL LENS' ELSE");
            sbSql.AppendLine(" CASE WHEN  SUMMARY.HOYA_LENSTYPE = 1 THEN 'SLR' ELSE");
            sbSql.AppendLine("CASE WHEN  SUMMARY.HOYA_LENSTYPE = 2 THEN 'MENISCUS' ELSE");
            sbSql.AppendLine(" CASE WHEN  SUMMARY.HOYA_LENSTYPE = 3 THEN 'NORMAL'");
            sbSql.AppendLine("  END END END END [TYPE]");
        }
        else
        {
            sbSql.AppendLine(",HOYA_Sozaidiv");
            sbSql.AppendLine(",CASE WHEN SUMMARY.HOYA_LENSTYPE  = 0 THEN null ");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 1 THEN 'SLR' ");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 2 THEN 'MENISCUS' ");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 3 THEN 'NORMAL'");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 4 THEN 'EB'");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 5 THEN 'FC'");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 6 THEN 'GB'");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 7 THEN 'HS'");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 8 THEN 'TP'");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 9 THEN 'MS'");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 10 THEN 'MMS'");
            sbSql.AppendLine("WHEN SUMMARY.HOYA_LENSTYPE  = 11 THEN 'PGS'");
            sbSql.AppendLine("END[TYPE]");

        }
        sbSql.AppendLine(",CASE WHEN SUMMARY.ITEMID LIKE 'O-%' THEN 0 ELSE  SUMMARY1.BOMQTY END [BOMQTY]");
        sbSql.AppendLine(",CASE WHEN SUMMARY.ITEMID LIKE 'O-%' THEN 0 ELSE  SUMMARY1.BOMCOST END [BOMCOST] ");
            
        sbSql.AppendLine(",ReceiveIn.ReceiveInQTY");
        sbSql.AppendLine(",ReceiveIn.ReceiveInCOST");
        sbSql.AppendLine(",InReturn.InReturnQTY");
        sbSql.AppendLine(",InReturn.InReturnCOST");
        sbSql.AppendLine(",CASE WHEN SUMMARY.ITEMID LIKE 'O-%' THEN 0 ELSE  SUMMARY.QTY END [EOMQTY]");
        sbSql.AppendLine(",CASE WHEN SUMMARY.ITEMID LIKE 'O-%' THEN 0 ELSE  SUMMARY.Cost END [EOMCOST] ");
        sbSql.AppendLine(",NG.NGQTY");
        sbSql.AppendLine(",NG.NGCOST");
        sbSql.AppendLine(",SALE.SALEQTY");
        sbSql.AppendLine(",SALE.SALECOST");
        sbSql.AppendLine(",DEAD.DEADQTY ");
        sbSql.AppendLine(",DEAD.DEADCOST");
        sbSql.AppendLine(",CASE WHEN SUMMARY.ITEMID LIKE 'O-%' THEN ReceiveIn.ReceiveInQTY ELSE  USED.USEDQTY  END [USEDQTY] ");
        sbSql.AppendLine(",CASE WHEN SUMMARY.ITEMID LIKE 'O-%' THEN ReceiveIn.ReceiveInCOST ELSE  USED.USEDCOST END [USEDCOST] ");
        sbSql.AppendLine(" FROM");

        sbSql.AppendLine(" (SELECT");
        sbSql.AppendLine("InventTrans.ITEMID [ITEMID]");
        sbSql.AppendLine(",INVENTTABLE.HOYA_PRODUCTIONITEM ");
        sbSql.AppendLine(",INVENTTABLE.HOYA_LENSTYPE");
        sbSql.AppendLine(",INVENTTABLE.HOYA_Sozaidiv");
        sbSql.AppendLine(",SUM(InventTrans.QTY) [QTY]");
        sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) [Cost]  ");
        sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
        sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
        sbSql.AppendLine("WHERE");
        sbSql.AppendLine("  INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", "1/01/2001") + "',103) ");
        sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

       sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
       sbSql.AppendLine(" AND INVENTTABLE.DATAAREAID = 'hoya'");
            
            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }

            sbSql.AppendLine(" GROUP BY InventTrans.ITEMID,INVENTTABLE.HOYA_PRODUCTIONITEM,INVENTTABLE.HOYA_LENSTYPE,HOYA_SOZAIDIV");
        sbSql.AppendLine(" )SUMMARY");
        sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP on INVENTITEMINVENTSETUP.ITEMID = SUMMARY.ITEMID");
       // sbSql.AppendLine("and INVENTITEMINVENTSETUP.DATAAREAID = SUMMARY.DATAAREAID");

        sbSql.AppendLine(" LEFT JOIN ");
        sbSql.AppendLine("( SELECT ");
        sbSql.AppendLine("INVENTTRANS.ITEMID");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) ReceiveInQTY");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) ReceiveInCOST");
        sbSql.AppendLine(",CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
        sbSql.AppendLine(" FROM INVENTTRANS ");
       
         sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine("AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID");
        sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN (SELECT ITEMID, PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID ");
        sbSql.AppendLine("FROM VENDINVOICETRANS");
       
        sbSql.AppendLine(" GROUP BY ITEMID,PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID) VENDINVOICETRANS");
        sbSql.AppendLine("ON INVENTTRANS.ITEMID=VENDINVOICETRANS.ITEMID AND INVENTTRANS.INVOICEID=VENDINVOICETRANS.INVOICEID");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
       
        sbSql.AppendLine("INNER JOIN VENDINVOICEJOUR ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
        sbSql.AppendLine("  AND VENDINVOICETRANS.INVOICEDATE=VENDINVOICEJOUR.INVOICEDATE AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP=VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
        sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID=VENDINVOICEJOUR.INTERNALINVOICEID");
        sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=VENDINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID");
    
       sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");

       sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID");
        sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");

        sbSql.AppendLine(" WHERE INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine("AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'--Physical=0, Financial=1");
        sbSql.AppendLine("AND INVENTTABLEMODULE.MODULETYPE='0'--Invent=0, Purch=1, Sales=2");
        sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
        sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY=3");
                        
            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");

            }

            sbSql.AppendLine(" AND  VENDINVOICEJOUR.INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
             sbSql.AppendLine(" AND   CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103) ");

        sbSql.AppendLine(" GROUP BY ");
        sbSql.AppendLine(" INVENTTRANS.ITEMID,INVENTTABLE.HOYA_PRODUCTIONITEM,ECORESPRODUCTTRANSLATION.NAME ) ReceiveIn ON ReceiveIn.ITEMID = SUMMARY.ITEMID");
        sbSql.AppendLine(" LEFT JOIN ");
        sbSql.AppendLine(" (SELECT ");
        sbSql.AppendLine(" INVENTTABLE.ITEMID");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) *-1 InReturnQTY");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) *-1 InReturnCOST");
        sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
        sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" WHERE INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
        sbSql.AppendLine(" AND INVENTJOURNALTABLE.HOYA_DIFTYPE in ('3')");
        sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
        sbSql.AppendLine(" AND   INVENTTRANS.DATEFINANCIAL BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine(" AND   CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103) ");
        sbSql.AppendLine(" GROUP BY  INVENTTABLE.ITEMID) InReturn ON InReturn.ITEMID = SUMMARY.ItemID");

        sbSql.AppendLine(" LEFT JOIN ");
        sbSql.AppendLine(" (SELECT ");
        sbSql.AppendLine(" INVENTTABLE.ITEMID");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) *-1 NGQTY");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) *-1 NGCOST");
        sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
        sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" WHERE INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
        sbSql.AppendLine(" AND INVENTJOURNALTABLE.HOYA_DIFTYPE IN ('2','7')");
        sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
        sbSql.AppendLine(" AND   INVENTTRANS.DATEFINANCIAL BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine(" AND   CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103) ");
        sbSql.AppendLine(" GROUP BY  INVENTTABLE.ITEMID) NG ON NG.ITEMID = SUMMARY.ItemID");


        sbSql.AppendLine(" LEFT JOIN ");
        sbSql.AppendLine(" (SELECT ");
        sbSql.AppendLine(" INVENTTABLE.ITEMID");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY)*-1 SALEQTY");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)*-1 SALECOST");
        sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
        sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" WHERE INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
        sbSql.AppendLine(" AND INVENTJOURNALTABLE.HOYA_DIFTYPE = '0'");
        sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
        sbSql.AppendLine(" AND   INVENTTRANS.DATEFINANCIAL BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine(" AND   CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103) ");
        sbSql.AppendLine(" GROUP BY  INVENTTABLE.ITEMID) SALE ON SALE.ITEMID = SUMMARY.ItemID");



        sbSql.AppendLine(" LEFT JOIN ");
        sbSql.AppendLine(" (SELECT ");
        sbSql.AppendLine(" INVENTTABLE.ITEMID");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY)*-1 DEADQTY");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)*-1 DEADCOST");
        sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
        sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" WHERE INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
        sbSql.AppendLine(" AND INVENTJOURNALTABLE.HOYA_DIFTYPE = '4'");
        sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
        sbSql.AppendLine(" AND   INVENTTRANS.DATEFINANCIAL BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine(" AND   CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103) ");
        sbSql.AppendLine(" GROUP BY  INVENTTABLE.ITEMID) DEAD ON DEAD.ITEMID = SUMMARY.ItemID");



        sbSql.AppendLine(" LEFT JOIN ");
        sbSql.AppendLine(" (SELECT ");
        sbSql.AppendLine(" INVENTTABLE.ITEMID");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.QTY) *-1 USEDQTY");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) *-1 USEDCOST");
        sbSql.AppendLine(" ,CASE WHEN SUM(INVENTTRANS.QTY)=0 THEN 0 ELSE SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/SUM(INVENTTRANS.QTY) END [COST/Unit]");
        sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" WHERE INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
        sbSql.AppendLine(" AND INVENTJOURNALTABLE.HOYA_DIFTYPE IN ('1','6')");
        sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
        sbSql.AppendLine(" AND   INVENTTRANS.DATEFINANCIAL BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine(" AND   CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103) ");
        sbSql.AppendLine(" GROUP BY  INVENTTABLE.ITEMID) USED ON USED.ITEMID = SUMMARY.ItemID");


  //BOM=====================================================================================================//
        sbSql.AppendLine(" LEFT JOIN ");
        sbSql.AppendLine(" (SELECT ");
        sbSql.AppendLine(" INVENTTRANS.ITEMID [ITEMID]");
        sbSql.AppendLine(" ,SUM(InventTrans.QTY) [BOMQTY]");
        sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)  [BOMCost]");

        sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
        sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
        sbSql.AppendLine("WHERE");

         sbSql.AppendLine("  INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", "1/01/2001") + "',103) ");
        sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", lastDayOfBeforeMonth) + "',103)");

         sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
         sbSql.AppendLine(" AND INVENTTABLE.DATAAREAID = 'hoya'");
         
            if (MaterialOBJ.Category == "All")
         {
             sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O')");
         }
         else
         {
             sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");

         }

       
        sbSql.AppendLine(" GROUP BY INVENTTRANS.ITEMID ");
        sbSql.AppendLine(" )SUMMARY1 ON SUMMARY.ITEMID = SUMMARY1.ITEMID");
        sbSql.AppendLine("WHERE INVENTITEMINVENTSETUP.STOPPED = 0");
        sbSql.AppendLine(")as Total");
        sbSql.AppendLine("WHERE NOT (total.ReceiveInQTY IS NUll  AND total.InReturnQTY IS NULL AND  Total.NGQTY IS NULL AND Total.SALEQTY IS NULL AND Total.DEADQTY IS NULL AND");
        sbSql.AppendLine(" ( Total.USEDQTY =0 AND  total.BOMQTY =0) )");

       
          
     
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterialWip(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine(" SELECT ");
        sbSql.AppendLine(" SUM(INVENTTRANS.QTY) ReceiveInQTY");
        sbSql.AppendLine(" ,SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) ReceiveInCOST");
         sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine("AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID");
        sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN (SELECT ITEMID, PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID ");
        sbSql.AppendLine("FROM VENDINVOICETRANS");
        sbSql.AppendLine(" GROUP BY ITEMID,PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID) VENDINVOICETRANS");
        sbSql.AppendLine("ON INVENTTRANS.ITEMID=VENDINVOICETRANS.ITEMID AND INVENTTRANS.INVOICEID=VENDINVOICETRANS.INVOICEID");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN VENDINVOICEJOUR ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
        sbSql.AppendLine("  AND VENDINVOICETRANS.INVOICEDATE=VENDINVOICEJOUR.INVOICEDATE AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP=VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
        sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID=VENDINVOICEJOUR.INTERNALINVOICEID");
        sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=VENDINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID");
        sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID");
        sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
        sbSql.AppendLine(" WHERE INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine("AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'--Physical=0, Financial=1");
        sbSql.AppendLine("AND INVENTTABLEMODULE.MODULETYPE='0'--Invent=0, Purch=1, Sales=2");

        sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
        sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY=3");

        sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('I')");

        sbSql.AppendLine("AND  VENDINVOICEJOUR.INVOICEDATE BETWEEN  CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
         sbSql.AppendLine("   AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103) ");



            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getLossSuri(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom  = new DateTime(MaterialOBJ.DateFrom.Year,MaterialOBJ.DateFrom.Month,1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

        sbSql.AppendLine(" SELECT ");
        sbSql.AppendLine(" INVENTTABLE.HOYA_GLASSTYPE [GLASSTYPE]");
        sbSql.AppendLine(",INVENTTABLE.HOYA_SOZAIDIV [SOZAIDIV]");

       while (dtFrom <= dtTo)
        {
           // sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY*-1 ELSE 0 END))[QTY]", dtFrom.Month));
           // sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY * (POSTEDVALUE/POSTEDQTY) ELSE 0 END))[COST]", dtFrom.Month));

            sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END) *-1 [QTY]", dtFrom.Month));
            sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END) *-1 [COST]", dtFrom.Month));

           
           
           sbSql.AppendLine(String.Format(",'' [PRICE]", dtFrom.Month));
            dtFrom = dtFrom.AddMonths(1);
        }


        sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
        sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
        sbSql.AppendLine(" WHERE INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
        sbSql.AppendLine(" AND INVENTITEMINVENTSETUP.STOPPED = '0'");

        sbSql.AppendLine(" AND INVENTSITEID = '" + MaterialOBJ.Factory + "'");
        sbSql.AppendLine("AND INVENTJOURNALTABLE.HOYA_DIFTYPE = '5'");
        sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");

        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

        //sbSql.AppendLine(" AND InventSum.PostedQty + InventSum.Received - InventSum.Deducted + InventSum.Registered - InventSum.Picked > 0");
        sbSql.AppendLine("GROUP BY INVENTTABLE.HOYA_GLASSTYPE ,INVENTTABLE.HOYA_SOZAIDIV  ");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterialReport(MaterialOBJ MaterialOBJ,string HOYA_DifType,string Qty,bool YEN)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine(" SELECT ");
            sbSql.AppendLine(" ECL_SUBGROUP [SubGroup]");


            while (dtFrom <= dtTo)
            {

                if (Qty == "Qty")
                {
                    sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY  ELSE 0 END))[QTY]", dtFrom.Month));
                }
                else if (Qty == "Cost")
                {
                    if (YEN)
                    {
                        sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))/1000[COST]", dtFrom.Month));
                    }
                    else
                    {
                        sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))[COST]", dtFrom.Month));

                    }
                   
                }
                else
                {
                    sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))/NULLIF(ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY  ELSE 0 END)),0)  [COST/QTY]", dtFrom.Month));
                }
                dtFrom = dtFrom.AddMonths(1);
            }


        sbSql.AppendLine(" FROM INVENTTRANS");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
        sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
        sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
        sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
        sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
        sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
        sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
       // sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
        sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");

        sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory +"'");
       // sbSql.AppendLine("AND  ECL_SUBGROUP IN ('EB','GB','FC','HS','TP') ");
        sbSql.AppendLine("AND  ECL_SUBGROUP IN ('EB','GB','FC','HS') ");
        sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");


        if (HOYA_DifType == "USED")
        {
            sbSql.AppendLine("AND INVENTJOURNALTABLE.HOYA_DIFTYPE IN ('1','5','2','3','6')");
        }
        else if (HOYA_DifType == "SALE")
        {
             sbSql.AppendLine("AND INVENTJOURNALTABLE.HOYA_DIFTYPE = '0'");
        }
        else
        {
            sbSql.AppendLine("AND NOT INVENTJOURNALTABLE.HOYA_DIFTYPE IN ('1','5','2','3','6','0')");
        }

        sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
        sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

        sbSql.AppendLine(" GROUP BY ECL_SUBGROUP");


        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;

        }

        public ADODB.Recordset getMaterialPurchase(MaterialOBJ MaterialOBJ, string PurQty)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine(" SELECT ");
            sbSql.AppendLine(" ECL_SUBGROUP [SubGroup]");


            while (dtFrom <= dtTo)
            {

                if (PurQty == "Qty")
                {
                    sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY  ELSE 0 END))[QTY]", dtFrom.Month));
                }
                else if (PurQty == "Cost")
                {
                    sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))[COST]", dtFrom.Month));
                }
                else
                {
                    sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))/NULLIF(ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY  ELSE 0 END)),0)  [COST/QTY]", dtFrom.Month));
                }
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS ");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
        sbSql.AppendLine("	    AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
        sbSql.AppendLine("	    AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
        sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID");
        sbSql.AppendLine("	    AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN (SELECT ITEMID, PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID ");
        sbSql.AppendLine("              FROM VENDINVOICETRANS");
        sbSql.AppendLine("		        GROUP BY ITEMID,PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID) VENDINVOICETRANS ");
        sbSql.AppendLine("		ON INVENTTRANS.ITEMID=VENDINVOICETRANS.ITEMID AND INVENTTRANS.INVOICEID=VENDINVOICETRANS.INVOICEID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN VENDINVOICEJOUR ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
        sbSql.AppendLine("      AND VENDINVOICETRANS.INVOICEDATE=VENDINVOICEJOUR.INVOICEDATE AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP=VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
        sbSql.AppendLine("      AND VENDINVOICETRANS.INTERNALINVOICEID=VENDINVOICEJOUR.INTERNALINVOICEID");
        sbSql.AppendLine("      AND VENDINVOICETRANS.DATAAREAID=VENDINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
        sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
        sbSql.AppendLine("	    AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
        sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
        sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'--Physical=0, Financial=1");
        sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'--Invent=0, Purch=1, Sales=2");
        sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY=3");

        sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
        //sbSql.AppendLine("AND  ECL_SUBGROUP IN ('EB','GB','FC','HS','TP') ");
        sbSql.AppendLine("AND  ECL_SUBGROUP IN ('EB','GB','FC','HS') ");

         if (MaterialOBJ.Category == "All") {
             sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
         }else{
             sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
         }


         foreach (DataRow dr in getAllSubSectionByFactory(MaterialOBJ.Factory).Rows)
         {
             MaterialOBJ.Section += dr["SubSection"] + ",";
         }

         sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + MaterialOBJ.Section.Replace(",", "','") + "')");

         sbSql.AppendLine(" AND  VENDINVOICEJOUR.INVOICEDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
         sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

         sbSql.AppendLine(" GROUP BY ECL_SUBGROUP");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSummaryMaterialBalance(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            var firstDayBeforeMonth = new DateTime(MaterialOBJ.DateFrom.AddMonths(-1).Year, MaterialOBJ.DateFrom.AddMonths(-1).Month, 1);
            var lastDayOfBeforeMonth = firstDayBeforeMonth.AddMonths(1).AddDays(-1);

            sbSql.AppendLine("SELECT SUBGROUP [SUBGROUP] ");
            sbSql.AppendLine(",CASE WHEN NOT SUMMARY.SUBGROUP IS NULL AND SUMMARY.HOYA_GLASSTYPE IS NULL THEN  SUMMARY.SUBGROUP +'-TOTAL' ELSE CASE WHEN SUMMARY.HOYA_GLASSTYPE IS NULL THEN 'GRAND TOTAL' ELSE SUMMARY.HOYA_GLASSTYPE END END  ");
            sbSql.AppendLine(",HOYA_SOZAIDIV");
            sbSql.AppendLine(",CASE WHEN NOT SUMMARY.HOYA_SOZAIDIV IS NULL  THEN 'KG' ELSE NULL END [Unit]");
            sbSql.AppendLine(",SUM(QTY) [QTY]");
            sbSql.AppendLine(",SUM(COST) [COST]");
            sbSql.AppendLine(",''[Unit/Cost]");
            sbSql.AppendLine("from(");
            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("ECL_SUBGROUP [SUBGROUP]");
            sbSql.AppendLine(",SUMMARY.[ItemID] ");
            sbSql.AppendLine(",NAME [NAME]");
            sbSql.AppendLine(",HOYA_PRODUCTIONITEM ");
            sbSql.AppendLine(",HOYA_GLASSTYPE");
            sbSql.AppendLine(",HOYA_SOZAIDIV");
            sbSql.AppendLine(",SUM(QTY) [QTY]");
            sbSql.AppendLine(",SUM(COST) [COST]");
            sbSql.AppendLine(",''[Unit/Cost]");
            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT ");
            sbSql.AppendLine("INVENTTABLE.ECL_SUBGROUP");
            sbSql.AppendLine(",InventTrans.ITEMID [ITEMID]");
            sbSql.AppendLine(",ECORESPRODUCTTRANSLATION.NAME [NAME]");
            sbSql.AppendLine(",INVENTTABLE.HOYA_PRODUCTIONITEM ");
            sbSql.AppendLine(",INVENTTABLE.HOYA_GLASSTYPE");
            sbSql.AppendLine(",INVENTTABLE.HOYA_SOZAIDIV");
            sbSql.AppendLine(",SUM(INVENTTRANS.QTY) QTY");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
            sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine("INNER JOIN  ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
            sbSql.AppendLine(" WHERE INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");
            sbSql.AppendLine(" AND ECL_SUBGROUP IN('EB','FC','GB')");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }
            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", "1/01/2001") + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", lastDayOfBeforeMonth) + "',103)");

            sbSql.AppendLine("GROUP BY INVENTTABLE.ECL_SUBGROUP,INVENTTRANS.ITEMID,ECORESPRODUCTTRANSLATION.NAME,INVENTTABLE.HOYA_PRODUCTIONITEM,INVENTTABLE.HOYA_GLASSTYPE,INVENTTABLE.HOYA_SOZAIDIV     --) begining LEFT OUTER join");

            sbSql.AppendLine("UNION ALL");

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("INVENTTABLE.ECL_SUBGROUP [SUBGROUP]");
            sbSql.AppendLine(",InventTrans.ITEMID [ITEMID]");
            sbSql.AppendLine(",ECORESPRODUCTTRANSLATION.NAME [NAME]");
            sbSql.AppendLine(",INVENTTABLE.HOYA_PRODUCTIONITEM");
            sbSql.AppendLine(",INVENTTABLE.HOYA_GLASSTYPE");
            sbSql.AppendLine(",INVENTTABLE.HOYA_SOZAIDIV");
            sbSql.AppendLine(",SUM(INVENTTRANS.QTY) QTY");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
            sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine("INNER JOIN  ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");

            sbSql.AppendLine(" WHERE INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");
            sbSql.AppendLine(" AND ECL_SUBGROUP IN('EB','FC','GB')");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }
            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

            sbSql.AppendLine("GROUP BY INVENTTABLE.ECL_SUBGROUP,INVENTTRANS.ITEMID,ECORESPRODUCTTRANSLATION.NAME,INVENTTABLE.HOYA_PRODUCTIONITEM,INVENTTABLE.HOYA_GLASSTYPE,INVENTTABLE.HOYA_SOZAIDIV     --) begining LEFT OUTER join");


            sbSql.AppendLine(") summary");
            sbSql.AppendLine("GROUP BY ECL_SUBGROUP,ITEMID,NAME,HOYA_PRODUCTIONITEM,HOYA_GLASSTYPE,HOYA_SOZAIDIV ");
            sbSql.AppendLine(") as Summary");
            sbSql.AppendLine("where QTY >0");
            sbSql.AppendLine("GROUP BY [SUBGROUP],HOYA_GLASSTYPE,HOYA_SOZAIDIV with rollup");
            sbSql.AppendLine("HAVING (NOT  HOYA_SOZAIDIV IS NULL) OR HOYA_GLASSTYPE IS NULL ");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

       public ADODB.Recordset BOM(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            var firstDayBeforeMonth = new DateTime(MaterialOBJ.DateFrom.AddMonths(-1).Year, MaterialOBJ.DateFrom.AddMonths(-1).Month, 1);
            var lastDayOfBeforeMonth = firstDayBeforeMonth.AddMonths(1).AddDays(-1);


            sbSql.AppendLine("SELECT  ");
            sbSql.AppendLine("INVENTTABLE.ECL_SUBGROUP ");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST  ");
            sbSql.AppendLine(",SUM(INVENTTRANS.QTY) [BOMQTY]  ");
            sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID  ");
            sbSql.AppendLine("INNER JOIN  ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT  ");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID  ");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID  ");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID  ");

          
            sbSql.AppendLine(" WHERE INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");
            sbSql.AppendLine(" AND ECL_SUBGROUP IN('EB','FC','GB','HS')");

            //All Category
            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }
            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", "1/01/2001") + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", lastDayOfBeforeMonth) + "',103)");
            sbSql.AppendLine("GROUP BY INVENTTABLE.ECL_SUBGROUP");

  

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }
    
        public ADODB.Recordset getGroupmateCompare(MaterialOBJ MaterialOBJ,bool opticlens)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);


            if (opticlens)
            {

                sbSql.AppendLine(" SELECT ");
                sbSql.AppendLine("CASE WHEN HOYA_LENSTYPE = 0 THEN 'NULL' ELSE ");
                sbSql.AppendLine("CASE WHEN  HOYA_LENSTYPE = 1 THEN 'SLR' ELSE");
                sbSql.AppendLine("CASE WHEN  HOYA_LENSTYPE = 2 THEN 'MENISCUS' ELSE");
                sbSql.AppendLine("CASE WHEN  HOYA_LENSTYPE = 3 THEN 'NORMAL'");
                sbSql.AppendLine("END END END END [TYPE]");

                sbSql.AppendLine(",CASE WHEN INVENTTABLE.ITEMID   IS NULL THEN 'TOTAL' ELSE  INVENTTABLE.ITEMID  END  [ITEM]");
                sbSql.AppendLine(" ,HOYA_PRODUCTIONITEM  [PRODUCTION ITEM]");


                while (dtFrom <= dtTo)
                {
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END) *-1[QTY]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END)*-1[COST]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",'' [/PCS]", dtFrom.Month));
                    dtFrom = dtFrom.AddMonths(1);
                }

                sbSql.AppendLine(" FROM INVENTTRANS ");
                sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
                sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
                sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
                sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
                sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
                sbSql.AppendLine(" WHERE INVENTTRANS.DATAAREAID='hoya'");
                sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
                sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
                sbSql.AppendLine(" AND INVENTSITEID = '" + MaterialOBJ.Factory + "'");
                sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
                sbSql.AppendLine(" AND INVENTTABLE.HOYA_LENSTYPE IN( '1','2','3')");
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','I')");
                sbSql.AppendLine(" AND HOYA_DIFTYPE in (1,6,3)");

                sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
                sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103) ");

                sbSql.AppendLine("   GROUP BY INVENTITEMGROUPITEM.ITEMGROUPID,HOYA_LENSTYPE,INVENTTABLE.ITEMID,HOYA_PRODUCTIONITEM WITH ROLLUP    ");
                sbSql.AppendLine(" HAVING ( NOT HOYA_PRODUCTIONITEM IS NULL) OR INVENTTABLE.ITEMID IS NULL AND  NOT HOYA_LENSTYPE IS NULL");

            }
            else
            {

                sbSql.AppendLine("SELECT");
                sbSql.AppendLine("'OPTICAL LENS'[TYPE]");
                sbSql.AppendLine(",CASE WHEN INVENTTRANS.ITEMID IS NULL THEN 'TOTAL' ELSE INVENTTRANS.ITEMID   END  [ITEMID]");
                sbSql.AppendLine(",''  [PRODUCTIONITEM]");


                while (dtFrom <= dtTo)
                {
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END) [QTY]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END)[COST]", dtFrom.Month));
                    sbSql.AppendLine(String.Format(",'' [/PCS]", dtFrom.Month));
                    dtFrom = dtFrom.AddMonths(1);
                }

                sbSql.AppendLine("FROM INVENTTRANS ");
                sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
                sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
                sbSql.AppendLine("AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
                sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
                sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
                sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
                sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
                sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID");
                sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN (SELECT ITEMID, PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID ");
                sbSql.AppendLine(" FROM VENDINVOICETRANS");

                sbSql.AppendLine("GROUP BY ITEMID,PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID) VENDINVOICETRANS ");
                sbSql.AppendLine("ON INVENTTRANS.ITEMID=VENDINVOICETRANS.ITEMID AND INVENTTRANS.INVOICEID=VENDINVOICETRANS.INVOICEID");
                sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN VENDINVOICEJOUR ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
                sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE=VENDINVOICEJOUR.INVOICEDATE AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP=VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
                sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID=VENDINVOICEJOUR.INTERNALINVOICEID");
                sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=VENDINVOICEJOUR.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
                sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID");
                sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
                sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID");
                sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
                sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
                sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
                sbSql.AppendLine("WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
                sbSql.AppendLine("AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'--Physical=0, Financial=1");
                sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'--Invent=0, Purch=1, Sales=2");
                sbSql.AppendLine(" AND INVENTSITEID = '" + MaterialOBJ.Factory + "'");
                sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY=3");
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('O')");

                foreach (DataRow dr in getAllSubSectionByFactory(MaterialOBJ.Factory).Rows)
                {
                    MaterialOBJ.Section += dr["SubSection"] + ",";
                }
                sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + MaterialOBJ.Section.Replace(",", "','") + "')");


                sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE   BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
                sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103) ");
                sbSql.AppendLine(" GROUP BY ");
                sbSql.AppendLine(" INVENTTRANS.ITEMID,ECORESPRODUCTTRANSLATION.NAME WITH ROLLUP");
                sbSql.AppendLine(" HAVING Not  ECORESPRODUCTTRANSLATION.NAME is null OR INVENTTRANS.ITEMID IS NULL");
            }




            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// End Group mat compare

        public ADODB.Recordset getCompareStock(MaterialOBJ MaterialOBJ, string filePath)
        {
            StringBuilder sbSql = new StringBuilder();

            

            if (MaterialOBJ.Factory == "RP")
            {

                sbSql.AppendLine(" SELECT  ");
                sbSql.AppendLine(" STOCK.DATE  ");
                sbSql.AppendLine(" ,STOCK.ITEMCD  ");
                sbSql.AppendLine(" ,STOCK.GLASSTYPE  ");
                sbSql.AppendLine(" ,STOCK.SOZAIDIV  ");
                sbSql.AppendLine(" ,STOCK.QUANTITY [Qty]  ");
                sbSql.AppendLine(" ,(INVENTSUM.PostedQty + INVENTSUM.Received - INVENTSUM.Deducted + INVENTSUM.Registered - INVENTSUM.Picked) as AXQty  ");
                sbSql.AppendLine("  from [192.1.87.242].[HOAX61LIVE].dbo.INVENTTABLE invent ");
                sbSql.AppendLine(" INNER  JOIN   ");
                sbSql.AppendLine(@"OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 8.0; Database=E:\Mat\Material.xlsx', 'SELECT * FROM [STOCK$]')STOCK ");
                sbSql.AppendLine(" ON invent.HOYA_PRODUCTIONITEM = STOCK.ITEMCD  ");
                sbSql.AppendLine(" AND invent.HOYA_GLASSTYPE = STOCK.GLASSTYPE  ");
                sbSql.AppendLine(" AND invent.HOYA_SOZAIDIV = STOCK.SOZAIDIV  ");
                sbSql.AppendLine("  INNER JOIN [192.1.87.242].[HOAX61LIVE].dbo.INVENTSUM  ");
                sbSql.AppendLine(" on INVENTSUM.ITEMID = invent.ITEMID  ");
                sbSql.AppendLine(" WHERE (INVENTSUM.PostedQty + INVENTSUM.Received - INVENTSUM.Deducted + INVENTSUM.Registered - INVENTSUM.Picked) > 0  ");
                sbSql.AppendLine(" AND NOT STOCK.ITEMCD IS NULL  ");

                sbSql.AppendLine(" UNION ALL  ");

                sbSql.AppendLine("SELECT");
                sbSql.AppendLine(" STOCK.DATE  ");
                sbSql.AppendLine(" ,STOCK.ITEMCD  ");
                sbSql.AppendLine(" ,STOCK.GLASSTYPE  ");
                sbSql.AppendLine(" ,STOCK.SOZAIDIV  ");
                sbSql.AppendLine(" ,STOCK.QUANTITY [Qty]  ");
                sbSql.AppendLine(" ,(INVENTSUM.PostedQty + INVENTSUM.Received - INVENTSUM.Deducted + INVENTSUM.Registered - INVENTSUM.Picked) as AXQty  ");
                sbSql.AppendLine("  from [192.1.87.242].[HOAX61LIVE].dbo.INVENTTABLE    invent   ");
                sbSql.AppendLine(" INNER  JOIN   ");
                sbSql.AppendLine(@" OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 8.0; Database=E:\Mat\Material.xlsx', 'SELECT * FROM [STOCK$]')STOCK   ");
                sbSql.AppendLine(" ON  ");
                sbSql.AppendLine("  invent.HOYA_GLASSTYPE = STOCK.GLASSTYPE   ");
                sbSql.AppendLine(" AND invent.HOYA_SOZAIDIV = STOCK.SOZAIDIV   ");
                sbSql.AppendLine(" INNER JOIN [192.1.87.242].[HOAX61LIVE].dbo.INVENTSUM   ");
                sbSql.AppendLine(" on INVENTSUM.ITEMID = invent.ITEMID  ");
                sbSql.AppendLine(" WHERE (INVENTSUM.PostedQty + INVENTSUM.Received - INVENTSUM.Deducted + INVENTSUM.Registered - INVENTSUM.Picked) > 0  ");
                sbSql.AppendLine(" AND ECL_SUBGROUP ='EB' AND STOCK.ITEMCD IS NULL  ");

            }
            else
            {
             

                sbSql.AppendLine(" SELECT '' [DATE] ");
                sbSql.AppendLine(" ,STOCK.ITEMCD  ");
                sbSql.AppendLine(" ,invent.HOYA_GLASSTYPE ,invent.HOYA_SOZAIDIV   ");
                sbSql.AppendLine(" ,STOCK.STOCK [Qty]  ");
                sbSql.AppendLine(" ,(INVENTSUM.PostedQty + INVENTSUM.Received - INVENTSUM.Deducted + INVENTSUM.Registered - INVENTSUM.Picked) as AXQty  ");
                sbSql.AppendLine(" from [192.1.87.242].[HOAX61LIVE].dbo.INVENTTABLE    invent   ");
                sbSql.AppendLine(" INNER  JOIN   ");
                sbSql.AppendLine(@" OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 8.0; Database=E:\Mat\Material.xlsx', 'SELECT * FROM [STOCK$]')STOCK  ");
                sbSql.AppendLine(" ON invent.HOYA_PRODUCTIONITEM = STOCK.ITEMCD  ");
                sbSql.AppendLine(" INNER JOIN [192.1.87.242].[HOAX61LIVE].dbo.INVENTSUM  ");
                sbSql.AppendLine(" on INVENTSUM.ITEMID = invent.ITEMID  ");
                sbSql.AppendLine(" WHERE (INVENTSUM.PostedQty + INVENTSUM.Received - INVENTSUM.Deducted + INVENTSUM.Registered - INVENTSUM.Picked) > 0  ");
                sbSql.AppendLine(" AND NOT STOCK.ITEMCD IS NULL  ");

            }

          
            
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.HOAX244Connect();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// End Group mat compare

        public ADODB.Recordset getMaterialReportMO(MaterialOBJ MaterialOBJ, string Qty, string TypeLens, bool com,string DifType)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine(" SELECT ");
            if (TypeLens == "OPTICAL LENS")
            {
                sbSql.AppendLine("''  [TYPE]");

            }
            else
            {
                sbSql.AppendLine(" CASE WHEN  HOYA_CUSTTYPE = 0 THEN 'NULL'  ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 1 THEN 'TA' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 2 THEN 'SY' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 3 THEN 'CA' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 4 THEN 'OT' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 5 THEN 'MI' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 6 THEN 'CH' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 7 THEN 'UO' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 8 THEN 'JL' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 9 THEN 'KN' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 10 THEN 'MX' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 11 THEN 'NC' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 12 THEN 'NK' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 13 THEN 'OL' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 14 THEN 'PX' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 15 THEN 'SW' ");
                sbSql.AppendLine(" END [TYPE]");

            }

            while (dtFrom <= dtTo)
            {

                if (Qty == "Qty")
                {
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY  ELSE 0 END) *-1[QTY]", dtFrom.Month));
                }
                else if (Qty == "Cost")
                {
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END) *-1[COST]", dtFrom.Month));
                }
                else
                {
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT END)/SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY  END)  [COST/QTY]", dtFrom.Month));
                }
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS ");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine(" WHERE INVENTSITEID='" + MaterialOBJ.Factory + "'");

            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
            sbSql.AppendLine(" AND INVENTTABLE.HOYA_LENSTYPE IN( '1','2','3')");

            if (TypeLens == "Normal")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
            }
            else if(DifType=="know")
            {

                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('O')");

            }

            if (DifType == "USED")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
                sbSql.AppendLine(" AND HOYA_DIFTYPE in (1,6,3)");
            }
            else if (DifType == "NG")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
                sbSql.AppendLine(" AND HOYA_DIFTYPE in (2,7)");

            }
            else if (DifType == "DS")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
                sbSql.AppendLine(" AND HOYA_DIFTYPE in (4)");
            }
            else if (DifType == "SALE")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
                sbSql.AppendLine(" AND HOYA_DIFTYPE in (0)");

            }

            

 
            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

            if (com)
            {
                sbSql.AppendLine(" AND INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT !=0");
           
            }

            if (TypeLens == "Normal")
            {
                sbSql.AppendLine(" GROUP BY HOYA_CUSTTYPE");
            }

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterialReportMOPurchase(MaterialOBJ MaterialOBJ, string Qty,string TypeLens,bool com)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine(" SELECT ");
            if (TypeLens == "OPTICAL LENS")
            {
                sbSql.AppendLine("''  [TYPE]");

            }
            else if (TypeLens == "DEAD")
            {

                sbSql.AppendLine("''  [TYPE]");
            }
            else 
            {
                sbSql.AppendLine(" CASE WHEN  HOYA_CUSTTYPE = 0 THEN 'NULL'  ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 1 THEN 'TA' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 2 THEN 'SY' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 3 THEN 'CA' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 4 THEN 'OT' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 5 THEN 'MI' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 6 THEN 'CH' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 7 THEN 'UO' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 8 THEN 'JL' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 9 THEN 'KN' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 10 THEN 'MX' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 11 THEN 'NC' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 12 THEN 'NK' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 13 THEN 'OL' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 14 THEN 'PX' ");
                sbSql.AppendLine(" WHEN  HOYA_CUSTTYPE = 15 THEN 'SW' ");
                sbSql.AppendLine(" END [TYPE]");

            }
            

            while (dtFrom <= dtTo)
            {

                if (Qty == "Qty")
                {
                    sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY  ELSE 0 END)) [QTY]", dtFrom.Month));
                }
                else if (Qty == "Cost")
                {
                    sbSql.AppendLine(String.Format(",ABS(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END) )[COST]", dtFrom.Month));
                }
                else
                {
                    sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT  END)/NULLIF(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY END),0)   [COST/QTY]", dtFrom.Month));
                }
                
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine("FROM INVENTTRANS ");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine("AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN (SELECT ITEMID, PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID ");
            sbSql.AppendLine(" FROM VENDINVOICETRANS");
            sbSql.AppendLine("GROUP BY ITEMID,PURCHID, INVOICEID, INVOICEDATE,NUMBERSEQUENCEGROUP,INTERNALINVOICEID,DATAAREAID) VENDINVOICETRANS ");
            sbSql.AppendLine("ON INVENTTRANS.ITEMID=VENDINVOICETRANS.ITEMID AND INVENTTRANS.INVOICEID=VENDINVOICETRANS.INVOICEID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN VENDINVOICEJOUR ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE=VENDINVOICEJOUR.INVOICEDATE AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP=VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID=VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=VENDINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID");
            sbSql.AppendLine(" AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP On INVENTITEMINVENTSETUP.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTRANS.DATAAREAID");
            
            
            
            sbSql.AppendLine("WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine("AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'--Physical=0, Financial=1");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'--Invent=0, Purch=1, Sales=2");
            
            sbSql.AppendLine(" AND INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY=3");
            sbSql.AppendLine(" AND INVENTITEMINVENTSETUP.STOPPED = 0");
            //sbSql.AppendLine(" AND HOYA_DIFTYPE in (1,6,3)");

            //====================================Check Opical lens================================//
            if (TypeLens == "OPTICAL LENS")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('O')");
                sbSql.AppendLine(" AND HOYA_LENSTYPE !=4");
               

            }
            else if (TypeLens == "DEAD")
            {
                sbSql.AppendLine(" AND HOYA_LENSTYPE =4");
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");

            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");

            }
            
            foreach (DataRow dr in getAllSubSectionByFactory(MaterialOBJ.Factory).Rows)
            {
                MaterialOBJ.Section += dr["SubSection"] + ",";
            }


            sbSql.AppendLine(" AND DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE IN ('" + MaterialOBJ.Section.Replace(",", "','") + "')");
            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE   BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

            if (com)
            {
                sbSql.AppendLine("AND INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT !=0");

            }

            sbSql.AppendLine(" GROUP BY HOYA_CUSTTYPE");
           
           


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterialReportMOBalance(MaterialOBJ MaterialOBJ,string Qty,string TypeLens,bool com)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);
          //  var firstDayBeforeMonth = new DateTime(MaterialOBJ.DateTo.AddMonths(1).Year, MaterialOBJ.DateTo.AddMonths(1).Month, 1);
            var lastDayOfBeforeMonth = MaterialOBJ.DateFrom.AddMonths(1).AddDays(-1);

              sbSql.AppendLine(" SELECT ");


                  sbSql.AppendLine(" CASE WHEN  TYPE = 0 THEN 'NULL'  ");
                  sbSql.AppendLine(" WHEN  TYPE = 1 THEN 'TA' ");
                  sbSql.AppendLine(" WHEN  TYPE = 2 THEN 'SY' ");
                  sbSql.AppendLine(" WHEN  TYPE = 3 THEN 'CA' ");
                  sbSql.AppendLine(" WHEN  TYPE = 4 THEN 'OT' ");
                  sbSql.AppendLine(" WHEN  TYPE = 5 THEN 'MI' ");
                  sbSql.AppendLine(" WHEN  TYPE = 6 THEN 'CH' ");
                  sbSql.AppendLine(" WHEN  TYPE = 7 THEN 'UO' ");
                  sbSql.AppendLine(" WHEN  TYPE = 8 THEN 'JL' ");
                  sbSql.AppendLine(" WHEN  TYPE = 9 THEN 'KN' ");
                  sbSql.AppendLine(" WHEN  TYPE = 10 THEN 'MX' ");
                  sbSql.AppendLine(" WHEN  TYPE = 11 THEN 'NC' ");
                  sbSql.AppendLine(" WHEN  TYPE = 12 THEN 'NK' ");
                  sbSql.AppendLine(" WHEN  TYPE = 13 THEN 'OL' ");
                  sbSql.AppendLine(" WHEN  TYPE = 14 THEN 'PX' ");
                  sbSql.AppendLine(" WHEN  TYPE = 15 THEN 'SW' ");
                  sbSql.AppendLine(" END [TYPE]");

              

                if (Qty == "Qty")
                {
                  sbSql.AppendLine(",SUM(QTY) [QTY]");
                }
                else if (Qty == "Cost")
                {
                    sbSql.AppendLine(",SUM(COST) [COST]");
                }
                else
                {
                   // sbSql.AppendLine(",SUM(COsT) /SUM(QTY) [Unit/Cost]");

                    sbSql.AppendLine(",CASE WHEN SUM(COST)=0 THEN NULL ELSE SUM(COST) /SUM(QTY)END [Unit/Cost]");
                }

                sbSql.AppendLine("FROM(");
                sbSql.AppendLine("SELECT ");
                sbSql.AppendLine("INVENTTABLE.HOYA_CUSTTYPE [TYPE]");
                sbSql.AppendLine(",InventTrans.ITEMID [ITEMID]");
                sbSql.AppendLine(",ECORESPRODUCTTRANSLATION.NAME [NAME]");
                sbSql.AppendLine(",INVENTTABLE.HOYA_PRODUCTIONITEM");
                sbSql.AppendLine(",INVENTTABLE.HOYA_GLASSTYPE");
                sbSql.AppendLine(",INVENTTABLE.HOYA_SOZAIDIV");
                sbSql.AppendLine(",SUM(INVENTTRANS.QTY) QTY");
                sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
                sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
                sbSql.AppendLine("INNER JOIN  ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
                sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
                sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
                sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
                sbSql.AppendLine(" AND INVENTSITEID='" + MaterialOBJ.Factory + "'");
                sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");

                if (TypeLens == "Normal")
                {
                    sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
                }
                else
                {

                    sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('O')");

                }

                sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", "1/01/2001") + "',103) ");
                sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", lastDayOfBeforeMonth) + "',103)");


                sbSql.AppendLine("GROUP BY HOYA_CUSTTYPE,INVENTTRANS.ITEMID,ECORESPRODUCTTRANSLATION.NAME,INVENTTABLE.HOYA_PRODUCTIONITEM,INVENTTABLE.HOYA_GLASSTYPE,INVENTTABLE.HOYA_SOZAIDIV     --) begining LEFT OUTER join");
               
  
                sbSql.AppendLine(") as Summary");
                sbSql.AppendLine("GROUP BY [TYPE] ");
    

            



            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterialReportRPBalance(MaterialOBJ MaterialOBJ, string Qty)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);
            //  var firstDayBeforeMonth = new DateTime(MaterialOBJ.DateTo.AddMonths(1).Year, MaterialOBJ.DateTo.AddMonths(1).Month, 1);
            var lastDayOfBeforeMonth = MaterialOBJ.DateFrom.AddMonths(1).AddDays(-1);

            sbSql.AppendLine(" SELECT ");
            sbSql.AppendLine(" ECL_SUBGROUP [SubGroup]");



            if (Qty == "Qty")
            {
                sbSql.AppendLine(",SUM(QTY) [QTY]");
            }
            else if (Qty == "Cost")
            {
                sbSql.AppendLine(",SUM(COST) [COST]");
            }
            else
            {
                // sbSql.AppendLine(",SUM(COsT) /SUM(QTY) [Unit/Cost]");

                sbSql.AppendLine(",CASE WHEN SUM(COST)=0 THEN NULL ELSE SUM(COST) /SUM(QTY)END [Unit/Cost]");
            }

            sbSql.AppendLine("FROM(");
            sbSql.AppendLine("SELECT ");
            sbSql.AppendLine("INVENTTABLE.ECL_SUBGROUP");
            sbSql.AppendLine(",SUM(INVENTTRANS.QTY) QTY");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) COST");
            sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine("INNER JOIN  ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
            sbSql.AppendLine(" AND INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");

            sbSql.AppendLine("AND  ECL_SUBGROUP IN ('EB','GB','FC','HS','TP') ");

            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", "1/01/2001") + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", lastDayOfBeforeMonth) + "',103)");


            sbSql.AppendLine("GROUP BY ECL_SUBGROUP   --) begining LEFT OUTER join");


            sbSql.AppendLine(") as Summary");
            sbSql.AppendLine("GROUP BY ECL_SUBGROUP ");






            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;
        }


        //RP
        public ADODB.Recordset getMaterailPurchaseYear(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine(" SELECT  ");
            sbSql.AppendLine(" CASE WHEN HOYA_GLASSTYPE IS NULL AND  ECL_SUBGROUP  IS NULL THEN 'GRAND TOTAL' ELSE ECL_SUBGROUP  END [SUBGROUP]");
            sbSql.AppendLine(",CASE WHEN HOYA_GLASSTYPE IS NULL  AND NOT ECL_SUBGROUP IS NULL  THEN 'TOTAL' ELSE HOYA_GLASSTYPE END [GLASSTYPE]");
            //sbSql.AppendLine(",HOYA_GLASSTYPE [GLASSTYPE]");
            sbSql.AppendLine(",HOYA_SOZAIDIV [SOSAIDIV]");
            sbSql.AppendLine(",VENDINVOICETRANS.PURCHUNIT [UNIT]");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN VENDINVOICETRANS.QTY ELSE 0 END)  [QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} AND VENDINVOICEJOUR.CURRENCYCODE ='JPS' THEN VENDINVOICETRANS.LINEAMOUNT  ELSE 0 END) [JPY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} AND VENDINVOICEJOUR.CURRENCYCODE ='USS' THEN VENDINVOICETRANS.LINEAMOUNT  ELSE 0 END) [USD]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0}  THEN VENDINVOICETRANS.LINEAMOUNT * (VENDINVOICEJOUR.EXCHRATE/100) + ISNULL(MARKUPTRANS.VALUE,0) ELSE 0 END) [THB]", dtFrom.Month));

                sbSql.AppendLine(String.Format(",'' [/KG]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine("  FROM VENDINVOICEJOUR ");
            sbSql.AppendLine("INNER JOIN VENDINVOICETRANS ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE = VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine("AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP = VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID = VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=VENDINVOICETRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID ");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");

            //10/15/2018
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTABLE.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");


            sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(MARKUPCODE) MARKUPCODE,MAX(TRANSRECID) TRANSRECID, SUM(VALUE) VALUE,MAX(MARKUPTRANS.CURRENCYCODE) CURRENCYCODE, COUNT(RECID) RECIDM ");
            sbSql.AppendLine("FROM MARKUPTRANS WHERE TRANSTABLEID='492'  GROUP BY TRANSRECID)");
            sbSql.AppendLine("MARKUPTRANS ON VENDINVOICETRANS.RECID=MARKUPTRANS.TRANSRECID ");
            sbSql.AppendLine("WHERE ");
            sbSql.AppendLine("  VENDINVOICEJOUR.INVOICEDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" AND INVENTSITEID = '" + MaterialOBJ.Factory + "'");
            //sbSql.AppendLine("AND  ECL_SUBGROUP IN ('EB','FC')");


            // Edit 12/10/2018
            sbSql.AppendLine("AND NOT ECL_SUBGROUP IN ('MO','FG')");
            
            sbSql.AppendLine("AND VENDINVOICEJOUR.DATAAREAID = 'hoya'");

           // if (Numbersequence)
           // {
                sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ('RP-MT','RP-CNM')");

           // }
           // else
           // {
            //    sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP = 'RP-CNM'");

           // }

                if (MaterialOBJ.Category == "All")
                {
                    sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
                }
                else
                {
                    sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
                }


            sbSql.AppendLine(" GROUP BY ECL_SUBGROUP ,HOYA_GLASSTYPE ,HOYA_SOZAIDIV,VENDINVOICETRANS.PURCHUNIT WITH ROLLUP  ");
            sbSql.AppendLine(" HAVING NOT VENDINVOICETRANS.PURCHUNIT  IS NULL OR HOYA_GLASSTYPE IS NULL");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterailPurchaseYearSummary(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine(" SELECT  ");
            sbSql.AppendLine("  CASE WHEN   ECL_SUBGROUP  IS NULL THEN 'TOTAL' ELSE ECL_SUBGROUP  END [SUBGROUP]");
            sbSql.AppendLine(",VENDINVOICETRANS.PURCHUNIT [UNIT]");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN VENDINVOICETRANS.QTY ELSE 0 END)  [QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} AND VENDINVOICEJOUR.CURRENCYCODE ='JPS' THEN VENDINVOICETRANS.LINEAMOUNT  ELSE 0 END) [JPY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} AND VENDINVOICEJOUR.CURRENCYCODE ='USS' THEN VENDINVOICETRANS.LINEAMOUNT  ELSE 0 END) [USD]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0}  THEN VENDINVOICETRANS.LINEAMOUNT * (VENDINVOICEJOUR.EXCHRATE/100) + ISNULL(MARKUPTRANS.VALUE,0) ELSE 0 END) [THB]", dtFrom.Month));

                sbSql.AppendLine(String.Format(",'' [/KG]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine("  FROM VENDINVOICEJOUR ");
            sbSql.AppendLine("INNER JOIN VENDINVOICETRANS ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE = VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine("AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP = VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID = VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=VENDINVOICETRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID ");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");


            //10/15/2018
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTABLE.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");


            sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(MARKUPCODE) MARKUPCODE,MAX(TRANSRECID) TRANSRECID, SUM(VALUE) VALUE,MAX(MARKUPTRANS.CURRENCYCODE) CURRENCYCODE, COUNT(RECID) RECIDM ");
            sbSql.AppendLine("FROM MARKUPTRANS WHERE TRANSTABLEID='492'  GROUP BY TRANSRECID)");
            sbSql.AppendLine("MARKUPTRANS ON VENDINVOICETRANS.RECID=MARKUPTRANS.TRANSRECID ");

            //sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(TRANSRECID) TRANSRECID, SUM(VALUE) VALUE,DATAAREAID ");
            //sbSql.AppendLine("FROM MARKUPTRANS WHERE TRANSTABLEID='492'  GROUP BY TRANSRECID,DATAAREAID)");
            //sbSql.AppendLine("MARKUPTRANS ON VENDINVOICETRANS.RECID=MARKUPTRANS.TRANSRECID ");
            //sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");

            sbSql.AppendLine("WHERE ");
            sbSql.AppendLine("  VENDINVOICEJOUR.INVOICEDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" AND INVENTSITEID = '" + MaterialOBJ.Factory + "'");
            //sbSql.AppendLine("AND  ECL_SUBGROUP IN ('EB','FC')");
            sbSql.AppendLine("AND NOT ECL_SUBGROUP IN ('MO','FG')");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }


            sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ('RP-MT','RP-CNM')");
            sbSql.AppendLine("AND VENDINVOICEJOUR.DATAAREAID = 'hoya'");

            sbSql.AppendLine("  GROUP BY ECL_SUBGROUP ,VENDINVOICETRANS.PURCHUNIT WITH ROLLUP  ");
            sbSql.AppendLine("  HAVING NOT VENDINVOICETRANS.PURCHUNIT  IS NULL OR ECL_SUBGROUP IS NULL");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }



        public ADODB.Recordset getDetailMaterialReportForGMO(MaterialOBJ MaterialOBJ, string hoya_diftype)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");

            //sbSql.AppendLine("  CASE WHEN INVENTTABLE.ITEMID IS NULL THEN 'TOTAL' +' "+hoya_diftype+"'  ELSE INVENTTABLE.ITEMID END [ITEM],");
            sbSql.AppendLine("  INVENTTABLE.ITEMID  [ITEM],");
            sbSql.AppendLine("  'PCS' [Unit],");
            sbSql.AppendLine(" CASE WHEN INVENTTABLE.ITEMID IS NULL THEN ECL_SUBGROUP ELSE ECL_SUBGROUP END  ECL_SUBGROUP");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END))*-1[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))*-1[COST]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            //sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");


            //USED-USED-RT-RETRUN 1,6,3  #SALE 0 #DEAD 4 #NG-NG-RT 2,7 
            if (hoya_diftype.ToString() == "USED")
            {
                sbSql.AppendLine("AND  HOYA_Diftype IN(1,6,3)");
            }
            else if (hoya_diftype.ToString() == "SALE")
            {
                sbSql.AppendLine("AND  HOYA_Diftype IN(0)");
            }
            else if (hoya_diftype.ToString() == "DEAD")
            {
                sbSql.AppendLine("AND  HOYA_Diftype IN(4)");
            }
            else if (hoya_diftype.ToString() == "NG")
            {
                sbSql.AppendLine("AND  HOYA_Diftype IN(2,7)");
            }
            


            sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }

            //Shipment
            //sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");

            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine("GROUP BY ECL_SUBGROUP,INVENTTABLE.ITEMID  --WITH ROLLUP");
            //sbSql.AppendLine("HAVING NOT ECL_SUBGROUP IS NULL");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getDetailMaterialReportForGMO2(MaterialOBJ MaterialOBJ,string GroupID)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");

            if (GroupID == "I")
            {
                 sbSql.AppendLine("   CASE WHEN INVENTTABLE.ITEMID IS NULL THEN 'TOTAL ' + ECL_SUBGROUP  ELSE INVENTTABLE.ITEMID END [ITEM],");
            }
            else
            {
                sbSql.AppendLine(" INVENTTABLE.ITEMID  [ITEM],");
            }
           
            sbSql.AppendLine("  'PCS' [Unit],");
            sbSql.AppendLine(" CASE WHEN INVENTTABLE.ITEMID IS NULL THEN ECL_SUBGROUP ELSE ECL_SUBGROUP END  ECL_SUBGROUP");

            while (dtFrom <= dtTo)
            {
               // sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END))*-1[QTY]", dtFrom.Month));
                //sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))*-1[COST]", dtFrom.Month));
                 sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END))[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))[COST]", dtFrom.Month));
                
                sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            //sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");


            //USED-USED-RT-RETRUN 1,6,3  #SALE 0 #DEAD 4 #NG-NG-RT 2,7 


            sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('"+GroupID+"')");
            /*
            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('O','I')");
            }
            else
            {
                //sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");

            }
            */
            //Shipment
            sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 3");

            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

            if (GroupID == "I")
            {
                sbSql.AppendLine("GROUP BY ECL_SUBGROUP,INVENTTABLE.ITEMID  WITH ROLLUP");
                sbSql.AppendLine("HAVING NOT ECL_SUBGROUP IS NULL");
            }
            else
            {
                sbSql.AppendLine("GROUP BY ECL_SUBGROUP,INVENTTABLE.ITEMID -- WITH ROLLUP");
            }


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSummaryMaterialReportForGMO(MaterialOBJ MaterialOBJ, string hoya_diftype)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" CASE WHEN  ECL_SUBGROUP != '' THEN [ECL_SUBGROUP] + ' " + hoya_diftype + "'   ELSE ECL_SUBGROUP END [SUBGROUP],");
            sbSql.AppendLine(" 'PCS' [Unit]");


            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END))*-1[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))*-1[COST]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            //sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");
           
            sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");

            //USED-USED-RT-RETRUN 1,6,3  #SALE 0 #DEAD 4 #NG-NG-RT 2,7 
            if (hoya_diftype.ToString() == "USED")
            {
                sbSql.AppendLine("AND  HOYA_Diftype IN(1,6,3)");
            }
            else if (hoya_diftype.ToString() == "SALE")
            {
                sbSql.AppendLine("AND  HOYA_Diftype IN(0)");
            }
            else if (hoya_diftype.ToString() == "DEAD")
            {
                sbSql.AppendLine("AND  HOYA_Diftype IN(4)");
            }
            else if (hoya_diftype.ToString() == "NG")
            {
                sbSql.AppendLine("AND  HOYA_Diftype IN(2,7)");
            }




            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }

            //Shipment
           // sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");

            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine("GROUP BY ECL_SUBGROUP ");
          //  sbSql.AppendLine("HAVING NOT ECL_SUBGROUP IS NULL");

           
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getSummaryMaterialReportForGMO2(MaterialOBJ MaterialOBJ,string GroupID)
        {
                
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine(" CASE WHEN  ECL_SUBGROUP != '' THEN [ECL_SUBGROUP]  ELSE ECL_SUBGROUP END [SUBGROUP],");
            sbSql.AppendLine(" 'PCS' [Unit]");


            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END))[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))[COST]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            //sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");

            sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");
            sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('"+GroupID+"')");
            /*
            */
            //Shipment
            sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 3");


            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine("GROUP BY ECL_SUBGROUP ");
            //  sbSql.AppendLine("HAVING NOT ECL_SUBGROUP IS NULL");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;


        }




        public ADODB.Recordset getDetailMaterialReportForPO(MaterialOBJ MaterialOBJ, string hoya_diftype)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");
           // sbSql.AppendLine("CASE WHEN INVENTTRANS.ITEMID IS NULL THEN 'TOTAL PRESS LENS' ELSE INVENTTRANS.ITEMID END [ITEMID]");
            sbSql.AppendLine("INVENTTRANS.ITEMID  [ITEMID]");
            sbSql.AppendLine(", 'PCS' [Unit]");
            sbSql.AppendLine("");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END))*-1[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))*-1[COST]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            //sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");

            sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','W')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }

            //Shipment
            sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");

            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" GROUP BY INVENTTRANS.ITEMID --WITH ROLLUP ");
            //  sbSql.AppendLine("HAVING NOT ECL_SUBGROUP IS NULL");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// end GetMaterailForPO

        public ADODB.Recordset getSummaryMaterialReportForPO(MaterialOBJ MaterialOBJ, string hoya_diftype)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN ECL_SUBGROUP IS NULL OR ECL_SUBGROUP = '' THEN 'TOTAL PRESS LENS' ELSE ECL_SUBGROUP END [ECL_SUPGROUP]");
            sbSql.AppendLine(", 'PCS' [Unit]");
            sbSql.AppendLine("");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.QTY ELSE 0 END))*-1[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",(SUM(CASE WHEN MONTH(DATEFINANCIAL)={0} THEN INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT ELSE 0 END))*-1[COST]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",'' [COST/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine(" FROM INVENTTRANS");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON INVENTTRANSPOSTING.DEFAULTDIMENSION=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUE ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE=DIMENSIONATTRIBUTEVALUE.RECID");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTE ON DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE=DIMENSIONATTRIBUTE.RECID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine("LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE DIMENSIONATTRIBUTE.NAME='D3_SUBSECTION' AND INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine(" AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            //sbSql.AppendLine("AND DISPLAYVALUE IN ('Z1BR','Z1AC','Z1QA','Z1ST')");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");

            sbSql.AppendLine("AND INVENTDIM.INVENTSITEID='" + MaterialOBJ.Factory + "'");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','W')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }

            //Shipment
            sbSql.AppendLine("AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");

            sbSql.AppendLine(" AND INVENTTRANS.DATEFINANCIAL  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine("  GROUP BY ECL_SUBGROUP");
            //  sbSql.AppendLine("HAVING NOT ECL_SUBGROUP IS NULL");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// end GetMaterailForPO




        public ADODB.Recordset getMaterailPurchaseForPO(DataTable dt,MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT purchase.ITEMID,purchase.NAME,purchase.[Vender],purchase.PURCHUNIT");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM([QTY" + String.Format("{0:yyMM}", dr[0]) + "]) [QTY" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([JPS" + String.Format("{0:yyMM}", dr[0]) + "]) [JPS" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([USS" + String.Format("{0:yyMM}", dr[0]) + "]) [USS" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([THB" + String.Format("{0:yyMM}", dr[0]) + "]) [THB" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([(THB)" + String.Format("{0:yyMM}", dr[0]) + "]) [(THB)" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([/KG" + String.Format("{0:yyMM}", dr[0]) + "]) [/KG" + String.Format("{0:yyMM}", dr[0]) + "] ");
            }

            sbSql.AppendLine("FROM(");

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("CASE WHEN INVENTTABLE.ITEMID IS NULL THEN 'TOTAL' ELSE INVENTTABLE.ITEMID END [ITEMID] ,");
            sbSql.AppendLine("ECORESPRODUCTTRANSLATION.NAME,");
            sbSql.AppendLine(" CASE WHEN HOYA_VENDERID  = '' THEN VENDINVOICEJOUR.INVOICEACCOUNT ELSE HOYA_VENDERID END  [Vender],");
            sbSql.AppendLine(" VENDINVOICETRANS.PURCHUNIT");


            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN VENDINVOICETRANS.QTY END )[QTY" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='JPS' AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END) [JPS" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='USS' AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END) [USS" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='THB'AND PurchTable.Payment !='NOCOM'  THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END) [THB" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN CASE WHEN  PurchTable.Payment !='NOCOM' THEN VENDINVOICETRANS.LINEAMOUNT * (VENDINVOICEJOUR.EXCHRATE/100) +ISNULL(MARKUPTRANS.VALUE,0) ELSE 0 END  END )[(THB)" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN '' ELSE 0 END )[/KG" + String.Format("{0:yyMM}", dr[0]) + "]");
              

            }


            sbSql.AppendLine("  FROM VENDINVOICEJOUR  ");
            sbSql.AppendLine("INNER JOIN VENDINVOICETRANS ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE = VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine("AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP = VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID = VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");


            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=VENDINVOICETRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTABLE.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");

            sbSql.AppendLine("INNER JOIN VENDTABLE On VENDTABLE.ACCOUNTNUM = VENDINVOICEJOUR.INVOICEACCOUNT");
            sbSql.AppendLine("AND VENDTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");

            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
             sbSql.AppendLine("INNER JOIN PURCHTABLE ON PURCHTABLE.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND PURCHTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");


            sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(TRANSRECID) TRANSRECID, SUM(VALUE) VALUE FROM MARKUPTRANS ");
            sbSql.AppendLine("WHERE TRANSTABLEID='492'   AND MARKUPTRANS.DATAAREAID = 'hoya' AND MARKUPCODE IN ('Insurance','Freight') ");
            sbSql.AppendLine(" AND  MARKUPTRANS.TRANSDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" GROUP BY TRANSRECID)");
            sbSql.AppendLine("MARKUPTRANS ON VENDINVOICETRANS.RECID=MARKUPTRANS.TRANSRECID ");
         
        

         
            sbSql.AppendLine("WHERE");
            sbSql.AppendLine("INVENTDIM.INVENTSITEID = '" + MaterialOBJ.strFactory + "'");
            sbSql.AppendLine("AND INVENTDIM.DATAAREAID = 'hoya'");
            sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ( 'PO-DM', 'PO-IM', 'PO-MT', 'PO-CND', 'PO-CNI', 'PO-CNM' ) ");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }


            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");



            sbSql.AppendLine("GROUP BY INVENTTABLE.ITEMID,ECORESPRODUCTTRANSLATION.NAME,HOYA_VENDERID, VENDINVOICEJOUR.INVOICEACCOUNT,VENDINVOICETRANS.PURCHUNIT --WITH ROLLUP");
            sbSql.AppendLine(")as purchase GROUP BY purchase.ITEMID,purchase.NAME,purchase.Vender,purchase.PURCHUNIT");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterailPurchaseSummaryForPO(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");
           // sbSql.AppendLine("ECL_SUBGROUP  [SUBGROUP],");
            sbSql.AppendLine("'PRESSED LENS'  [SUBGROUP],");
            sbSql.AppendLine(" VENDINVOICETRANS.PURCHUNIT");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN VENDINVOICETRANS.QTY ELSE 0 END)[QTY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='JPS' AND PurchTable.Payment !='NOCOM'  THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END)[JPY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='USS' AND PurchTable.Payment !='NOCOM'  THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END)[USD]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='THB' AND PurchTable.Payment !='NOCOM'  THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END)[THB]", dtFrom.Month));
                //sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN VENDINVOICETRANS.LINEAMOUNT * (VENDINVOICEJOUR.EXCHRATE/100) +ISNULL(MARKUPTRANS.VALUE,0) ELSE 0  END)[(THB)]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  PurchTable.Payment !='NOCOM'  THEN VENDINVOICETRANS.LINEAMOUNT * (VENDINVOICEJOUR.EXCHRATE/100) + ISNULL(MARKUPTRANS.VALUE,0)  ELSE 0 END ELSE 0  END)[(THB)]", dtFrom.Month));
    


                sbSql.AppendLine(String.Format(",'' [/KG]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }

            sbSql.AppendLine("  FROM VENDINVOICEJOUR  ");
            sbSql.AppendLine("INNER JOIN VENDINVOICETRANS ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE = VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine("AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP = VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID = VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");

            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=VENDINVOICETRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTABLE.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");

            sbSql.AppendLine("INNER JOIN VENDTABLE On VENDTABLE.ACCOUNTNUM = VENDINVOICEJOUR.INVOICEACCOUNT");
            sbSql.AppendLine("AND VENDTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");

            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            

            sbSql.AppendLine("INNER JOIN PURCHTABLE ON PURCHTABLE.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND PURCHTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");


            //sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(TRANSRECID) TRANSRECID, SUM(VALUE) VALUE FROM MARKUPTRANS ");
            //sbSql.AppendLine("WHERE TRANSTABLEID='492'   AND MARKUPTRANS.DATAAREAID = 'hoya' AND MARKUPCODE IN ('Insurance','Freight') GROUP BY TRANSRECID)");
            //sbSql.AppendLine("MARKUPTRANS ON VENDINVOICETRANS.RECID=MARKUPTRANS.TRANSRECID ");

            sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(TRANSRECID) TRANSRECID, SUM(VALUE) VALUE FROM MARKUPTRANS ");
            sbSql.AppendLine("WHERE TRANSTABLEID='492'   AND MARKUPTRANS.DATAAREAID = 'hoya' AND MARKUPCODE IN ('Insurance','Freight') ");
            sbSql.AppendLine(" AND  MARKUPTRANS.TRANSDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" GROUP BY TRANSRECID)");
            sbSql.AppendLine("MARKUPTRANS ON VENDINVOICETRANS.RECID=MARKUPTRANS.TRANSRECID ");
         


            sbSql.AppendLine("WHERE");
            sbSql.AppendLine("INVENTDIM.INVENTSITEID = '" + MaterialOBJ.strFactory + "'");
            sbSql.AppendLine("AND INVENTDIM.DATAAREAID = 'hoya'");

            sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ( 'PO-DM', 'PO-IM', 'PO-MT', 'PO-CND', 'PO-CNI', 'PO-CNM' ) ");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }


            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");



            sbSql.AppendLine("GROUP BY VENDINVOICETRANS.PURCHUNIT --WITH ROLLUP");
            //sbSql.AppendLine("HAVING NOT VENDINVOICETRANS.PURCHUNIT IS NULL OR ECL_SUBGROUP IS NULL");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

       

        public ADODB.Recordset getMaterailPurchaseForGMO(DataTable dt,MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT purchase.SUBGROUP,CASE WHEN purchase.ITEMID IS NULL THEN 'TOTAL' ELSE purchase.ITEMID END [ITEMID] ,purchase.NAME,purchase.[Vender],purchase.PURCHUNIT");

            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM([QTY" + String.Format("{0:yyMM}", dr[0]) + "]) [QTY" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([JPS" + String.Format("{0:yyMM}", dr[0]) + "]) [JPS" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([USS" + String.Format("{0:yyMM}", dr[0]) + "]) [USS" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([THB" + String.Format("{0:yyMM}", dr[0]) + "]) [THB" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([(THB)" + String.Format("{0:yyMM}", dr[0]) + "]) [(THB)" + String.Format("{0:yyMM}", dr[0]) + "] ");
                sbSql.AppendLine(" ,SUM([/KG" + String.Format("{0:yyMM}", dr[0]) + "]) [/KG" + String.Format("{0:yyMM}", dr[0]) + "] ");
            }

            sbSql.AppendLine("FROM(");

            sbSql.AppendLine("SELECT ECL_SUBGROUP [SUBGROUP],");
            sbSql.AppendLine(" INVENTTABLE.ITEMID [ITEMID] ,");
            sbSql.AppendLine("ECORESPRODUCTTRANSLATION.NAME,");
            sbSql.AppendLine(" CASE WHEN HOYA_VENDERID  = '' THEN VENDINVOICEJOUR.INVOICEACCOUNT ELSE HOYA_VENDERID END  [Vender],");
            sbSql.AppendLine(" VENDINVOICETRANS.PURCHUNIT");


            foreach (DataRow dr in dt.Rows)
            {
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN VENDINVOICETRANS.QTY END )[QTY" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='JPS' AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END) [JPS" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='USS' AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END) [USS" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='THB'AND PurchTable.Payment !='NOCOM'  THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END) [THB" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN CASE WHEN  PurchTable.Payment !='NOCOM' THEN VENDINVOICETRANS.LINEAMOUNT * (VENDINVOICEJOUR.EXCHRATE/100) +ISNULL(MARKUPTRANS.VALUE,0) ELSE 0 END  END )[(THB)" + String.Format("{0:yyMM}", dr[0]) + "]");
                sbSql.AppendLine(" ,SUM(CASE WHEN CONVERT(CHAR(4),VENDINVOICEJOUR.INVOICEDATE,12)='" + String.Format("{0:yyMM}", dr[0]) + "' THEN '' ELSE 0 END )[/KG" + String.Format("{0:yyMM}", dr[0]) + "]");


            }


            sbSql.AppendLine("  FROM VENDINVOICEJOUR  ");
            sbSql.AppendLine("INNER JOIN VENDINVOICETRANS ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE = VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine("AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP = VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID = VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");

            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=VENDINVOICETRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            
            sbSql.AppendLine("INNER JOIN INVENTDIM ON VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTABLE.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");

            sbSql.AppendLine("INNER JOIN VENDTABLE On VENDTABLE.ACCOUNTNUM = VENDINVOICEJOUR.INVOICEACCOUNT");
            sbSql.AppendLine("AND VENDTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("INNER JOIN PURCHTABLE ON PURCHTABLE.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND PURCHTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");


            sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(TRANSRECID) TRANSRECID, SUM(VALUE) VALUE FROM MARKUPTRANS ");
            sbSql.AppendLine("WHERE TRANSTABLEID='492'   AND MARKUPTRANS.DATAAREAID = 'hoya' AND MARKUPCODE IN ('Insurance','Freight') ");
            sbSql.AppendLine(" AND  MARKUPTRANS.TRANSDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" GROUP BY TRANSRECID)");
            sbSql.AppendLine("MARKUPTRANS ON VENDINVOICETRANS.RECID=MARKUPTRANS.TRANSRECID ");




            sbSql.AppendLine("WHERE");
            sbSql.AppendLine("INVENTDIM.INVENTSITEID = '" + MaterialOBJ.strFactory + "'");
            sbSql.AppendLine("AND INVENTDIM.DATAAREAID = 'hoya'");
            sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ('MO-DM', 'MO-IM', 'MO-MT', 'MO-CND', 'MO-CNI', 'MO-CNM')  ");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }


            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");



            sbSql.AppendLine("GROUP BY ECL_SUBGROUP,INVENTTABLE.ITEMID,ECORESPRODUCTTRANSLATION.NAME,HOYA_VENDERID, VENDINVOICEJOUR.INVOICEACCOUNT,VENDINVOICETRANS.PURCHUNIT --WITH ROLLUP");
            sbSql.AppendLine(")as purchase GROUP BY purchase.SUBGROUP,purchase.ITEMID,purchase.NAME,purchase.Vender,purchase.PURCHUNIT WITH ROLLUP");
            sbSql.AppendLine("HAVING NOT purchase.PURCHUNIT IS NULL OR purchase.ITEMID IS NULL AND NOT purchase.SUBGROUP IS NULL");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getMaterailPurchaseSummaryForGMO(MaterialOBJ MaterialOBJ)
        {

            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT");
            sbSql.AppendLine("ECL_SUBGROUP  [SUBGROUP],");
            sbSql.AppendLine(" VENDINVOICETRANS.PURCHUNIT");

            while (dtFrom <= dtTo)
            {
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN VENDINVOICETRANS.QTY ELSE 0 END)[QTY]", dtFrom.Month));
                //sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='JPS' AND ECL_SUBGROUP != 'OPTICAL'  AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END)[JPY]", dtFrom.Month));
                //sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='USS' AND ECL_SUBGROUP != 'OPTICAL' AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END)[USD]", dtFrom.Month));
               // sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='THB' AND ECL_SUBGROUP != 'OPTICAL' AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END)[THB]", dtFrom.Month));


                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='JPS' AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END)[JPY]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='USS' AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END)[USD]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  VENDINVOICEJOUR.CURRENCYCODE ='THB' AND PurchTable.Payment !='NOCOM' THEN  VENDINVOICETRANS.LINEAMOUNT ELSE 0 END END)[THB]", dtFrom.Month));
               
                
                
                
                sbSql.AppendLine(String.Format(",SUM(CASE WHEN MONTH(VENDINVOICEJOUR.INVOICEDATE)={0} THEN CASE WHEN  PurchTable.Payment !='NOCOM' THEN VENDINVOICETRANS.LINEAMOUNT * (VENDINVOICEJOUR.EXCHRATE/100) + ISNULL(MARKUPTRANS.VALUE,0)  ELSE 0 END ELSE 0  END)[(THB)]", dtFrom.Month));
                sbSql.AppendLine(String.Format(",'' [/QTY]", dtFrom.Month));
                dtFrom = dtFrom.AddMonths(1);
            }


            sbSql.AppendLine("  FROM VENDINVOICEJOUR  ");
            sbSql.AppendLine("INNER JOIN VENDINVOICETRANS ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE = VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine("AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP = VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID = VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");


            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=VENDINVOICETRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTABLE.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");

            sbSql.AppendLine("INNER JOIN VENDTABLE On VENDTABLE.ACCOUNTNUM = VENDINVOICEJOUR.INVOICEACCOUNT");
            sbSql.AppendLine("AND VENDTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");

            sbSql.AppendLine("INNER JOIN PURCHTABLE ON PURCHTABLE.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND PURCHTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");


           // sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(TRANSRECID) TRANSRECID, SUM(VALUE) VALUE FROM MARKUPTRANS ");
           // sbSql.AppendLine("WHERE TRANSTABLEID='492'   AND MARKUPTRANS.DATAAREAID = 'hoya' AND MARKUPCODE IN ('Insurance','Freight') GROUP BY TRANSRECID)");
           // sbSql.AppendLine("MARKUPTRANS ON VENDINVOICETRANS.RECID=MARKUPTRANS.TRANSRECID ");

            sbSql.AppendLine("LEFT OUTER JOIN (SELECT MAX(TRANSRECID) TRANSRECID, SUM(VALUE) VALUE FROM MARKUPTRANS ");
            sbSql.AppendLine("WHERE TRANSTABLEID='492'   AND MARKUPTRANS.DATAAREAID = 'hoya' AND MARKUPCODE IN ('Insurance','Freight') ");
            sbSql.AppendLine(" AND  MARKUPTRANS.TRANSDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" GROUP BY TRANSRECID)");
            sbSql.AppendLine("MARKUPTRANS ON VENDINVOICETRANS.RECID=MARKUPTRANS.TRANSRECID ");



            sbSql.AppendLine("WHERE");
            sbSql.AppendLine("INVENTDIM.INVENTSITEID = '" + MaterialOBJ.strFactory + "'");
            sbSql.AppendLine("AND INVENTDIM.DATAAREAID = 'hoya'");
            sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ('MO-DM', 'MO-IM', 'MO-MT', 'MO-CND', 'MO-CNI', 'MO-CNM') ");

            if (MaterialOBJ.Category == "All")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
            }
            else
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
            }


            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");



            sbSql.AppendLine("GROUP BY ECL_SUBGROUP, VENDINVOICETRANS.PURCHUNIT,VENDINVOICEJOUR.CURRENCYCODE");
           // sbSql.AppendLine("HAVING NOT VENDINVOICETRANS.PURCHUNIT IS NULL OR INVENTTABLE.ITEMID IS NULL AND NOT ECL_SUBGROUP IS NULL");

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }




        public ADODB.Recordset getMaterailPurchaseVender(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT [Vender]");
            sbSql.AppendLine("FROM (SELECT");

            sbSql.AppendLine("CASE WHEN HOYA_VENDERID  = '' THEN VENDINVOICEJOUR.INVOICEACCOUNT ELSE HOYA_VENDERID END  [Vender]");

            sbSql.AppendLine("  FROM VENDINVOICEJOUR  ");
            sbSql.AppendLine("INNER JOIN VENDINVOICETRANS ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE = VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine("AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP = VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID = VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");

            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=VENDINVOICETRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            
            //sbSql.AppendLine("INNER JOIN INVENTDIM ON VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
           // sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            
            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTABLE.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");

            sbSql.AppendLine("INNER JOIN VENDTABLE On VENDTABLE.ACCOUNTNUM = VENDINVOICEJOUR.INVOICEACCOUNT");
            sbSql.AppendLine("AND VENDTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");

            sbSql.AppendLine("WHERE");
          //  sbSql.AppendLine("INVENTDIM.INVENTSITEID = '" + MaterialOBJ.strFactory + "'");



            if (MaterialOBJ.Category == "All")
            {


                if (MaterialOBJ.Factory == "PO")
                {
                   // sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ( 'PO-DM', 'PO-IM', 'PO-MT', 'PO-CND', 'PO-CNI', 'PO-CNM' ) ");
                    sbSql.AppendLine(" VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ( 'PO-DM', 'PO-IM', 'PO-MT', 'PO-CND', 'PO-CNI', 'PO-CNM' ) ");
                    sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
                }
                else
                {
                   // sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
                     sbSql.AppendLine("  INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
                    sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ('MO-DM', 'MO-IM', 'MO-MT', 'MO-CND', 'MO-CNI', 'MO-CNM') ");
                }

            }
            else
            {
                sbSql.AppendLine("  INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
                if (MaterialOBJ.Factory == "PO")
                {
                    sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ( 'PO-DM', 'PO-IM', 'PO-MT', 'PO-CND', 'PO-CNI', 'PO-CNM' ) ");
                    // sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
                }
                else
                {
                    // sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
                    sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ('MO-DM', 'MO-IM', 'MO-MT', 'MO-CND', 'MO-CNI', 'MO-CNM') ");

                }
            }


            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

            sbSql.AppendLine("GROUP BY HOYA_VENDERID ,VENDINVOICEJOUR.INVOICEACCOUNT ) as VenderID");
            sbSql.AppendLine("GROUP BY VenderID.Vender");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }


        public ADODB.Recordset getMaterailPurchaseSubGroup(MaterialOBJ MaterialOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            DateTime dtFrom = new DateTime(MaterialOBJ.DateFrom.Year, MaterialOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(MaterialOBJ.DateTo.Year, MaterialOBJ.DateTo.Month, 1);

            sbSql.AppendLine("SELECT [SUBGROUP],'PCS'");
            sbSql.AppendLine("FROM (SELECT");

            //sbSql.AppendLine("CASE WHEN HOYA_VENDERID  = '' THEN VENDINVOICEJOUR.INVOICEACCOUNT ELSE HOYA_VENDERID END  [Vender]");
            sbSql.AppendLine("INVENTTABLE.ECL_SUBGROUP [SUBGROUP]");
            sbSql.AppendLine("  FROM VENDINVOICEJOUR  ");
            sbSql.AppendLine("INNER JOIN VENDINVOICETRANS ON VENDINVOICETRANS.PURCHID = VENDINVOICEJOUR.PURCHID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEID = VENDINVOICEJOUR.INVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.INVOICEDATE = VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine("AND VENDINVOICETRANS.NUMBERSEQUENCEGROUP = VENDINVOICEJOUR.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("AND VENDINVOICETRANS.INTERNALINVOICEID = VENDINVOICEJOUR.INTERNALINVOICEID");
            sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");

            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=VENDINVOICETRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=VENDINVOICETRANS.DATAAREAID");

            //sbSql.AppendLine("INNER JOIN INVENTDIM ON VENDINVOICETRANS.INVENTDIMID=INVENTDIM.INVENTDIMID");
            // sbSql.AppendLine("AND VENDINVOICETRANS.DATAAREAID=INVENTDIM.DATAAREAID");

            sbSql.AppendLine("INNER JOIN INVENTITEMGROUPITEM ON INVENTTABLE.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");

            sbSql.AppendLine("INNER JOIN VENDTABLE On VENDTABLE.ACCOUNTNUM = VENDINVOICEJOUR.INVOICEACCOUNT");
            sbSql.AppendLine("AND VENDTABLE.DATAAREAID = VENDINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");

            sbSql.AppendLine("WHERE");
            //  sbSql.AppendLine("INVENTDIM.INVENTSITEID = '" + MaterialOBJ.strFactory + "'");



            if (MaterialOBJ.Category == "All")
            {


                if (MaterialOBJ.Factory == "PO")
                {
                    // sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ( 'PO-DM', 'PO-IM', 'PO-MT', 'PO-CND', 'PO-CNI', 'PO-CNM' ) ");
                    sbSql.AppendLine(" VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ( 'PO-DM', 'PO-IM', 'PO-MT', 'PO-CND', 'PO-CNI', 'PO-CNM' ) ");
                    sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
                }
                else
                {
                    // sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
                    sbSql.AppendLine("  INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
                    sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ('MO-DM', 'MO-IM', 'MO-MT', 'MO-CND', 'MO-CNI', 'MO-CNM') ");
                }

            }
            else
            {
                sbSql.AppendLine("  INVENTITEMGROUPITEM.ITEMGROUPID='" + MaterialOBJ.Category + "'");
                if (MaterialOBJ.Factory == "PO")
                {
                    sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ( 'PO-DM', 'PO-IM', 'PO-MT', 'PO-CND', 'PO-CNI', 'PO-CNM' ) ");
                    // sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y')");
                }
                else
                {
                    // sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('Z','Y','O','I')");
                    sbSql.AppendLine("AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP IN ('MO-DM', 'MO-IM', 'MO-MT', 'MO-CND', 'MO-CNI', 'MO-CNM') ");

                }
            }


            sbSql.AppendLine(" AND VENDINVOICEJOUR.INVOICEDATE  BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", MaterialOBJ.DateTo) + "',103)");

            sbSql.AppendLine("GROUP BY INVENTTABLE.ECL_SUBGROUP) as VenderID");
            //sbSql.AppendLine("GROUP BY VenderID.Vender");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

    }//end class




}
