using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;

namespace NewVersion.Material
{
    class MaterialDAL
    {
        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();

      

    public DataTable getCustomerGroup(){
        StringBuilder sbSql = new StringBuilder();
        sbSql.AppendLine(" SELECT CustGroup,Name FROM CUSTGROUP");
        sbSql.AppendLine(" ORDER BY CustGroup");
        DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
        return dt;
    }

    public DataTable getDimAttrValSetItemAX(string D1,string D2, string D3, string D4)
    {
            StringBuilder sbSql = new StringBuilder ();

        try{

           sbSql.AppendLine(" SELECT * FROM ");
            sbSql.AppendLine("  (SELECT DIMENSIONATTRIBUTEVALUESETITEM.* FROM DIMENSIONATTRIBUTE ");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUE ");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTE.RECID = DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTEVALUE.RECID = DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("  WHERE NAME = 'D1_Factory') FACTORY");
            sbSql.AppendLine(" LEFT OUTER JOIN");
            sbSql.AppendLine("  (SELECT DIMENSIONATTRIBUTEVALUESETITEM.* FROM DIMENSIONATTRIBUTE ");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUE ");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTE.RECID = DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTEVALUE.RECID = DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("  WHERE NAME = 'D2_Section') SECTION ON FACTORY.DIMENSIONATTRIBUTEVALUESET = SECTION.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" LEFT OUTER JOIN");

            sbSql.AppendLine("  (SELECT DIMENSIONATTRIBUTEVALUESETITEM.* FROM DIMENSIONATTRIBUTE ");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUE ");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTE.RECID = DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTEVALUE.RECID = DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("  WHERE NAME = 'D3_SubSection') SUBSECTION ON FACTORY.DIMENSIONATTRIBUTEVALUESET = SUBSECTION.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" LEFT OUTER JOIN");

            sbSql.AppendLine("  (SELECT DIMENSIONATTRIBUTEVALUESETITEM.* FROM DIMENSIONATTRIBUTE ");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUE ");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTE.RECID = DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTEVALUE.RECID = DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("  WHERE NAME = 'D4_Related') RELATED ON FACTORY.DIMENSIONATTRIBUTEVALUESET = RELATED.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" LEFT OUTER JOIN");
            sbSql.AppendLine("  (SELECT DIMENSIONATTRIBUTEVALUESETITEM.* FROM DIMENSIONATTRIBUTE ");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUE ");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTE.RECID = DIMENSIONATTRIBUTEVALUE.DIMENSIONATTRIBUTE");
            sbSql.AppendLine("      INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM");
            sbSql.AppendLine("          ON DIMENSIONATTRIBUTEVALUE.RECID = DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("  WHERE NAME = 'D5_Project') PROJECT ON FACTORY.DIMENSIONATTRIBUTEVALUESET = PROJECT.DIMENSIONATTRIBUTEVALUESET");
            sbSql.AppendLine(" WHERE FACTORY.DISPLAYVALUE ='" + D1 + "'");
            sbSql.AppendLine("  AND SECTION.DISPLAYVALUE = '" + D2 + "' ");
            sbSql.AppendLine("  AND SUBSECTION.DISPLAYVALUE = '" + D3 + "'");
            sbSql.AppendLine("  AND RELATED.DISPLAYVALUE = '" + D4 + "'") ;
            sbSql.AppendLine("  AND PROJECT.DISPLAYVALUE = '9999999999'");





            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;

       }catch (Exception ex){
           MessageBox.Show(ex.Message);
           return null;
       }

    }//end getDimAttrValSetItemAX

       public DataTable getDimAttrValueLine(string ItemCD, string GlassType, string SozaiDiv)
       {
            StringBuilder sbSql = new StringBuilder ();

        try{

          sbSql.AppendLine(" SELECT top 1 INVENTTABLE.ITEMID,ECL_SubGroup  FROM  INVENTTABLE ");
            sbSql.AppendLine("INNER JOIN INVENTTRANS ON INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");

            if (ItemCD != "")
            {
                sbSql.AppendLine("WHERE HOYA_PRODUCTIONITEM = '" + ItemCD + "'");
                sbSql.AppendLine("AND HOYA_GLASSTYPE = '" + GlassType + "'");
                sbSql.AppendLine("AND HOYA_Sozaidiv = '" + SozaiDiv + "'");
            }
            else if (GlassType == "" && SozaiDiv == "")
            {
                sbSql.AppendLine("WHERE HOYA_PRODUCTIONITEM = '" + ItemCD + "'");
            }
            else
            {
                sbSql.AppendLine("WHERE HOYA_GLASSTYPE = '" + GlassType + "'");
                sbSql.AppendLine("AND HOYA_Sozaidiv = '" + SozaiDiv + "'");
                sbSql.AppendLine("AND ECL_SUBGROUP='EB'");
            }

            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;

       }catch (Exception ex){
           MessageBox.Show(ex.Message);
           return null;
       }

  }//end getDimAttrValueLine

       public DataTable getCost(string ITEMID)
       {
           StringBuilder sbSql = new StringBuilder();
           try
           {

            sbSql.AppendLine(" SELECT top 1((INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)*-1 / INVENTTRANS.QTY) *-1 [Cost]");
            sbSql.AppendLine(" FROM INVENTTRANS");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSPOSTING ON INVENTTRANS.VOUCHER=INVENTTRANSPOSTING.VOUCHER");
            sbSql.AppendLine("AND INVENTTRANS.DATEFINANCIAL=INVENTTRANSPOSTING.TRANSDATE");
            sbSql.AppendLine(" AND INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSPOSTING.INVENTTRANSORIGIN");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTTRANSPOSTING.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID=INVENTDIM.INVENTDIMID AND INVENTTRANS.DATAAREAID=INVENTDIM.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID=INVENTTRANS.ITEMID AND INVENTTABLE.DATAAREAID=INVENTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine(" INNER JOIN INVENTTABLEMODULE ON INVENTTRANS.ITEMID=INVENTTABLEMODULE.ITEMID AND INVENTTRANS.DATAAREAID=INVENTTABLEMODULE.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTTRANSORIGIN ON INVENTTRANS.INVENTTRANSORIGIN=INVENTTRANSORIGIN.RECID AND INVENTTRANS.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" LEFT OUTER JOIN INVENTJOURNALTABLE ON INVENTJOURNALTABLE.JOURNALID=INVENTTRANSORIGIN.REFERENCEID AND INVENTJOURNALTABLE.DATAAREAID=INVENTTRANSORIGIN.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMINVENTSETUP ON INVENTITEMINVENTSETUP.ITEMID = INVENTTABLE.ITEMID");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.DATAAREAID = INVENTTABLE.DATAAREAID");

            sbSql.AppendLine("WHERE  INVENTTABLE.ITEMID = '" + ITEMID + "'");

            sbSql.AppendLine(" AND  INVENTTRANS.DATAAREAID='hoya'");
            sbSql.AppendLine("AND INVENTTRANSPOSTING.INVENTTRANSPOSTINGTYPE='1'");
            sbSql.AppendLine(" AND INVENTTABLEMODULE.MODULETYPE='0'");
            sbSql.AppendLine(" AND STATUSISSUE = 1");
            sbSql.AppendLine(" AND INVENTTRANSORIGIN.REFERENCECATEGORY = 4");
            sbSql.AppendLine("AND INVENTITEMINVENTSETUP.STOPPED = '0'");
            sbSql.AppendLine("AND ((INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)*-1 / INVENTTRANS.QTY) *-1 >0");
            sbSql.AppendLine("order by inventtrans.DATEPHYSICAL desc");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
               return null;
           }

       }//end gotCost

       public DataTable getAddressByLocation(string strSite,string strLocation)
       {
           StringBuilder sbSql = new StringBuilder();
           try
           {

            sbSql.AppendLine(" SELECT INVENTDIM.INVENTDIMID");
            sbSql.AppendLine(" ,INVENTLOCATION.ECL_BRANCHID");
            sbSql.AppendLine(" ,INVENTLOCATION.ECL_MANAGER");
            sbSql.AppendLine(" ,LOGISTICSLOCATION.[DESCRIPTION]");
            sbSql.AppendLine(" ,LOGISTICSPOSTALADDRESS.RECID LogisticsPostalAddress");
            sbSql.AppendLine(" FROM InventLocationLogisticsLocation");
            sbSql.AppendLine(" INNER JOIN InventLocationLogisticsLocationRole ON InventLocationLogisticsLocation.RecId = InventLocationLogisticsLocationRole.LocationLogisticsLocation");
            sbSql.AppendLine(" INNER JOIN LogisticsLocationRole ON InventLocationLogisticsLocationRole.LocationRole = LogisticsLocationRole.RecId");
            sbSql.AppendLine(" INNER JOIN INVENTLOCATION ON InventLocationLogisticsLocation.InventLocation = INVENTLOCATION.recid");
            sbSql.AppendLine(" INNER JOIN INVENTDIM ON INVENTDIM.INVENTLOCATIONID=INVENTLOCATION.INVENTLOCATIONID");
            sbSql.AppendLine(" INNER JOIN LOGISTICSLOCATION ON LOGISTICSLOCATION.RECID=INVENTLOCATIONLOGISTICSLOCATION.LOCATION");
            sbSql.AppendLine(" INNER JOIN LOGISTICSPOSTALADDRESS on LOGISTICSLOCATION.RECID=LOGISTICSPOSTALADDRESS.LOCATION");
            sbSql.AppendLine(" WHERE LogisticsLocationRole.NAME='English Address'");
            sbSql.AppendLine(" AND LOGISTICSPOSTALADDRESS.VALIDFROM<=GETDATE() AND LOGISTICSPOSTALADDRESS.VALIDTO>GETDATE()");
            sbSql.AppendLine(" AND INVENTDIM.INVENTLOCATIONID='" + strSite + "'");
            sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + strLocation + "'");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
               return null;
           }

       }//end getAddressByLocation



    }//end Class
}
