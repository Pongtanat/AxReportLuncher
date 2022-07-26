using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.StockCard
{
    class StockCardDAL
    {

        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();


        public ADODB.Recordset getStockCard(StockCardOBJ StockCardOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            var firstDayBeforeMonth = new DateTime(StockCardOBJ.DateFrom.AddMonths(-1).Year, StockCardOBJ.DateFrom.AddMonths(-1).Month, 1);
            var lastDayOfBeforeMonth = firstDayBeforeMonth.AddMonths(1).AddDays(-1);

            sbSql.AppendLine("SELECT DISTINCT * FROM(");
            sbSql.AppendLine("SELECT ");
            sbSql.AppendLine("Balance.[ItemID] ");
            sbSql.AppendLine(",Balance.[NAME]");
            sbSql.AppendLine(",Balance.[LOCATIONID]");
            sbSql.AppendLine(",Begining.BOMQTY  [BOMQTY]");
            sbSql.AppendLine(",Begining.BOMCOST  [BOMCOST] ");
            sbSql.AppendLine(",Received.[Received QTY]");
            sbSql.AppendLine(",Received.[Received Cost]");
            sbSql.AppendLine(",Issue.[Issue QTY]");
            sbSql.AppendLine(",Issue.[Issue Cost]");
            sbSql.AppendLine(",Balance.QTY [Balance QTY]");
            sbSql.AppendLine(",Balance.Cost [Balance COST]");
            sbSql.AppendLine(",Balance.[Unit Cost]");


            //============= Balance =================//
            sbSql.AppendLine(" FROM(SELECT");
            sbSql.AppendLine("InventTrans.ITEMID [ITEMID]");
            sbSql.AppendLine(",ECORESPRODUCTTRANSLATION.[NAME] ");
            sbSql.AppendLine(",INVENTLOCATION.INVENTLOCATIONID [LOCATIONID]");
    
            sbSql.AppendLine(",INVENTTABLE.HOYA_LENSTYPE");
            sbSql.AppendLine(",SUM(InventTrans.QTY) [QTY]");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT) [Cost]  ");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)/NULLIF(SUM(InventTrans.QTY),0)[Unit Cost]");
            sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("INNER JOIN inventlocation ON inventlocation.inventlocationid = inventdim.inventlocationid");
            sbSql.AppendLine("INNER JOIN inventsite on inventsite.siteid = inventdim.inventsiteid");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");

            sbSql.AppendLine("WHERE");        
            sbSql.AppendLine("  INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'1/01/1991',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", StockCardOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='"+StockCardOBJ.Factory+"'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");
            sbSql.AppendLine("AND INVENTTABLE.ITEMTYPE != '2'");

            if (StockCardOBJ.GroupID != "")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('" + StockCardOBJ.GroupID + "')");
            }
    
            sbSql.AppendLine("  GROUP BY InventTrans.ITEMID,ECORESPRODUCTTRANSLATION.NAME,INVENTLOCATION.INVENTLOCATIONID,INVENTTABLE.HOYA_LENSTYPE");
            sbSql.AppendLine(" )Balance");
            sbSql.AppendLine("INNER JOIN INVENTITEMINVENTSETUP on INVENTITEMINVENTSETUP.ITEMID = Balance.ITEMID");


            //======================== Received ====================//
            sbSql.AppendLine("LEFT JOIN");
            sbSql.AppendLine("(SELECT");
            sbSql.AppendLine(" INVENTTRANS.ITEMID [ITEMID]");
            sbSql.AppendLine(",SUM(InventTrans.QTY) [Received QTY]");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)  [Received Cost]");
            sbSql.AppendLine("from INVENTTRANS INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("INNER JOIN inventlocation ON inventlocation.inventlocationid = inventdim.inventlocationid");
            sbSql.AppendLine("INNER JOIN inventsite on inventsite.siteid = inventdim.inventsiteid");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");

            sbSql.AppendLine("WHERE");
            sbSql.AppendLine("  INVENTTRANS.DATEFINANCIAL between  CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", StockCardOBJ.DateFrom) + "',103)");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", StockCardOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + StockCardOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");
            sbSql.AppendLine("AND STATUSRECEIPT = '1'");
            sbSql.AppendLine("AND INVENTTABLE.ITEMTYPE != '2'");
            sbSql.AppendLine("GROUP BY INVENTTRANS.ITEMID ");
            sbSql.AppendLine(")Received ON Balance.ITEMID = Received.ITEMID");


            //================== Issue ====================//
            sbSql.AppendLine("LEFT JOIN");
            sbSql.AppendLine("(SELECT ");
            sbSql.AppendLine(" INVENTTRANS.ITEMID [ITEMID]");
            sbSql.AppendLine(",SUM(InventTrans.QTY)*-1 [Issue QTY]");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)  *-1[Issue Cost]");
            sbSql.AppendLine("from INVENTTRANS INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID = INVENTTABLE.DATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("INNER JOIN inventlocation ON inventlocation.inventlocationid = inventdim.inventlocationid");
            sbSql.AppendLine("INNER JOIN inventsite on inventsite.siteid = inventdim.inventsiteid");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("WHERE");
            sbSql.AppendLine("  INVENTTRANS.DATEFINANCIAL between  CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", StockCardOBJ.DateFrom) + "',103)");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", StockCardOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + StockCardOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");
            sbSql.AppendLine("AND STATUSISSUE = '1'");
            sbSql.AppendLine("AND INVENTTABLE.ITEMTYPE != '2'");
            sbSql.AppendLine("GROUP BY INVENTTRANS.ITEMID ");
            sbSql.AppendLine(")Issue ON Balance.ITEMID = Issue.ITEMID");


            //========================= BOM=========================//
            sbSql.AppendLine(" LEFT JOIN ");
            sbSql.AppendLine(" (SELECT ");
            sbSql.AppendLine(" INVENTTRANS.ITEMID [ITEMID]");
            sbSql.AppendLine(",SUM(InventTrans.QTY) [BOMQTY]");
            sbSql.AppendLine(",SUM(INVENTTRANS.COSTAMOUNTPOSTED+INVENTTRANS.COSTAMOUNTADJUSTMENT)  [BOMCost]");
            sbSql.AppendLine("FROM InventTrans INNER JOIN INVENTTABLE on INVENTTABLE.ITEMID = INVENTTRANS.ITEMID");
            sbSql.AppendLine(" INNER JOIN INVENTITEMGROUPITEM ON INVENTTRANS.ITEMID=INVENTITEMGROUPITEM.ITEMID");
            sbSql.AppendLine("AND INVENTTRANS.DATAAREAID=INVENTITEMGROUPITEM.ITEMDATAAREAID");
            sbSql.AppendLine("INNER JOIN INVENTDIM ON INVENTTRANS.INVENTDIMID = INVENTDIM.INVENTDIMID");
            sbSql.AppendLine("INNER JOIN inventlocation ON inventlocation.inventlocationid = inventdim.inventlocationid");
            sbSql.AppendLine("INNER JOIN inventsite on inventsite.siteid = inventdim.inventsiteid");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATION ON INVENTTABLE.PRODUCT=ECORESPRODUCTTRANSLATION.PRODUCT");
            sbSql.AppendLine("WHERE");

            sbSql.AppendLine("  INVENTTRANS.DATEFINANCIAL between CONVERT(datetime,'1/01/1991',103) ");
            sbSql.AppendLine("  AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", lastDayOfBeforeMonth) + "',103)");
            sbSql.AppendLine(" AND INVENTDIM.INVENTSITEID='" + StockCardOBJ.Factory + "'");
            sbSql.AppendLine("AND INVENTTABLE.DATAAREAID = 'hoya'");

            if (StockCardOBJ.GroupID != "")
            {
                sbSql.AppendLine(" AND INVENTITEMGROUPITEM.ITEMGROUPID IN ('" + StockCardOBJ.GroupID + "')");
            }
            sbSql.AppendLine("AND INVENTTABLE.ITEMTYPE != '2'");
            sbSql.AppendLine(" GROUP BY INVENTTRANS.ITEMID ");
            sbSql.AppendLine(" )Begining ON Balance.ITEMID = Begining.ITEMID");
            //sbSql.AppendLine("WHERE INVENTITEMINVENTSETUP.STOPPED = 0");
           
            sbSql.AppendLine(")as Total");

            if (StockCardOBJ.ItemID != "")
            {
                sbSql.AppendLine("WHERE Total.ITEMID = '" + StockCardOBJ.ItemID + "'");
            }

          

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

    }

}
