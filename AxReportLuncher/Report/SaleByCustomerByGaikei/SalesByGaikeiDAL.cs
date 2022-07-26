using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NewVersion.Report.SaleByCustomerByGaikei
{
    class SalesByGaikeiDAL
    {
        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();


        public ADODB.Recordset getSaleByCustomer(SalesByGaikeiOBJ SalesByGaikeiOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            sbSql.AppendLine("SELECT ");
            sbSql.AppendLine("CASE WHEN tb_Sales.ITEMCD  IS NULL AND tb_Sales.TRADEPART IS NULL THEN 'GRAND TOTAL' ELSE  ");
            sbSql.AppendLine("CASE WHEN  tb_Sales.ITEMCD IS NULL THEN 'TOTAL' ELSE tb_Sales.ITEMCD END END [ItemCD]");

            sbSql.AppendLine(",CASE WHEN TRADEPART = '' THEN tb_Sales.CUSTNAME ELSE TRADEPART END [Customer]");
            sbSql.AppendLine(",tb_Sales.GLASSTYPE ,tb_Sales.GAIKEI,tb_Sales.PWT,SUM(OUT_PCS)PCS,SUM(OUT_KGS)KGS");
            sbSql.AppendLine(",CURRENCY,UNITCUR [UNit Price],SUM(SALESCUR) [Fur Cur.],SUM(SALESBAHT) BAHT ");
            sbSql.AppendLine("from tb_Sales INNER JOIN tb_SP ON tb_SP.ITEMCD = tb_Sales.ITEMCD");
            sbSql.AppendLine("AND tb_Sales.LOTNO = tb_SP.LOTNO AND tb_Sales.YEARMONTH = tb_SP.YEARMONTH");



           // sbSql.AppendLine(" WHERE  tb_Sales.FACTORYCD ='" + SalesByGaikeiOBJ.Factory + "'");
            sbSql.AppendLine("WHERE tb_Sales.YEARMONTH  = '"+SalesByGaikeiOBJ.dtDate+"'");

            //sbSql.AppendLine("GROUP BY TRADEPART,tb_Sales.CUSTNAME,tb_Sales.ITEMCD,tb_Sales.GLASSTYPE,tb_Sales.GAIKEI,tb_Sales.PWT,CURRENCY,UNITCUR WITH ROLLUP");
            //sbSql.AppendLine("HAVING NOT tb_Sales.UNITCUR IS NULL OR tb_Sales.CUSTNAME IS NULL");
            sbSql.AppendLine("GROUP BY TRADEPART,tb_Sales.ITEMCD,tb_Sales.CUSTNAME,tb_Sales.GLASSTYPE,tb_SP.SOZAIDIV,tb_Sales.GAIKEI,tb_Sales.PWT,CURRENCY,UNITCUR WITH ROLLUP");
            sbSql.AppendLine("HAVING NOT tb_Sales.UNITCUR IS NULL OR tb_Sales.ITEMCD IS NULL");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.RPContingConnect();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;


        }



        public ADODB.Recordset getSaleByCustomerByGaikei(SalesByGaikeiOBJ SalesByGaikeiOBJ,String Gaikei,String Fac)
        {
            StringBuilder sbSql = new StringBuilder();


            sbSql.AppendLine("SELECT ");
            sbSql.AppendLine("tb_Sales.ITEMCD ");
            sbSql.AppendLine(",CASE WHEN TRADEPART = '' THEN tb_Sales.CUSTNAME ELSE TRADEPART END [Customer]");
            sbSql.AppendLine(",tb_Sales.GLASSTYPE ,tb_SP.SOZAIDIV,tb_Sales.GAIKEI,tb_Sales.PWT,SUM(OUT_PCS)PCS,SUM(OUT_KGS)KGS");
            sbSql.AppendLine(",CURRENCY,UNITCUR [UNit Price],SUM(SALESCUR) [Fur Cur.],SUM(SALESBAHT) BAHT ");
            sbSql.AppendLine("from tb_Sales INNER JOIN tb_SP ON tb_SP.ITEMCD = tb_Sales.ITEMCD");
            sbSql.AppendLine("AND tb_Sales.LOTNO = tb_SP.LOTNO AND tb_Sales.YEARMONTH = tb_SP.YEARMONTH");


            if (Fac != "All")
            {

                sbSql.AppendLine(" WHERE  tb_Sales.FACTORYCD ='" + Fac + "'");
                sbSql.AppendLine("And tb_Sales.YEARMONTH  = '" + SalesByGaikeiOBJ.dtDate + "'");
            }
            else
            {
               // sbSql.AppendLine(" WHERE  tb_Sales.FACTORYCD ='" + SalesByGaikeiOBJ.Factory + "'");
                sbSql.AppendLine("WHERE tb_Sales.YEARMONTH  = '" + SalesByGaikeiOBJ.dtDate + "'");
            }

           
         
            sbSql.AppendLine("AND " + Gaikei);

            sbSql.AppendLine("GROUP BY TRADEPART,tb_Sales.ITEMCD,tb_Sales.CUSTNAME,tb_Sales.GLASSTYPE,tb_SP.SOZAIDIV,tb_Sales.GAIKEI,tb_Sales.PWT,CURRENCY,UNITCUR ");
           
            
            
            // sbSql.AppendLine("HAVING NOT tb_Sales.UNITCUR IS NULL OR tb_Sales.CUSTNAME IS NULL");

           
         
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.RPContingConnect();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;


        }



        public ADODB.Recordset getSaleByCustomerByGaikei2(SalesByGaikeiOBJ SalesByGaikeiOBJ, String Gaikei)
        {
            StringBuilder sbSql = new StringBuilder();


            sbSql.AppendLine("SELECT ");
            sbSql.AppendLine("tb_Sales.ITEMCD ");
            sbSql.AppendLine(",CASE WHEN TRADEPART = '' THEN tb_Sales.CUSTNAME ELSE TRADEPART END [Customer]");
            sbSql.AppendLine(",tb_Sales.GLASSTYPE ,tb_Sales.GAIKEI,tb_Sales.PWT,SUM(OUT_PCS)PCS,SUM(OUT_KGS)KGS");
            sbSql.AppendLine(",CURRENCY,UNITCUR [UNit Price],SUM(SALESCUR) [Fur Cur.],SUM(SALESBAHT) BAHT ");
            sbSql.AppendLine("from tb_Sales INNER JOIN tb_SP ON tb_SP.ITEMCD = tb_Sales.ITEMCD");
            sbSql.AppendLine("AND tb_Sales.LOTNO = tb_SP.LOTNO AND tb_Sales.YEARMONTH = tb_SP.YEARMONTH");


            sbSql.AppendLine("WHERE tb_Sales.YEARMONTH  = '" + SalesByGaikeiOBJ.dtDate + "'");
            



            sbSql.AppendLine("AND " + Gaikei);

            //sbSql.AppendLine("GROUP BY TRADEPART,tb_Sales.CUSTNAME,tb_Sales.ITEMCD,tb_Sales.GLASSTYPE,tb_Sales.GAIKEI,tb_Sales.PWT,CURRENCY,UNITCUR ");
            // sbSql.AppendLine("HAVING NOT tb_Sales.UNITCUR IS NULL OR tb_Sales.CUSTNAME IS NULL");
            sbSql.AppendLine("GROUP BY TRADEPART,tb_Sales.ITEMCD,tb_Sales.CUSTNAME,tb_Sales.GLASSTYPE,tb_SP.SOZAIDIV,tb_Sales.GAIKEI,tb_Sales.PWT,CURRENCY,UNITCUR ");
           
      


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.RPContingConnect();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;


        }






    }
}
