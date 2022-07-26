using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.CompareBudomari
{
    class BudomariDAL
    {

        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();
        public DataTable getCustomer(string strSearchField, string strSearchValue)
        {
            StringBuilder sbSql = new StringBuilder();

            sbSql.AppendLine(" SELECT AccountNum,Name FROM CUSTTABLE");
            sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON CUSTTABLE.Party=DIRPARTYTABLE.RECID");
            sbSql.AppendLine(" WHERE " + strSearchField + " LIKE '%" + strSearchValue + "%'");
            sbSql.AppendLine(" ORDER BY AccountNum");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }



        public ADODB.Recordset GetBudomari(string strGroup, BudomariOBJ BudomariOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            sbSql.AppendLine(" SELECT");
            sbSql.AppendLine(" G.[CONDITIONJ] [CONDITION],");
            sbSql.AppendLine(" CASE WHEN Budomari.LITEM IS NULL THEN Budomari.TITEM ELSE Budomari.LITEM END [ITEM]");
            sbSql.AppendLine(",Budomari.LBOM [LASTBOM]");

            sbSql.AppendLine(",Budomari.LIN [LAST IN WIP],''[LAST RT WH],''[LAST NET IN WIP]");
            sbSql.AppendLine(",Budomari.LSHIPPING [LAST SHIPPING IN],'' [LAST SH OUR RT],''[LAST SH out NET]");

            sbSql.AppendLine(",Budomari.[LDIFF][LAST NGDIFF],Budomari.LNGTESTMC [LAST NG TEST MC],Budomari.LREJECTSPECIAL [LAST REJECT SPECIAL]");
            sbSql.AppendLine(",Budomari.LSALELENINWIP [LAST SALELENS IN WIP],Budomari.LDS [LAST DS],Budomari.LTINREROUTE [LAST INREROUTE]");
            sbSql.AppendLine(",'' [LAST NET NG],Budomari.TBOM [LAST EOM]");

            sbSql.AppendLine(",'' [LAST Budomari],'',Budomari.TBOM [THIS BOM]");
            sbSql.AppendLine(",Budomari.TINWIP [THIS IN WIP],''[THIS RT WH] ,''[THIS NET IN WIP],Budomari.TSHIPPING [THIS SHIPPIN IN]");
            sbSql.AppendLine(",'' [THIS SH OUT RT],'' [THIS SH IN NET],Budomari.TDIFF [THIS NG+DIFF],Budomari.TNGTESTMC [THIS NG TEST MC]");

            sbSql.AppendLine(",Budomari.TREJECTSPECIAL [THIS REJECTSPECIAL],Budomari.TSALELENINWIP [THIS SALES IN WIP]");
            sbSql.AppendLine(",Budomari.TDS [THIS DS],Budomari.TTINREROUTE [THIS IN ROUTE],''[THIS NET NG]");
            sbSql.AppendLine(",Budomari.TEOM [THIS EOM],'' [THIS Budomari]");

            sbSql.AppendLine("FROM(SELECT ");
            sbSql.AppendLine("CASE WHEN L.CONDITION  IS NULL THEN T.CONDITION  ELSE L.CONDITION   END [CONDITION] ");
            sbSql.AppendLine(",L.[ITEM NO] [LITEM],L.CONDITION1 [LCONDITION1],L.BOM [LBOM]");

            sbSql.AppendLine(",L.[IN] [LIN],L.[SHIPPING IN] [LSHIPPING],L.[NG+DIFF] [LDIFF],L.[NG TEST MC] [LNGTESTMC]");
            sbSql.AppendLine(",L.[REJECT SPECIAL] [LREJECTSPECIAL],L.[SALE LENS IN WIP][LSALELENINWIP],L.[DS][LDS]");
            sbSql.AppendLine(",L.[IN-REROUTE][LTINREROUTE],L.[NET NG][LNETNG],L.[EOM][LEOM],T.CONDITION [TCONDITION]");

            sbSql.AppendLine(",T.[ITEM NO] [TITEM],T.BOM [TBOM],T.[IN WIP] [TINWIP],T.[SHIPPING IN] [TSHIPPING],T.[NG+DIFF] [TDIFF]");
            sbSql.AppendLine(",T.[NG TEST MC] [TNGTESTMC],T.[REJECT SPECIAL] [TREJECTSPECIAL],T.[SALE LENS IN WIP][TSALELENINWIP]");
            sbSql.AppendLine(",T.[DS][TDS],T.[IN-REROUTE][TTINREROUTE],T.[NET NG][TNETNG],T.[EOM][TEOM]");

            sbSql.AppendLine(@" FROM((SELECT  * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 8.0; Database=E:\CompareBudomari\Budomari.xlsx', 'SELECT * FROM [" + BudomariOBJ.GetSheet1 + "$]')) L");
            sbSql.AppendLine("FULL JOIN ");
            sbSql.AppendLine(@"(SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 8.0; Database=E:\CompareBudomari\Budomari.xlsx', 'SELECT * FROM [" + BudomariOBJ.GetSheet2 + "$]')) T");
            //sbSql.AppendLine("ON L.CONDITION  = T.CONDITION AND L.[ITEM NO] = T.[ITEM NO] )) AS Budomari");
            sbSql.AppendLine("ON  L.[ITEM NO] = T.[ITEM NO])) AS Budomari");
            sbSql.AppendLine("INNER JOIN ");
            sbSql.AppendLine(@"(SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 8.0; Database=E:\CompareBudomari\MasterGroup.xlsx', 'SELECT * FROM [MasterGroup$]')) G");
            //sbSql.AppendLine("ON G.[CONDITION (EN)] = Budomari.LCONDITION");
            sbSql.AppendLine("ON G.[CONDITIONE] = Budomari.[CONDITION]");
            sbSql.AppendLine("WHERE  Budomari.[CONDITION]  = '" + strGroup + "'");
         
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }


        public DataTable   GetMasterGroup()
        {
            StringBuilder sbSql = new StringBuilder();


            sbSql.AppendLine(@"SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 8.0; Database=E:\CompareBudomari\MasterGroup.xlsx', 'SELECT * FROM [MasterGroup$]') G");


            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;

        }

    
 




    }

}
