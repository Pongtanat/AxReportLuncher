using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.Class.DAL
{
 public  class FindDAL
    {
        SQLConnectionDAL ConnectionDAL = new SQLConnectionDAL();


       public DataTable findVender(string strSearchField,string strSearchValue){
           StringBuilder sbSql = new StringBuilder();
           sbSql.AppendLine(" SELECT AccountNum,Name FROM VendTable");
           sbSql.AppendLine(" INNER JOIN VendDirPartyTableView ON VendTable.Party=VendDirPartyTableView.Party");

           if (strSearchField != "")
           {
               sbSql.AppendLine("WHERE" + strSearchField + "LIKE '%" + strSearchValue + "%'");
           }
           sbSql.AppendLine(" ORDER BY AccountNum");
          DataTable dt = ConnectionDAL.QueryDataTable(sbSql.ToString());
          return dt;

        }



       public DataTable findCustomer(string strSearchField, string strSearchValue)
       {
           StringBuilder sbSql = new StringBuilder();
         sbSql.AppendLine(" SELECT AccountNum,Name FROM CUSTTABLE");
         sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON CUSTTABLE.Party=DIRPARTYTABLE.RECID");
         
           if (strSearchField != "")
           {
               sbSql.AppendLine("WHERE" + strSearchField + "LIKE '%" + strSearchValue + "%'");
           }
           sbSql.AppendLine(" ORDER BY AccountNum");
           DataTable dt = ConnectionDAL.QueryDataTable(sbSql.ToString());
           return dt;

       }


    }
}
