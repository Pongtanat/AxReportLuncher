using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace NewVersion
{
    
	public class AXReportLuncherDAL
	{
        SQLConnectionDAL ConnectionDAL = new SQLConnectionDAL();
        
        public DataTable getAllMenu(){

            StringBuilder sbSql = new StringBuilder();
             sbSql.AppendLine(" SELECT NOTE FROM HcmWorkerTask");
             sbSql.AppendLine(" WHERE WORKERTASKID LIKE 'Report %'");
             sbSql.AppendLine(" AND NOTE<>''");
             sbSql.AppendLine(" ORDER BY WorkerTaskID");

             DataTable dt = new DataTable();
             dt = ConnectionDAL.QueryDataTable(sbSql.ToString());
             return dt;
        }

        public DataTable getRoleuser(string strUser){

        StringBuilder sbSql = new StringBuilder();
        sbSql.AppendLine(" SELECT KNOWNAS,OFFICELOCATION");
        sbSql.AppendLine(" FROM DirPartyTable");
        sbSql.AppendLine(" INNER JOIN DIRPERSON ON DIRPARTYTABLE.RECID = DIRPERSON.RECID");
        sbSql.AppendLine(" INNER JOIN HCMWORKER ON HCMWORKER.PERSON = DIRPERSON.RECID ");
        sbSql.AppendLine(" INNER JOIN HCMWORKERTITLE ON HCMWORKERTITLE.WORKER = HCMWORKER.RECID ");
        sbSql.AppendLine("WHERE KNOWNAS = '" + strUser + "' AND HCMWORKERTITLE.VALIDTO > DateAdd(HH,+7,GETDATE())");

        DataTable dt = new DataTable();
        dt = ConnectionDAL.QueryDataTable(sbSql.ToString());
        return dt;

        }

        public DataTable getMenuByUser(string strUser)
        {
            StringBuilder sbSql = new StringBuilder();
         sbSql.AppendLine(" SELECT HcmWorkerTask.WORKERTASKID");
        sbSql.AppendLine(" FROM HCMWORKER");
        sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON DIRPARTYTABLE.RECID=HCMWORKER.PERSON");
        sbSql.AppendLine(" INNER JOIN HCMWORKERTASKASSIGNMENT ON HCMWORKERTASKASSIGNMENT.WORKER=HCMWORKER.RECID");
        sbSql.AppendLine(" INNER JOIN HcmWorkerTask ON HcmWorkerTask.RECID=HCMWORKERTASKASSIGNMENT.WORKERTASK");
        sbSql.AppendLine(" WHERE HcmWorkerTask.WORKERTASKID LIKE 'Report %'");
        sbSql.AppendLine(" AND DIRPARTYTABLE.KNOWNAS='"+strUser+"'");

        DataTable dt = new DataTable();
        dt = ConnectionDAL.QueryDataTable(sbSql.ToString());
        return dt;

        }

public DataTable getVenderGroup(string strVendGroup){

    StringBuilder sbSql = new StringBuilder();
         sbSql.AppendLine(" SELECT DISTINCT VendGroup AS VendGroup");
        sbSql.AppendLine(" FROM VendGroup");
      
        sbSql.AppendLine(" WHERE NOT(VendGroup IN ('" + strVendGroup + "'))");

        DataTable dt = new DataTable();
        dt = ConnectionDAL.QueryDataTable(sbSql.ToString());
        return dt;

}



	}
}
