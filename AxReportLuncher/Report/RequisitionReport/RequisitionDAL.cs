using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion.Report.RequisitionReport
{
    class RequisitionDAL
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


        public DataTable getSection(RequisitionOBJ RequisitionOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(" SELECT EmpOrgName [Section] FROM HOYA_vwRequisitionEmpData ");
            sbSql.AppendLine("WHERE");

            if (RequisitionOBJ.Factory == "GMO")
            {
                sbSql.AppendLine("  EmpWorkPlace='GMO'");
                sbSql.AppendLine(" AND NOT EmpOrgName LIKE '%NP1%' AND NOT EmpOrgName LIKE '%2%'");
            }
            else if (RequisitionOBJ.Factory == "SLR")
            {

                sbSql.AppendLine("  EmpWorkPlace='GMO'");
                sbSql.AppendLine(" AND (EmpOrgName LIKE '%NP1%' OR EmpOrgName LIKE '%2%') ");

            }
            else
            {
                sbSql.AppendLine("  EmpWorkPlace='" + RequisitionOBJ.Factory + "'");

            }

            sbSql.AppendLine(" GROUP BY EmpOrgName  ORDER BY EmpOrgName");


            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public ADODB.Recordset getRequistionList(RequisitionOBJ RequistionOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(RequistionOBJ.DateFrom.Year, RequistionOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(RequistionOBJ.DateTo.Year, RequistionOBJ.DateTo.Month, 1);

            String strFac = RequistionOBJ.strFactory;

            if (RequistionOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = RequistionOBJ.strFactory;
            }


            sbSql.AppendLine("SELECT ");
            sbSql.AppendLine("CASE WHEN EmpOrgName IS NULL THEN 'GRAND TOTAL' ELSE  EmpOrgName END [SECTION]");
            sbSql.AppendLine(",CASE WHEN inventtable.ITEMID IS NULL AND NOT EmpOrgName  IS NULL THEN 'TOTAL' ELSE  inventtable.ITEMID END");
            sbSql.AppendLine(",REQUSER [USER]");
            sbSql.AppendLine(",ECORESPRODUCTTRANSLATIONS.PRODUCTNAME");
            
            sbSql.AppendLine(",CASE  HOYA_IRTABLE.REQREQUISITIONTYPE");
            sbSql.AppendLine("WHEN 0 THEN 'DAMAGE' ");
            sbSql.AppendLine("WHEN 1 THEN 'YEARLY'");
            sbSql.AppendLine("WHEN 2 THEN  'PASSPRO'");
            sbSql.AppendLine("WHEN 3 THEN  'NEW'");
            sbSql.AppendLine("END");

            sbSql.AppendLine(",SUM(REQQTY) [QTY]");
            sbSql.AppendLine("FROM HOYA_IRTable INNER JOIN HOYA_IRLINE ON HOYA_IRTABLE.REQID = HOYA_IRLINE.REQID");
            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID = HOYA_IRLINE.ITEMID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATIONS ON ECORESPRODUCTTRANSLATIONS.PRODUCT = INVENTTABLE.PRODUCT");
            sbSql.AppendLine(" INNER JOIN [192.1.87.221].Employee.dbo.tb_EmployeeMstr emp ON  HOYA_IRLine.empcode=emp.EmpCode  ");
            
            sbSql.AppendLine(" WHERE REQFactory='" + RequistionOBJ.Factory + "'");
            //Confirm
            sbSql.AppendLine("AND REQ_STATUS = '2' AND  EmpOrgName != ''");

            sbSql.AppendLine(" AND CONVERT(date,HOYA_IRTable.REQADMRECEIPTDATE,103) BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", RequistionOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", RequistionOBJ.DateTo) + "',103)");
            //sbSql.AppendLine("  GROUP BY EmpOrgName,inventtable.ITEMID,REQUSER,ECORESPRODUCTTRANSLATIONS.PRODUCTNAME WITH ROLLUP");
            //sbSql.AppendLine(" HAVING  NOT ECORESPRODUCTTRANSLATIONS.PRODUCTNAME IS NULL OR inventtable.ITEMID IS NULL");

            sbSql.AppendLine("  GROUP BY EmpOrgName,inventtable.ITEMID,REQUSER,ECORESPRODUCTTRANSLATIONS.PRODUCTNAME, HOYA_IRTABLE.REQREQUISITIONTYPE WITH ROLLUP");
            sbSql.AppendLine(" HAVING  NOT HOYA_IRTABLE.REQREQUISITIONTYPE IS NULL OR inventtable.ITEMID IS NULL");
            
        
                      

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getAnnualReport(RequisitionOBJ RequistionOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(RequistionOBJ.DateFrom.Year, RequistionOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(RequistionOBJ.DateTo.Year, RequistionOBJ.DateTo.Month, 1);


            sbSql.AppendLine("SELECT CODE,NAME,EmpOrgName,EmpEmployDate,Common,Size,CASE WHEN Qty IS NULL THEN 0 ELSE QTY END QTY FROM HOYA_vwRequisitionEmpData ");
            sbSql.AppendLine("WHERE");

            

            if (RequistionOBJ.Factory == "SLR")
            {
                sbSql.AppendLine("  EmpWorkPlace='GMO'");

                if (RequistionOBJ.Section == "All")
                {
                    sbSql.AppendLine(" AND ( EmpOrgName LIKE '%NP1%' OR EmpOrgName LIKE '%2%') ");
                }
                else
                {
                    sbSql.AppendLine(" AND ( EmpOrgName = '"+RequistionOBJ.Section+"') ");
                }
                
            }
            else if (RequistionOBJ.Factory == "GMO")
            {
                sbSql.AppendLine("  EmpWorkPlace='GMO'");

                if (RequistionOBJ.Section == "All")
                {
                    sbSql.AppendLine(" AND NOT EmpOrgName LIKE '%NP1%' AND NOT EmpOrgName LIKE '%2%'");
                }
                else
                {
                    sbSql.AppendLine(" AND ( EmpOrgName = '" + RequistionOBJ.Section + "') ");
                }

            }
            else
            {
                sbSql.AppendLine("  EmpWorkPlace='" + RequistionOBJ.Factory + "'");
             
                if (RequistionOBJ.Section != "All")
                {
                   sbSql.AppendLine(" AND ( EmpOrgName = '" + RequistionOBJ.Section + "') ");
                }
            }

           sbSql.AppendLine("AND  Size !=''");
            //sbSql.AppendLine("AND  Qty !=0");
            sbSql.AppendLine(" AND MONTH(EmpEmployDate) = '" + RequistionOBJ.DateFrom.Month + "'");
          // sbSql.AppendLine(" AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", RequistionOBJ.DateTo) + "',103)");
            sbSql.AppendLine("   ORDER BY EmpOrgName");
   





            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getAnnualReport1(RequisitionOBJ RequistionOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(RequistionOBJ.DateFrom.Year, RequistionOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(RequistionOBJ.DateTo.Year, RequistionOBJ.DateTo.Month, 1);


            sbSql.AppendLine("SELECT EmpCode CODE,NAME,EmpOrgName,EmpEmployDate,Common,Size,CASE WHEN Qty IS NULL THEN 0 ELSE QTY END QTY FROM HOYA_vwAnnualReport ");
            sbSql.AppendLine("WHERE");



            if (RequistionOBJ.Factory == "NP1")
            {
                sbSql.AppendLine("  EmpWorkPlace='NP1'");

                if (RequistionOBJ.Section == "All")
                {
                    sbSql.AppendLine(" AND ( EmpOrgName LIKE '%NP1%' OR EmpOrgName LIKE '%2%') ");
                }
                else
                {
                    sbSql.AppendLine(" AND ( EmpOrgName = '" + RequistionOBJ.Section + "') ");
                }

            }
            else if (RequistionOBJ.Factory == "GMO")
            {
                sbSql.AppendLine("  EmpWorkPlace='GMO'");

                if (RequistionOBJ.Section == "All")
                {
                    sbSql.AppendLine(" AND NOT EmpOrgName LIKE '%NP1%' AND NOT EmpOrgName LIKE '%2%'");
                }
                else
                {
                    sbSql.AppendLine(" AND ( EmpOrgName = '" + RequistionOBJ.Section + "') ");
                }

            }
            else
            {
                sbSql.AppendLine("  EmpWorkPlace='" + RequistionOBJ.Factory + "'");

                if (RequistionOBJ.Section != "All")
                {
                    sbSql.AppendLine(" AND ( EmpOrgName = '" + RequistionOBJ.Section + "') ");
                }
            }

            sbSql.AppendLine("AND  Size !=''");
            //sbSql.AppendLine("AND  Qty !=0");
            sbSql.AppendLine(" AND MONTH(EmpEmployDate) = '" + RequistionOBJ.DateFrom.Month + "'");
            // sbSql.AppendLine(" AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", RequistionOBJ.DateTo) + "',103)");
            sbSql.AppendLine("   ORDER BY EmpOrgName");






            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset SummaryByItem(RequisitionOBJ RequistionOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            DateTime dtFrom = new DateTime(RequistionOBJ.DateFrom.Year, RequistionOBJ.DateFrom.Month, 1);
            DateTime dtTo = new DateTime(RequistionOBJ.DateTo.Year, RequistionOBJ.DateTo.Month, 1);

            String strFac = RequistionOBJ.strFactory;

            if (RequistionOBJ.strFactory == "GMO")
            {
                strFac = "MO";
            }
            else
            {
                strFac = RequistionOBJ.strFactory;
            }


            sbSql.AppendLine("SELECT ");
           // sbSql.AppendLine("CASE WHEN HOYA_IRSECTIONMAP.SECTIONHR IS NULL THEN 'GRAND TOTAL' ELSE  HOYA_IRSECTIONMAP.SECTIONHR END [SECTION]");
           // sbSql.AppendLine(",CASE WHEN inventtable.ITEMID IS NULL AND NOT HOYA_IRSECTIONMAP.SECTIONHR  IS NULL THEN 'TOTAL' ELSE  inventtable.ITEMID END");
           // sbSql.AppendLine(",REQUSER [USER]");
            sbSql.AppendLine("inventtable.ITEMID");
            sbSql.AppendLine(",ECORESPRODUCTTRANSLATIONS.PRODUCTNAME");
            sbSql.AppendLine(",SUM(REQQTY) [QTY]");
            sbSql.AppendLine("FROM HOYA_IRTable INNER JOIN HOYA_IRLINE ON HOYA_IRTABLE.REQID = HOYA_IRLINE.REQID");
            sbSql.AppendLine("INNER JOIN INVENTTABLE ON INVENTTABLE.ITEMID = HOYA_IRLINE.ITEMID");
            sbSql.AppendLine("INNER JOIN ECORESPRODUCTTRANSLATIONS ON ECORESPRODUCTTRANSLATIONS.PRODUCT = INVENTTABLE.PRODUCT");
            sbSql.AppendLine("INNER JOIN [192.1.87.221].Employee.dbo.tb_EmployeeMstr emp ON  HOYA_IRLine.empcode=emp.EmpCode ");

            sbSql.AppendLine(" WHERE REQFactory='" + RequistionOBJ.Factory + "'");
            sbSql.AppendLine("AND REQ_STATUS = '2' AND  EmpOrgName != ''");

            sbSql.AppendLine(" AND CONVERT(date,HOYA_IRTable.REQADMRECEIPTDATE,103) BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", RequistionOBJ.DateFrom) + "',103) ");
          
            //sbSql.AppendLine(" AND CONVERT(date,HOYA_IRTable.CREATEDDATETIME,103) BETWEEN CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", RequistionOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", RequistionOBJ.DateTo) + "',103)");
            //sbSql.AppendLine("  GROUP BY SECTIONHR,inventtable.ITEMID,REQUSER,ECORESPRODUCTTRANSLATIONS.PRODUCTNAME WITH ROLLUP");
            //sbSql.AppendLine(" HAVING  NOT ECORESPRODUCTTRANSLATIONS.PRODUCTNAME IS NULL OR inventtable.ITEMID IS NULL");

            sbSql.AppendLine("  GROUP BY inventtable.ITEMID,ECORESPRODUCTTRANSLATIONS.PRODUCTNAME");



            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }





    }
}
