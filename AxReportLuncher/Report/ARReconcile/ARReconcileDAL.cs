using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;

namespace NewVersion.Report.ARReconcile
{
    class ARReconcileDAL
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

        public DataTable getInvoiceDueDate2(ARReconcileOBJ ARReconcileOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(" SELECT * FROM (");
            sbSql.AppendLine(" SELECT [DueDate] FROM ( SELECT");
            sbSql.AppendLine(" CONVERT(DATETIME,CONVERT(CHAR(4),CASE PaymTerm.NUMOFDAYS  WHEN 30 THEN DATEADD(m,1,CUSTINVOICEJOUR.INVOICEDATE)");
            sbSql.AppendLine(" WHEN 60 THEN DATEADD(m,2,CUSTINVOICEJOUR.INVOICEDATE)");
            sbSql.AppendLine(" WHEN 90 THEN DATEADD(m,3,CUSTINVOICEJOUR.INVOICEDATE)");
            sbSql.AppendLine(" END,12)+'01',12) [DueDate]");
            sbSql.AppendLine(" FROM SALESTABLE");
            sbSql.AppendLine(" INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
            sbSql.AppendLine("  AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("  AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
            sbSql.AppendLine("  AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN CustTransOpen on CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("  AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
            sbSql.AppendLine("  AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT");
            sbSql.AppendLine("  AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE SALESTABLE.DATAAREAID='HOYA' AND CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "' AND PaymTerm.NUMOFDAYS>0");
            sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE <= CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103) ");
            sbSql.AppendLine(" ) AS DueDate");
            sbSql.AppendLine(" GROUP BY [DueDate]");
            sbSql.AppendLine(")Duedate");
            sbSql.AppendLine("WHERE DueDate >= CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103) ");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }


        /*
        public DataTable getInvoiceAccountCurr(ARReconcileOBJ ARReconcileOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
        sbSql.AppendLine(" SELECT CUSTINVOICEJOUR.INVOICEACCOUNT,DIRPARTYTABLE.NAME,CURRENCY.CURRENCYCODEISO [Curr]");
        //sbSql.AppendLine(" ,SALESTABLE.PAYMENT,ISNULL(PaymTerm.NUMOFDAYS,90)/30 MONTHCOUNT")
        //sbSql.AppendLine(" ,ISNULL(MAX(PaymTerm.NUMOFDAYS),90)/30 MONTHCOUNT")
        sbSql.AppendLine(" FROM SALESTABLE");
        sbSql.AppendLine(" INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
        sbSql.AppendLine(" 	AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine(" 	AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
        sbSql.AppendLine(" 	AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
        sbSql.AppendLine(" INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
        sbSql.AppendLine(" INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine(" LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");
        sbSql.AppendLine(" WHERE SALESTABLE.DATAAREAID='HOYA' AND NOT(SALESTABLE.PAYMENT='NOCOM') AND CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "'");
        //sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE BETWEEN CONVERT(DATETIME,'" & Format(ARReconcileOBJ.DateFrom, "dd/MM/yyyy") & "',103) AND CONVERT(DATETIME,'" & Format(ARReconcileOBJ.DateTo, "dd/MM/yyyy") & "',103)")
        sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE <= CONVERT(DATETIME,'" +String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103)");
        sbSql.AppendLine(" GROUP BY CUSTINVOICEJOUR.INVOICEACCOUNT,DIRPARTYTABLE.NAME,CURRENCY.CURRENCYCODEISO");
        //sbSql.AppendLine(" ,SALESTABLE.PAYMENT,PaymTerm.NUMOFDAYS")
        sbSql.AppendLine(" ORDER BY CUSTINVOICEJOUR.INVOICEACCOUNT");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }
         */

        public DataTable getReconcileSummaryAcc(ARReconcileOBJ ARReconcileOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
        sbSql.AppendLine(" SELECT DIMENSIONATTRIBUTEVALUECOMBINATION.DISPLAYVALUE");
        sbSql.AppendLine(" ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNTMST END) [BAHT]");
        sbSql.AppendLine(" FROM SALESTABLE");
        sbSql.AppendLine(" INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
        sbSql.AppendLine("     AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("     AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
        sbSql.AppendLine("     AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("     AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN CUSTLEDGERACCOUNTS ON CUSTLEDGERACCOUNTS.NUM=CUSTTABLE.CUSTGROUP ");
        sbSql.AppendLine(" 	   AND CUSTTRANS.POSTINGPROFILE=CUSTLEDGERACCOUNTS.POSTINGPROFILE");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUECOMBINATION ON DIMENSIONATTRIBUTEVALUECOMBINATION.RECID=CUSTLEDGERACCOUNTS.SUMMARYLEDGERDIMENSION");
        sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
        sbSql.AppendLine(" INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
        sbSql.AppendLine(" INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("     AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
        sbSql.AppendLine("     AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine(" LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT");
        sbSql.AppendLine("     AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");
        sbSql.AppendLine(" WHERE CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "'");
        sbSql.AppendLine("     AND NOT(SALESTABLE.PAYMENT='NOCOM')");
        sbSql.AppendLine("     AND SALESTABLE.DATAAREAID='HOYA'");
        sbSql.AppendLine("     AND CUSTINVOICEJOUR.INVOICEDATE <= CONVERT(DATETIME,'" +String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103)");
        sbSql.AppendLine(" GROUP BY DIMENSIONATTRIBUTEVALUECOMBINATION.DISPLAYVALUE");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public DataTable getReconcileSummaryCurr(ARReconcileOBJ ARReconcileOBJ)
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(" SELECT CURRENCY.CURRENCYCODEISO");
            sbSql.AppendLine(" ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNTMST END) [Amt]");
            sbSql.AppendLine(" FROM SALESTABLE");
            sbSql.AppendLine(" INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
            sbSql.AppendLine("     AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("     AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
            sbSql.AppendLine("     AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
            sbSql.AppendLine(" INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("     AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
            sbSql.AppendLine("     AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" WHERE CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "'");
            sbSql.AppendLine("     AND NOT(SALESTABLE.PAYMENT='NOCOM')");
            sbSql.AppendLine("     AND SALESTABLE.DATAAREAID='HOYA'");
            sbSql.AppendLine("     AND CUSTINVOICEJOUR.INVOICEDATE <= CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" GROUP BY CURRENCY.CURRENCYCODEISO");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }


        public ADODB.Recordset getReconcileSummary(DataTable dt, ARReconcileOBJ ARReconcileOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            try
            {
               // sbSql.AppendLine(" SELECT ROW_NUMBER() OVER(order by NAME)[No],NAME,CURR");
                sbSql.AppendLine("SELECT ROW_NUMBER() OVER(order by NAME)[No],NAME + ' - ' +CURR");
                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(String.Format(" ,SUM([{0}THB])[{0}THB]", String.Format("{0:yyMM}", dr["DUEDATE"])));
                    sbSql.Append(String.Format(" ,SUM([{0}YEN])[{0}YEN]", String.Format("{0:yyMM}", dr["DUEDATE"])));
                    sbSql.Append(String.Format(" ,SUM([{0}USD])[{0}USD]", String.Format("{0:yyMM}", dr["DUEDATE"])));
                    sbSql.Append(String.Format(" ,SUM([{0}CNY])[{0}CNY]", String.Format("{0:yyMM}", dr["DUEDATE"])));
                    sbSql.Append(String.Format(" ,SUM([{0}TOTAL])[{0}TOTAL]", String.Format("{0:yyMM}", dr["DUEDATE"])));
                }

               // sbSql.AppendLine(" ,SUM([InvoiceAmtCurr]) [InvoiceAmtCurr],SUM([InvoiceAmtBHT]) [InvoiceAmtBHT]");
                sbSql.AppendLine(" FROM (");

                sbSql.AppendLine(" SELECT NAME ,[Curr]");
                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", dr["DUEDATE"]) + "',103) AND CURR='THB' THEN SUM([InvoiceAmt]) ELSE 0 END [" + String.Format("{0:yyMM}", dr["DUEDATE"]) + "THB]");
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", dr["DUEDATE"]) + "',103) AND CURR='JPY' THEN SUM([InvoiceAmt]) ELSE 0 END [" + String.Format("{0:yyMM}", dr["DUEDATE"]) + "YEN]");
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", dr["DUEDATE"]) + "',103) AND CURR='USD' THEN SUM([InvoiceAmt]) ELSE 0 END[" + String.Format("{0:yyMM}", dr["DUEDATE"]) + "USD]");
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", dr["DUEDATE"]) + "',103) AND CURR='CNY' THEN SUM([InvoiceAmt]) ELSE 0 END [" + String.Format("{0:yyMM}", dr["DUEDATE"]) + "CNY]");
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", dr["DUEDATE"]) + "',103) THEN SUM([InvoiceAmtBHT]) ELSE 0 END [" + String.Format("{0:yyMM}", dr["DUEDATE"]) + "TOTAL]");

                }


                sbSql.AppendLine(" ,SUM([InvoiceAmt]) [InvoiceAmtCurr],SUM([InvoiceAmtBHT]) [InvoiceAmtBHT],DUEDATE");
            sbSql.AppendLine(" FROM ( ");
            sbSql.AppendLine(" SELECT CUSTTABLE.ACCOUNTNUM");
            sbSql.AppendLine("  ,CASE WHEN CUSTINVOICEJOUR.INVOICEACCOUNT LIKE 'ARAO%' THEN CASE WHEN  CUSTINVOICEJOUR.PRINTMGMTSITEID ='GMO' THEN DIRPARTYTABLE.NAME +' (KATA)' ELSE DIRPARTYTABLE.NAME END ELSE DIRPARTYTABLE.NAME END [NAME]");
            //sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTTRANS.DUEDATE,12)+'01',12) DUEDATE");
            sbSql.AppendLine("  ,CASE PaymTerm.NUMOFDAYS  WHEN 30 THEN DATEADD(m,1,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
            sbSql.AppendLine("                              WHEN 60 THEN DATEADD(m,2,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
            sbSql.AppendLine("                              WHEN 90 THEN DATEADD(m,3,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
            sbSql.AppendLine("  END [DueDate]");
            sbSql.AppendLine(" ,CURRENCY.CURRENCYCODEISO [CURR] ");
            //sbSql.AppendLine(" ,CUSTTABLE.ECL_REASON [TradingPartner]");
            sbSql.AppendLine(" ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNT END) [InvoiceAmt] ");
            sbSql.AppendLine(" ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNTMST END) [InvoiceAmtBHT]");
            sbSql.AppendLine("  FROM SALESTABLE");
            sbSql.AppendLine(" INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID");
            sbSql.AppendLine("  AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
            sbSql.AppendLine("  AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("  AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
            sbSql.AppendLine("  AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("  AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
            sbSql.AppendLine(" INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
            sbSql.AppendLine(" INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("  AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
            sbSql.AppendLine("  AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT");
            sbSql.AppendLine("  AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE SALESTABLE.DATAAREAID='HOYA' AND CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "'");
            //sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE BETWEEN CONVERT(DATETIME,'" & Format(ARReconcileOBJ.DateFrom, "dd/MM/yyyy") & "',103) AND CONVERT(DATETIME,'" & Format(ARReconcileOBJ.DateTo, "dd/MM/yyyy") & "',103)")
            sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE <= CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103)");
            sbSql.AppendLine("AND ECL_SALESCOMERCIAL = 1");
            sbSql.AppendLine(" GROUP BY CUSTTABLE.ACCOUNTNUM,NAME,CUSTINVOICEJOUR.INVOICEACCOUNT,CUSTINVOICEJOUR.PRINTMGMTSITEID ");
            sbSql.AppendLine(" ,PaymTerm.NUMOFDAYS");
            sbSql.AppendLine(" ,CUSTINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine(",CURRENCY.CURRENCYCODEISO");
            sbSql.AppendLine(",CUSTTABLE.ECL_REASON");
            sbSql.AppendLine(" ) ARDetail");
            sbSql.AppendLine("WHERE [InvoiceAmt] != 0 OR [InvoiceAmtBHT] !=0");
            sbSql.AppendLine(" GROUP BY ACCOUNTNUM,NAME,DUEDATE,[Curr]");
            sbSql.AppendLine(" ) ARByDueDate");
            sbSql.AppendLine(" GROUP BY NAME,CURR");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;

            } //end Try

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }


        public ADODB.Recordset getReconcileSummaryDetail(DataTable dt, ARReconcileOBJ ARReconcileOBJ,string Type)
        {
            StringBuilder sbSql = new StringBuilder();

            try
            {
                sbSql.AppendLine(" SELECT NAME + ' - ' +CURR ,[Trad Part]");
                sbSql.AppendLine(" FROM (");

                sbSql.AppendLine(" SELECT NAME ,[Curr],[Trad Part]");
                sbSql.AppendLine(" FROM ( ");

                sbSql.AppendLine(" SELECT CUSTTABLE.ACCOUNTNUM");
                sbSql.AppendLine("  ,CASE WHEN CUSTINVOICEJOUR.INVOICEACCOUNT LIKE 'ARAO%' THEN CASE WHEN  CUSTINVOICEJOUR.PRINTMGMTSITEID ='GMO' THEN DIRPARTYTABLE.NAME +' (KATA)' ELSE DIRPARTYTABLE.NAME END ELSE DIRPARTYTABLE.NAME END [NAME]");
                sbSql.AppendLine("  ,CASE PaymTerm.NUMOFDAYS  WHEN 30 THEN DATEADD(m,1,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
                sbSql.AppendLine("                              WHEN 60 THEN DATEADD(m,2,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
                sbSql.AppendLine("                              WHEN 90 THEN DATEADD(m,3,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
                sbSql.AppendLine("  END [DueDate]");
                sbSql.AppendLine(" ,CURRENCY.CURRENCYCODEISO [CURR] ");
                sbSql.AppendLine(" ,CUSTTABLE.ECL_REASON [Trad Part]");
                sbSql.AppendLine(" ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNT END) [InvoiceAmt] ");
                sbSql.AppendLine(" ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNTMST END) [InvoiceAmtBHT]");
                sbSql.AppendLine("  FROM SALESTABLE");
                sbSql.AppendLine(" INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID");
                sbSql.AppendLine("  AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
                sbSql.AppendLine("  AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
                sbSql.AppendLine("  AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
                sbSql.AppendLine("  AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
                sbSql.AppendLine("  AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
                sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
                sbSql.AppendLine(" INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
                sbSql.AppendLine(" INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
                sbSql.AppendLine("  AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
                sbSql.AppendLine("  AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
                sbSql.AppendLine(" LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT");
                sbSql.AppendLine("  AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");
                sbSql.AppendLine(" WHERE SALESTABLE.DATAAREAID='HOYA' AND CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "'");
                sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE <= CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103)");
                sbSql.AppendLine("AND ECL_SALESCOMERCIAL = 1");

                if (Type == "1")
                {
                    sbSql.AppendLine("AND ECL_REASON != '' AND CUSTINVOICEJOUR.INVOICEACCOUNT LIKE 'AREX%'");
                    //AND ECL_REASON != 'J118'

                }
                else if (Type == "2")
                {
                    sbSql.AppendLine("AND ECL_REASON = '' AND CUSTINVOICEJOUR.INVOICEACCOUNT LIKE 'AREX%'");

                }
                else if (Type == "3")
                {
                    sbSql.AppendLine("AND ECL_REASON != '' AND CUSTINVOICEJOUR.INVOICEACCOUNT LIKE 'ARAO%'");

                }
                else if (Type == "4")
                {
                    sbSql.AppendLine("AND ECL_REASON != '' AND CUSTINVOICEJOUR.INVOICEACCOUNT LIKE 'ARIN%'");

                }

                sbSql.AppendLine(" GROUP BY CUSTTABLE.ACCOUNTNUM,NAME,CUSTINVOICEJOUR.INVOICEACCOUNT,CUSTINVOICEJOUR.PRINTMGMTSITEID ");
                sbSql.AppendLine(" ,PaymTerm.NUMOFDAYS");
                sbSql.AppendLine(" ,CUSTINVOICEJOUR.INVOICEDATE");
                sbSql.AppendLine(",CURRENCY.CURRENCYCODEISO");
                sbSql.AppendLine(",CUSTTABLE.ECL_REASON");
                sbSql.AppendLine(" ) ARDetail");
                sbSql.AppendLine("WHERE    [InvoiceAmt] != 0 OR [InvoiceAmtBHT] !=0");
                sbSql.AppendLine(" GROUP BY ACCOUNTNUM,NAME,DUEDATE,[Curr],[Trad Part]");
                sbSql.AppendLine(" ) ARByDueDate");
                sbSql.AppendLine(" GROUP BY NAME,CURR,[Trad Part]");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;

            } //end Try

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getInvoiceDetail(DataTable dt,ARReconcileOBJ ARReconcileOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            try
            {
                sbSql.AppendLine(" SELECT ROW_NUMBER() OVER(order by NAME)[No],NAME,CURR");
                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(String.Format(" ,SUM([{0}THB])[{0}THB]",String.Format("{yyMM}",dr["DUEDATE"])));
                    sbSql.Append(String.Format(" ,SUM([{0}YEN])[{0}YEN]", String.Format("{yyMM}", dr["DUEDATE"])));
                    sbSql.Append(String.Format(" ,SUM([{0}USD])[{0}USD]", String.Format("{yyMM}", dr["DUEDATE"])));
                    sbSql.Append(String.Format(" ,SUM([{0}CNY])[{0}CNY]", String.Format("{yyMM}", dr["DUEDATE"])));
                    sbSql.Append(String.Format(" ,SUM([{0}TOTAL])[{0}TOTAL]", String.Format("{yyMM}", dr["DUEDATE"])));
                }

                sbSql.AppendLine(" ,SUM([InvoiceAmtCurr]) [InvoiceAmtCurr],SUM([InvoiceAmtBHT]) [InvoiceAmtBHT]");
                sbSql.AppendLine(" FROM (");

                sbSql.AppendLine(" SELECT NAME ,[Curr]");
                foreach (DataRow dr in dt.Rows)
                {
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{dd/MM/yyyy}", dr["DUEDATE"]) + "',103) AND CURR='THB' THEN SUM([InvoiceAmt]) ELSE 0 END [" + String.Format("{yyMM}", dr["DUEDATE"]) + "THB]");
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{dd/MM/yyyy}", dr["DUEDATE"]) + "',103) AND CURR='JPY' THEN SUM([InvoiceAmt]) ELSE 0 END [" + String.Format("{yyMM}", dr["DUEDATE"]) + "YEN]");
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{dd/MM/yyyy}", dr["DUEDATE"]) + "',103) AND CURR='USD' THEN SUM([InvoiceAmt]) ELSE 0 END[" + String.Format("{yyMM}", dr["DUEDATE"]) + "USD]");
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{dd/MM/yyyy}", dr["DUEDATE"]) + "',103) AND CURR='CNY' THEN SUM([InvoiceAmt]) ELSE 0 END [" + String.Format("{yyMM}", dr["DUEDATE"]) + "CNY]");
                    sbSql.AppendLine(" ,CASE WHEN DUEDATE=CONVERT(DATETIME,'" + String.Format("{dd/MM/yyyy}", dr["DUEDATE"]) + "',103) THEN SUM([InvoiceAmtBHT]) ELSE 0 END [" + String.Format("{yyMM}", dr["DUEDATE"]) + "TOTAL]");

                }


           sbSql.AppendLine(" ,SUM([InvoiceAmt]) [InvoiceAmtCurr],SUM([InvoiceAmtBHT]) [InvoiceAmtBHT],DUEDATE");
            sbSql.AppendLine(" FROM ( ");
            sbSql.AppendLine(" SELECT CUSTTABLE.ACCOUNTNUM, DIRPARTYTABLE.NAME");
            //sbSql.AppendLine(" ,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTTRANS.DUEDATE,12)+'01',12) DUEDATE")
            sbSql.AppendLine("  ,CASE PaymTerm.NUMOFDAYS  WHEN 30 THEN DATEADD(m,1,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
            sbSql.AppendLine("                              WHEN 60 THEN DATEADD(m,2,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
            sbSql.AppendLine("                              WHEN 90 THEN DATEADD(m,3,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
            sbSql.AppendLine("  END [DueDate]");
            sbSql.AppendLine(" ,CURRENCY.CURRENCYCODEISO [CURR] ");
            sbSql.AppendLine(" ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNT END) [InvoiceAmt] ");
            sbSql.AppendLine(" ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNTMST END) [InvoiceAmtBHT]");
            sbSql.AppendLine("  FROM SALESTABLE");
            sbSql.AppendLine(" INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID");
            sbSql.AppendLine("  AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
            sbSql.AppendLine("  AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("  AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
            sbSql.AppendLine("  AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("  AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
            sbSql.AppendLine(" INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
            sbSql.AppendLine(" INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("  AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
            sbSql.AppendLine("  AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine(" LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT");
            sbSql.AppendLine("  AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");
            sbSql.AppendLine(" WHERE SALESTABLE.DATAAREAID='HOYA' AND CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "'");
            //sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE BETWEEN CONVERT(DATETIME,'" & Format(ARReconcileOBJ.DateFrom, "dd/MM/yyyy") & "',103) AND CONVERT(DATETIME,'" & Format(ARReconcileOBJ.DateTo, "dd/MM/yyyy") & "',103)")
            sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE <= CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" GROUP BY CUSTTABLE.ACCOUNTNUM,NAME");
            sbSql.AppendLine(" ,PaymTerm.NUMOFDAYS");
            sbSql.AppendLine(" ,CUSTINVOICEJOUR.INVOICEDATE");
            //sbSql.AppendLine(",CONVERT(DATETIME,CONVERT(CHAR(4),CUSTTRANS.DUEDATE,12)+'01',12)")
            sbSql.AppendLine(",CURRENCY.CURRENCYCODEISO");
            sbSql.AppendLine(" ) ARDetail");
            sbSql.AppendLine(" GROUP BY ACCOUNTNUM,NAME,DUEDATE,[Curr]");
            sbSql.AppendLine(" ) ARByDueDate");
            sbSql.AppendLine(" GROUP BY NAME,CURR");

            }
            catch (Exception ex)
            {
            MessageBox.Show(ex.Message);
            return null;

            } //end Try

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

        public ADODB.Recordset getInvoiceDetail3( ARReconcileOBJ ARReconcileOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

           sbSql.AppendLine(" SELECT");
        sbSql.AppendLine("  CASE WHEN [Voucher No] IS NULL THEN NULL ELSE [Due Date] END [Due Date] ");
        //sbSql.AppendLine("''  [Due Date] ")
        sbSql.AppendLine("  ,[Voucher No],[Invoice No],[Invoice Date]");
        //sbSql.AppendLine("  ,CASE WHEN (NUMOFDAYS IS NULL AND PAYMENT IS NULL) THEN 'Grand Total' ELSE CASE WHEN [Curr] IS NULL THEN 'Total' ELSE [Curr] END END [Curr]")
        sbSql.AppendLine("  ,CASE WHEN ([Due Date] IS NULL AND [Voucher No] IS NULL) THEN 'Grand Total' ELSE CASE WHEN [Voucher No] IS NULL THEN 'Total' ELSE [Curr] END END [Curr]");
        sbSql.AppendLine("  ,SUM(AmtCurr),SUM([Baht]),Remark");
        sbSql.AppendLine(" FROM ( ");
        sbSql.AppendLine("  SELECT ");
        sbSql.AppendLine("      REPLACE(SUBSTRING(CONVERT(CHAR(10),");
        sbSql.AppendLine("          CASE PaymTerm.NUMOFDAYS  WHEN 30 THEN DATEADD(m,1,CUSTINVOICEJOUR.INVOICEDATE)");
        sbSql.AppendLine("                              WHEN 60 THEN DATEADD(m,2,CUSTINVOICEJOUR.INVOICEDATE)");
        sbSql.AppendLine("                              WHEN 90 THEN DATEADD(m,3,CUSTINVOICEJOUR.INVOICEDATE)");
        sbSql.AppendLine("          END ");
        sbSql.AppendLine("      ,6),4,6),' ','-')");
        sbSql.AppendLine("  [Due Date]");
        sbSql.AppendLine("  ,CASE PaymTerm.NUMOFDAYS  WHEN 30 THEN DATEADD(m,1,CUSTINVOICEJOUR.INVOICEDATE)");
        sbSql.AppendLine("                              WHEN 60 THEN DATEADD(m,2,CUSTINVOICEJOUR.INVOICEDATE)");
        sbSql.AppendLine("                              WHEN 90 THEN DATEADD(m,3,CUSTINVOICEJOUR.INVOICEDATE)");
        sbSql.AppendLine("  END [DueDate]");
        sbSql.AppendLine("  ,CUSTINVOICEJOUR.LedgerVoucher [Voucher No]");
        sbSql.AppendLine("  ,CUSTINVOICEJOUR.INVOICEID [Invoice No]");
        sbSql.AppendLine("  ,CUSTINVOICEJOUR.INVOICEDATE [Invoice Date]");
        sbSql.AppendLine("  ,CURRENCY.CURRENCYCODEISO [Curr]");
        sbSql.AppendLine("  ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNT END) AmtCurr");
        sbSql.AppendLine("  ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNTMST END) [Baht]");
        //sbSql.AppendLine("  ,CASE WHEN SALESTABLE.PAYMENT='Nocom' THEN 'NOCOM' ELSE '' END Remark")
        sbSql.AppendLine("  ,CASE WHEN SALESTABLE.ECL_SALESCOMERCIAL = 1 THEN 'COMERCIAL' ELSE 'NOCOMERCIAL' END Remark");

        sbSql.AppendLine("  ,SALESTABLE.PAYMENT");
        sbSql.AppendLine("  ,PaymTerm.NUMOFDAYS");
        sbSql.AppendLine("  FROM SALESTABLE");
        sbSql.AppendLine("  INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine("  INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("  INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("      AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("  INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
        sbSql.AppendLine("  INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
        sbSql.AppendLine("  INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("      AND CustTransOpen.REFRECID=CUSTTRANS.RECID"); 
        sbSql.AppendLine("      AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("  LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT");
        sbSql.AppendLine("      AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");
        sbSql.AppendLine("  WHERE SALESTABLE.DATAAREAID='HOYA'");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory +"'");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.INVOICEACCOUNT='" + ARReconcileOBJ.InvoiceAccount + "'");
        sbSql.AppendLine("      AND CURRENCY.CURRENCYCODEISO='" + ARReconcileOBJ.CurrencyISO + "'");
        sbSql.AppendLine(" AND ECL_SALESCOMERCIAL = 1 AND INVOICEDATE<=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103) ");
        //sbSql.AppendLine("      AND CUSTINVOICEJOUR.INVOICEDATE BETWEEN CONVERT(DATETIME,'" & Format(ARReconcileOBJ.DateFrom, "dd/MM/yyyy") & "',103) AND CONVERT(DATETIME,'" & Format(ARReconcileOBJ.DateTo, "dd/MM/yyyy") & "',103)")
        sbSql.AppendLine("  GROUP BY");
        sbSql.AppendLine("      PaymTerm.NUMOFDAYS,CUSTINVOICEJOUR.LedgerVoucher,CUSTINVOICEJOUR.INVOICEID");
        sbSql.AppendLine("      ,CUSTINVOICEJOUR.INVOICEDATE,CURRENCY.CURRENCYCODEISO,SALESTABLE.PAYMENT,SALESTABLE.ECL_SALESCOMERCIAL");
        sbSql.AppendLine(" ) AS ARReconcile");
        sbSql.AppendLine(" WHERE [Invoice Date]<=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103) AND AmtCurr != 0 OR [Baht] !=0 ");
        sbSql.AppendLine(" GROUP BY [Due Date],[Invoice Date],[Invoice No],[Voucher No],[Curr],Remark WITH ROLLUP");

        //sbSql.AppendLine(" HAVING NOT(Remark IS NULL) OR ([Invoice Date] IS NULL)")
        sbSql.AppendLine("HAVING NOT Remark IS NULL OR ( [Voucher No] IS NULL AND [Invoice No] IS NULL)");

        //  sbSql.AppendLine(" ORDER BY [Due Date] ASC")

            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }

    public ADODB.Recordset getInvoiceDetail3Edit( ARReconcileOBJ ARReconcileOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

         sbSql.AppendLine(" SELECT");

         sbSql.AppendLine("   CASE WHEN [Invoice Date] IS NULL THEN   ARReconcil.[Due Date]   ELSE '' END [Due Date]");
        sbSql.AppendLine(",CASE WHEn ARReconcil.[Voucher No] IS NULL AND NOT ARReconcil.[Due Date] IS NULl THEN 'TOTAL' ELse ARReconcil.[Voucher No] END [Voucher No]");
       // sbSql.AppendLine(" ,CASE WHEN ARReconcil.[Invoice No] IS NULL AND ARReconcil.[Due Date] IS NULL  THEN 'Summary for All Invoice'  ");
        //sbSql.AppendLine(" 	 WHEN ARReconcil.[Invoice No] IS NULL  THEN  'Summary for Invoice' ");
        //sbSql.AppendLine("   ELSE ARReconcil.[Invoice No] END [Invoice No]");
        sbSql.AppendLine(",ARReconcil.[Invoice No]  [Invoice No]");

         //sbSql.AppendLine("CASE WHEN ARReconcil.[Voucher No] IS NULL THEN ARReconcil.[Due Date] ELSE NULL  END [Due Date]");
         //sbSql.AppendLine(",CASE WHEN NOT ARReconcil.[Due Date] IS NULL AND  ARReconcil.[Invoice Date] IS NULL THEN 'Summary for invoice' ELSE");
        // sbSql.AppendLine("CASE WHEN ARReconcil.[Voucher No] IS NULL AND  ARReconcil.[Invoice Date] IS NULL THEN 'Summary for all invoice' ELSE");
         //sbSql.AppendLine(" ARReconcil.[Voucher No] END END [Voucher No]");

        sbSql.AppendLine(" ,ARReconcil.[Invoice Date]");
        sbSql.AppendLine("  ,ARReconcil.Curr");
        sbSql.AppendLine("  ,SUM(ARReconcil.AmtCurr)[AmountCurr]");
        sbSql.AppendLine("  ,SUM(ARReconcil.AmountTHB)[AmountTHB]");
       // sbSql.AppendLine(" ,''[Remark]");

        sbSql.AppendLine("   FROM (");
        sbSql.AppendLine(" SELECT");

        sbSql.AppendLine("  REPLACE(SUBSTRING(CONVERT(CHAR(10),");
        sbSql.AppendLine(" CASE PaymTerm.NUMOFDAYS  WHEN 30 THEN DATEADD(m,1,dateadd(dd,-(day(CUSTINVOICEJOUR.INVOICEDATE)-1),CUSTINVOICEJOUR.INVOICEDATE) )");
        sbSql.AppendLine(" WHEN 60 THEN DATEADD(m,2,dateadd(dd,-(day(CUSTINVOICEJOUR.INVOICEDATE)-1),CUSTINVOICEJOUR.INVOICEDATE) )");
        sbSql.AppendLine(" WHEN 90 THEN DATEADD(m,3,dateadd(dd,-(day(CUSTINVOICEJOUR.INVOICEDATE)-1),CUSTINVOICEJOUR.INVOICEDATE) )");
        sbSql.AppendLine("  END ");

        sbSql.AppendLine("    ,3),4,5),' ','-')");
       // sbSql.AppendLine(" ,5),1,8),'','-')");
        sbSql.AppendLine("   [Due Date]");

        sbSql.AppendLine(" ,CUSTINVOICEJOUR.LedgerVoucher [Voucher No]");
        sbSql.AppendLine(" ,CUSTINVOICEJOUR.INVOICEID [Invoice No]");
        sbSql.AppendLine(",CUSTINVOICEJOUR.INVOICEDATE [Invoice Date]");
        sbSql.AppendLine(" ,CURRENCY.CURRENCYCODEISO [Curr]");
        sbSql.AppendLine("  ,CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNT END AmtCurr");
        sbSql.AppendLine(" ,CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNTMST END [AmountTHB]");
       // sbSql.AppendLine(" ,CASE WHEN SALESTABLE.ECL_SALESCOMERCIAL = 1 THEN 'COMERCIAL' ELSE 'NOCOMERCIAL' END Remark");
        sbSql.AppendLine(" FROM SALESTABLE");

        sbSql.AppendLine("INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine("INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
        sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
        sbSql.AppendLine(" AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");

        sbSql.AppendLine("INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
        sbSql.AppendLine(" INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
        sbSql.AppendLine(" INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");

        sbSql.AppendLine(" AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
        sbSql.AppendLine(" AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT");
        sbSql.AppendLine(" AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");
        sbSql.AppendLine("   WHERE SALESTABLE.DATAAREAID='HOYA'");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "'");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.INVOICEACCOUNT='" + ARReconcileOBJ.InvoiceAccount + "'");
        sbSql.AppendLine("      AND CURRENCY.CURRENCYCODEISO='" + ARReconcileOBJ.CurrencyISO + "'");
        sbSql.AppendLine(" AND ECL_SALESCOMERCIAL = 1 AND INVOICEDATE<=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103) )as ARReconcil  ");
        sbSql.AppendLine(" WHERE  AmtCurr != 0 OR [AmountTHB] !=0");
        sbSql.AppendLine(" GROUP BY  ARReconcil.[Due Date] ,ARReconcil.[Voucher No],ARReconcil.[Invoice No],ARReconcil.[Invoice Date],ARReconcil.Curr");
        sbSql.AppendLine("WITH ROLLUP HAVING NOT ARReconcil.Curr IS NULL OR ARReconcil.[Voucher No] iS NULL AND NOT ARReconcil.[Due Date] IS NULL");
        sbSql.AppendLine("ORDER BY GROUPING(ARReconcil.[Due Date]),ARReconcil.[Due Date]");
        sbSql.AppendLine(",GROUPING(ARReconcil.[Voucher No]),ARReconcil.[Voucher No]");
        sbSql.AppendLine(",GROUPING(ARReconcil.[Invoice Date]),ARReconcil.[Invoice Date]");
        sbSql.AppendLine(",GROUPING(ARReconcil.Curr),ARReconcil.Curr");

        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

       return rs;

        }

    public ADODB.Recordset getInvoiceDetail3New(ARReconcileOBJ ARReconcileOBJ)
    {
        StringBuilder sbSql = new StringBuilder();
        DateTime d = DateTime.Now;

        sbSql.AppendLine(" SELECT");
        sbSql.AppendLine("  CASE WHEN [Due Date] IS NULL AND Curr IS NULL THEN 'GRAND TOTAL' ELSE CASE WHEN Curr IS NULL THEN [Due Date] ELSE '' END END   [Due Date] ");
        sbSql.AppendLine("  ,CASE WHEN [Voucher No] IS NULL AND NOT [Due Date] IS NULL THEN 'TOTAL' ELSE [Voucher No] END   [Voucher No]");
        sbSql.AppendLine("  ,[Invoice No]");
        sbSql.AppendLine(" ,[Invoice Date]");
        sbSql.AppendLine(" ,Curr");
        sbSql.AppendLine(" ,SUM(AmtCurr),SUM([Baht])");
        sbSql.AppendLine(" ,Remark");



        sbSql.AppendLine(" FROM ( ");
        sbSql.AppendLine("  SELECT ");
        sbSql.AppendLine("      REPLACE(SUBSTRING(CONVERT(CHAR(10),");
        sbSql.AppendLine("          CASE PaymTerm.NUMOFDAYS  WHEN 30 THEN DATEADD(m,1,CUSTINVOICEJOUR.INVOICEDATE)");
        sbSql.AppendLine("                              WHEN 60 THEN DATEADD(m,2,CUSTINVOICEJOUR.INVOICEDATE)");
        sbSql.AppendLine("                              WHEN 90 THEN DATEADD(m,3,CUSTINVOICEJOUR.INVOICEDATE)");
        sbSql.AppendLine("          END ");
        sbSql.AppendLine("      ,6),4,6),' ','-')");
        sbSql.AppendLine("  [Due Date]");

        sbSql.AppendLine("  ,CUSTINVOICEJOUR.LedgerVoucher [Voucher No]");
        sbSql.AppendLine("  ,CUSTINVOICEJOUR.INVOICEID [Invoice No]");
        sbSql.AppendLine("  ,CUSTINVOICEJOUR.INVOICEDATE [Invoice Date]");
        sbSql.AppendLine("  ,CURRENCY.CURRENCYCODEISO [Curr]");
        sbSql.AppendLine("  ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNT END) AmtCurr");
        sbSql.AppendLine("  ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNTMST END) [Baht]");

        sbSql.AppendLine("  ,CASE WHEN SALESTABLE.ECL_SALESCOMERCIAL = 1 THEN 'COMERCIAL' ELSE 'NOCOMERCIAL' END [Remark]");
        sbSql.AppendLine("  ,SALESTABLE.PAYMENT");
        sbSql.AppendLine("  ,PaymTerm.NUMOFDAYS");

        sbSql.AppendLine("  FROM SALESTABLE");

        sbSql.AppendLine("  INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine("  INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("  INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("      AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("  INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
        sbSql.AppendLine("  INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
        sbSql.AppendLine("  INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("      AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
        sbSql.AppendLine("      AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("  LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT");
        sbSql.AppendLine("      AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");
        sbSql.AppendLine("  WHERE SALESTABLE.DATAAREAID='HOYA'");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "'");
        sbSql.AppendLine("      AND CUSTINVOICEJOUR.INVOICEACCOUNT='" + ARReconcileOBJ.InvoiceAccount + "'");
        sbSql.AppendLine("      AND CURRENCY.CURRENCYCODEISO='" + ARReconcileOBJ.CurrencyISO + "'");
        sbSql.AppendLine(" AND ECL_SALESCOMERCIAL = 1 ");
        sbSql.AppendLine(" AND ECL_SALESCOMERCIAL = 1 AND INVOICEDATE<=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103) ");
        
        sbSql.AppendLine("  GROUP BY");
        sbSql.AppendLine("      PaymTerm.NUMOFDAYS,CUSTINVOICEJOUR.LedgerVoucher,CUSTINVOICEJOUR.INVOICEID");
        sbSql.AppendLine("      ,CUSTINVOICEJOUR.INVOICEDATE,CURRENCY.CURRENCYCODEISO,SALESTABLE.PAYMENT,SALESTABLE.ECL_SALESCOMERCIAL");
        sbSql.AppendLine(" ) AS ARReconcile");

        sbSql.AppendLine(" WHERE  AmtCurr != 0 OR [Baht] !=0");
        sbSql.AppendLine(" GROUP BY [Due Date],[Voucher No],[Invoice No],[Invoice Date],Curr,Remark WITH ROLLUP");
        sbSql.AppendLine("HAVING NOT Remark IS NULL OR ( [Voucher No] IS NULL AND [Invoice No] IS NULL)");
        
       

        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;

    }

    public DataTable getInvoiceAccountCurr(ARReconcileOBJ ARReconcileOBJ)
    {
        StringBuilder sbSql = new StringBuilder();
        sbSql.AppendLine("  SELECT  [INVOICEACCOUNT],NAME,CURR [Curr]");
        sbSql.AppendLine("  FROM (");
        sbSql.AppendLine(" SELECT NAME ,[Curr],[INVOICEACCOUNT]");
        sbSql.AppendLine("FROM(");
        sbSql.AppendLine(" SELECT CUSTTABLE.ACCOUNTNUM");
        sbSql.AppendLine(" ,CASE WHEN CUSTINVOICEJOUR.INVOICEACCOUNT LIKE 'ARAO%' THEN CASE WHEN  CUSTINVOICEJOUR.PRINTMGMTSITEID ='GMO' THEN DIRPARTYTABLE.NAME +' (KATA)' ELSE DIRPARTYTABLE.NAME END ELSE DIRPARTYTABLE.NAME END [NAME]");
        sbSql.AppendLine("  ,CASE PaymTerm.NUMOFDAYS  WHEN 30 THEN DATEADD(m,1,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
        sbSql.AppendLine("  WHEN 60 THEN DATEADD(m,2,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
        sbSql.AppendLine("WHEN 90 THEN DATEADD(m,3,CONVERT(DATETIME,CONVERT(CHAR(4),CUSTINVOICEJOUR.INVOICEDATE,12)+'01',12))");
        sbSql.AppendLine(" END [DueDate]");
        sbSql.AppendLine(",CURRENCY.CURRENCYCODEISO [CURR] ");
        sbSql.AppendLine(" ,CUSTINVOICEJOUR.INVOICEACCOUNT");
        sbSql.AppendLine(" ,SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNT END) [InvoiceAmt] ");
        sbSql.AppendLine(",SUM(CASE WHEN SALESTABLE.PAYMENT='NOCOM' THEN 0 ELSE CUSTINVOICEJOUR.INVOICEAMOUNTMST END) [InvoiceAmtBHT]");
        sbSql.AppendLine(" FROM SALESTABLE");
        sbSql.AppendLine(" INNER JOIN CUSTINVOICEJOUR ON SALESTABLE.SALESID=CUSTINVOICEJOUR.SALESID");
        sbSql.AppendLine("AND SALESTABLE.DATAAREAID=CUSTINVOICEJOUR.DATAAREAID");
        sbSql.AppendLine("INNER JOIN CUSTTRANS ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
        sbSql.AppendLine("INNER JOIN CURRENCY ON CUSTINVOICEJOUR.CURRENCYCODE=CURRENCY.CURRENCYCODE");
        sbSql.AppendLine("INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
        sbSql.AppendLine("AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("LEFT OUTER JOIN PaymTerm ON PaymTerm.PAYMTERMID=SALESTABLE.PAYMENT");
        sbSql.AppendLine("AND PaymTerm.DATAAREAID=SALESTABLE.DATAAREAID");


        sbSql.AppendLine(" WHERE SALESTABLE.DATAAREAID='HOYA' AND CUSTINVOICEJOUR.PRINTMGMTSITEID='" + ARReconcileOBJ.Factory + "'");
       
        sbSql.AppendLine(" AND CUSTINVOICEJOUR.INVOICEDATE <= CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103)");
        sbSql.AppendLine("AND ECL_SALESCOMERCIAL = 1");

        sbSql.AppendLine(" GROUP BY CUSTTABLE.ACCOUNTNUM,NAME,CUSTINVOICEJOUR.INVOICEACCOUNT,CUSTINVOICEJOUR.PRINTMGMTSITEID");
        sbSql.AppendLine(",PaymTerm.NUMOFDAYS,CUSTINVOICEJOUR.INVOICEDATE,CURRENCY.CURRENCYCODEISO,CUSTTABLE.ECL_REASON");
        sbSql.AppendLine(" ) ARDetail");
        sbSql.AppendLine("WHERE  [InvoiceAmt] != 0 OR [InvoiceAmtBHT] !=0");
        sbSql.AppendLine(" GROUP BY ACCOUNTNUM,NAME,DUEDATE,[Curr],INVOICEACCOUNT ");
        sbSql.AppendLine(" ) ARByDueDate");
        sbSql.AppendLine(" GROUP BY NAME,CURR,INVOICEACCOUNT");

        DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
        return dt;
    }





    public ADODB.Recordset getInvoiceDetail4New(ARReconcileOBJ ARReconcileOBJ)
    {
        StringBuilder sbSql = new StringBuilder();

        try
        {

            sbSql.AppendLine("select  ");
            sbSql.AppendLine(" CUSTTRANS.ACCOUNTNUM [ACCOUNT]");
            sbSql.AppendLine(",CASE WHEN DIRPARTYTABLE.NAME IS NULL AND NOT CUSTTRANS.ACCOUNTNUM IS NULL  THEN  'TOTAL' ELSE DIRPARTYTABLE.NAME END [Customer ACC]");
            sbSql.AppendLine(",CUSTTRANS.TRANSDATE");
            sbSql.AppendLine(",CUSTINVOICEJOUR.INVOICEID [InvoiceID]");
            sbSql.AppendLine(",CUSTTRANS.VOUCHER [voucher]");
            sbSql.AppendLine(",CUSTTRANS.TXT [Des]");
            sbSql.AppendLine(",CUSTTRANS.CURRENCYCODE [CURR]");
            sbSql.AppendLine(",SUM(CUSTTRANS.AMOUNTCUR)[AmountCurr]");
            sbSql.AppendLine(",CASE WHEN NOT DIRPARTYTABLE.NAME IS NULL THEN SUM(CUSTTRANS.EXCHRATE/100) ELSE NULL END [Rate]");
            sbSql.AppendLine(",SUM(CUSTTRANS.AMOUNTMST) [AmountMstr]");



            sbSql.AppendLine(" from CUSTTRANS");
            sbSql.AppendLine("INNER JOIN CUSTINVOICEJOUR ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
            sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
            sbSql.AppendLine("AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
            sbSql.AppendLine("INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
            sbSql.AppendLine("INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
            sbSql.AppendLine("AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
            sbSql.AppendLine("AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");

            sbSql.AppendLine(" LEFT OUTER JOIN (");
            sbSql.AppendLine("SELECT DIMENSIONATTRIBUTEVALUESET,ECL_SHORTNAME FROM DIMENSIONATTRIBUTEVALUESETITEM");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.RECID=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine(" INNER JOIN DimensionFinancialTag ON DimensionFinancialTag.recid=DimensionAttributeValue.ENTITYINSTANCE");
            sbSql.AppendLine(" INNER JOIN DimensionAttribute ON DimensionAttribute.recid=DimensionAttributeValue.DIMENSIONATTRIBUTE");
            sbSql.AppendLine("WHERE DIMENSIONATTRIBUTE.NAME = 'D1_Factory'");
            sbSql.AppendLine(") Factory ON CUSTTRANS.DEFAULTDIMENSION=Factory.DIMENSIONATTRIBUTEVALUESET");

            sbSql.AppendLine(" WHERE ECL_SHORTNAME = '" + ARReconcileOBJ.Factory + "'");
            sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEACCOUNT LIKE '" + ARReconcileOBJ.InvoiceAccount + "%'");
            sbSql.AppendLine("AND TRANSTYPE = 0 ");
            sbSql.AppendLine("AND  CUSTTRANS.TRANSDATE<=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103) ");
            sbSql.AppendLine("AND (CUSTTRANS.AMOUNTCUR) > 0");

            sbSql.AppendLine("  GROUP BY CUSTTRANS.ACCOUNTNUM ,DIRPARTYTABLE.NAME ,CUSTTRANS.CURRENCYCODE,CUSTTRANS.TRANSDATE,");
            sbSql.AppendLine("CUSTINVOICEJOUR.INVOICEID ,CUSTTRANS.VOUCHER ,CUSTTRANS.TXT WITH ROLLUP");
            sbSql.AppendLine("HAVING NOT CUSTTRANS.TXT IS NULL OR DIRPARTYTABLE.NAME IS NULL  AND NOT CUSTTRANS.ACCOUNTNUM IS NULL");
  
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
            return null;

        } //end Try

        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;

    }

    public DataTable getInvoiceAccountNum(ARReconcileOBJ ARReconcileOBJ)
    {
        StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine("SELECT * FROM (");


        sbSql.AppendLine("select  ");
        sbSql.AppendLine("SUBSTRING(CUSTTRANS.ACCOUNTNUM, 1, 4) ACCOUNTNUMM");


        sbSql.AppendLine(" from CUSTTRANS");
        sbSql.AppendLine("INNER JOIN CUSTINVOICEJOUR ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
        sbSql.AppendLine("INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
        sbSql.AppendLine("AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");

        sbSql.AppendLine(" LEFT OUTER JOIN (");
        sbSql.AppendLine("SELECT DIMENSIONATTRIBUTEVALUESET,ECL_SHORTNAME FROM DIMENSIONATTRIBUTEVALUESETITEM");
        sbSql.AppendLine("INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.RECID=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
        sbSql.AppendLine(" INNER JOIN DimensionFinancialTag ON DimensionFinancialTag.recid=DimensionAttributeValue.ENTITYINSTANCE");
        sbSql.AppendLine(" INNER JOIN DimensionAttribute ON DimensionAttribute.recid=DimensionAttributeValue.DIMENSIONATTRIBUTE");
        sbSql.AppendLine("WHERE DIMENSIONATTRIBUTE.NAME = 'D1_Factory'");
        sbSql.AppendLine(") Factory ON CUSTTRANS.DEFAULTDIMENSION=Factory.DIMENSIONATTRIBUTEVALUESET");

        sbSql.AppendLine(" WHERE ECL_SHORTNAME = '" + ARReconcileOBJ.Factory + "'");
        sbSql.AppendLine("AND TRANSTYPE = 0 ");
        sbSql.AppendLine("AND  CUSTTRANS.TRANSDATE<=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103) ");
        sbSql.AppendLine("AND (CUSTTRANS.AMOUNTCUR) > 0");

        sbSql.AppendLine(" GROUP BY CUSTTRANS.ACCOUNTNUM ) as ACCOUNTNUM ");
        sbSql.AppendLine(" GROUP by  ACCOUNTNUM.ACCOUNTNUMM");


        DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
        return dt;
    }


    public ADODB.Recordset getInvoiceAccountCURR(ARReconcileOBJ ARReconcileOBJ)
    {
        StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine("select  ");
        sbSql.AppendLine("CUSTTRANS.CURRENCYCODE [Curr]");


        sbSql.AppendLine(" from CUSTTRANS");
        sbSql.AppendLine("INNER JOIN CUSTINVOICEJOUR ON CUSTINVOICEJOUR.INVOICEID=CUSTTRANS.INVOICE");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEACCOUNT=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEDATE=CUSTTRANS.TRANSDATE");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN CUSTTABLE ON CUSTTABLE.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CUSTTABLE.DATAAREAID=CUSTTRANS.DATAAREAID");
        sbSql.AppendLine("INNER JOIN DIRPARTYTABLE ON CUSTTABLE.PARTY=DIRPARTYTABLE.RECID");
        sbSql.AppendLine("INNER JOIN CustTransOpen ON CustTransOpen.ACCOUNTNUM=CUSTTRANS.ACCOUNTNUM");
        sbSql.AppendLine("AND CustTransOpen.REFRECID=CUSTTRANS.RECID");
        sbSql.AppendLine("AND CustTransOpen.DATAAREAID=CUSTTRANS.DATAAREAID");

        sbSql.AppendLine(" LEFT OUTER JOIN (");
        sbSql.AppendLine("SELECT DIMENSIONATTRIBUTEVALUESET,ECL_SHORTNAME FROM DIMENSIONATTRIBUTEVALUESETITEM");
        sbSql.AppendLine("INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.RECID=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
        sbSql.AppendLine(" INNER JOIN DimensionFinancialTag ON DimensionFinancialTag.recid=DimensionAttributeValue.ENTITYINSTANCE");
        sbSql.AppendLine(" INNER JOIN DimensionAttribute ON DimensionAttribute.recid=DimensionAttributeValue.DIMENSIONATTRIBUTE");
        sbSql.AppendLine("WHERE DIMENSIONATTRIBUTE.NAME = 'D1_Factory'");
        sbSql.AppendLine(") Factory ON CUSTTRANS.DEFAULTDIMENSION=Factory.DIMENSIONATTRIBUTEVALUESET");

        sbSql.AppendLine(" WHERE ECL_SHORTNAME = '" + ARReconcileOBJ.Factory + "'");
        sbSql.AppendLine("AND CUSTINVOICEJOUR.INVOICEACCOUNT LIKE '" + ARReconcileOBJ.InvoiceAccount + "%'");
        sbSql.AppendLine("AND TRANSTYPE = 0 ");
        sbSql.AppendLine("AND  CUSTTRANS.TRANSDATE<=CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", ARReconcileOBJ.DateTo) + "',103) ");
        sbSql.AppendLine("AND (CUSTTRANS.AMOUNTCUR) > 0");

        sbSql.AppendLine(" GROUP BY CUSTTRANS.CURRENCYCODE ");



        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();

        rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

        return rs;

    }










    }

}
