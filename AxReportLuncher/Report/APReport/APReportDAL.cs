using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;

namespace NewVersion.Report.APReport
{
    class APReportDAL
    {
        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();

        public DataTable findVendor(string strSearchField,string strSearchValue)
        {
            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine(" SELECT AccountNum,Name FROM VendTable");
        sbSql.AppendLine(" INNER JOIN VendDirPartyTableView ON VendTable.Party=VendDirPartyTableView.Party");
        sbSql.AppendLine(" WHERE " + strSearchField + " LIKE '%" + strSearchValue + "%'");
        sbSql.AppendLine(" ORDER BY AccountNum");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public DataTable getVenderGroup(string strVendGroup)
        {
            StringBuilder sbSql = new StringBuilder();

                sbSql.AppendLine(" SELECT DISTINCT VendGroup AS VendGroup");
                sbSql.AppendLine(" FROM VendGroup");
                sbSql.AppendLine(" WHERE NOT(VendGroup IN ('" + strVendGroup + "'))");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public DataTable getVendor(APReportOBJ APReportOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine(" SELECT VENDTABLE.ACCOUNTNUM,VendDirPartyTableView.Name,VENDTABLE.VendGroup");
        sbSql.AppendLine(" FROM VENDTABLE");
        sbSql.AppendLine(" INNER JOIN VendTrans ON VENDTABLE.ACCOUNTNUM=VendTrans.ACCOUNTNUM");
        sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET=VENDTRANS.DefaultDimension");
        sbSql.AppendLine(" INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.RECID=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
        sbSql.AppendLine(" INNER JOIN DimensionAttribute ON DimensionAttribute.RECID=DimensionAttributeValue.DIMENSIONATTRIBUTE");
        sbSql.AppendLine(" INNER JOIN DIMENSIONFINANCIALTAG ON DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE=DIMENSIONFINANCIALTAG.VALUE");
        sbSql.AppendLine(" INNER JOIN VendDirPartyTableView ON VendTable.Party=VendDirPartyTableView.Party");
        sbSql.AppendLine(" INNER JOIN VendGroup ON VENDTABLE.VendGroup=VendGroup.VendGroup ");
        sbSql.AppendLine(" WHERE DimensionAttribute.name='D1_Factory' and convert(char(10),VENDTRANS.CLOSED,103)='01/01/1900'");
        sbSql.AppendLine(" AND VendTrans.PROMISSORYNOTESTATUS=6 AND BLOCKED=0");

        sbSql.AppendLine(" AND  VENDTRANS.TransDate <= CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", APReportOBJ.DateTo) + "',103) ");

        //sbSql.AppendLine(" AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", SalesReturnOBJ.DateTo) + "',103)");
        if (APReportOBJ.vendercode != "")
        {
            sbSql.AppendLine(" AND VENDTRANS.ACCOUNTNUM='" + APReportOBJ.vendercode + "'");
        }

        if (APReportOBJ.venderGroup != "")
        {
            APReportOBJ.venderGroup = APReportOBJ.venderGroup.Replace(",", "','");
            sbSql.AppendLine(" AND VENDTABLE.VendGroup IN ('" + APReportOBJ.venderGroup + "')");
        }

        if (APReportOBJ.Factory != "")
        {
            sbSql.AppendLine(" AND ECL_SHORTNAME ='" + APReportOBJ.Factory + "'");
        }


        sbSql.AppendLine(" GROUP BY VENDTABLE.ACCOUNTNUM , VendDirPartyTableView.Name,VENDTABLE.VendGroup");
        sbSql.AppendLine(" ORDER BY VENDTABLE.VendGroup, AccountNum");

            DataTable dt = QueryDAL.QueryDataTable(sbSql.ToString());
            return dt;
        }

        public ADODB.Recordset getCurrRate(APReportOBJ APReportOBJ)
        {
            DateTime dt = new DateTime(APReportOBJ.DateTo.Year,APReportOBJ.DateTo.Month,1).AddMonths(1);
            StringBuilder sbSql = new StringBuilder();

        sbSql.AppendLine(" SELECT EXCHANGERATE FROM (");
        sbSql.AppendLine(" SELECT 1 EXCHANGERATE,'THB' CURRENCYCODEISO");
        sbSql.AppendLine(" UNION ALL");
        sbSql.AppendLine(" SELECT ExchangeRate.EXCHANGERATE/100 EXCHANGERATE,CURRENCY.CURRENCYCODEISO");
        sbSql.AppendLine(" FROM EXCHANGERATECURRENCYPAIR");
        sbSql.AppendLine(" INNER JOIN EXCHANGERATETYPE ON EXCHANGERATECURRENCYPAIR.EXCHANGERATETYPE=EXCHANGERATETYPE.RECID");
        sbSql.AppendLine(" INNER JOIN ExchangeRate ON ExchangeRate.EXCHANGERATECURRENCYPAIR=EXCHANGERATECURRENCYPAIR.RECID");
        sbSql.AppendLine(" INNER JOIN CURRENCY ON EXCHANGERATECURRENCYPAIR.FROMCURRENCYCODE=CURRENCY.CURRENCYCODE");
        sbSql.AppendLine(" WHERE EXCHANGERATETYPE.NAME='Default' AND EXCHANGERATECURRENCYPAIR.FROMCURRENCYCODE IN ('JPS','USS','HKS','SGS')");

        sbSql.AppendLine(" AND ExchangeRate.VALIDFROM=CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103)");
        sbSql.AppendLine(" AND ExchangeRate.VALIDTO=CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", dt) + "',103))RATE ");

        sbSql.AppendLine(" ORDER BY CASE RATE.CURRENCYCODEISO WHEN 'JPY' THEN 1");
        sbSql.AppendLine("                                  WHEN 'USD' THEN 2");
        sbSql.AppendLine("                                  WHEN 'THB' THEN 3");
        sbSql.AppendLine("                                  WHEN 'HKD' THEN 4");
        sbSql.AppendLine("                                  WHEN 'SGD' THEN 5 END");

        ADODB.Recordset rs = new ADODB.Recordset();
        ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
        ADODBConnection.Open();
        object ret = null;
        rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

        return rs;
        }

        public ADODB.Recordset getAPDueDate(APReportOBJ APReportOBJ)
        {
            //DateTime dt = new DateTime(APReportOBJ.DateTo.Year, APReportOBJ.DateTo.Month, 1).AddMonths(1);
            StringBuilder sbSql = new StringBuilder();

            try
            {
            sbSql.AppendLine(" SELECT [InvoiceDate],[Invoice No],VOUCHER,[Receiving Date],[AWB Date],CURRENCYCODEISO,AMOUNTCUR,EXCHRATE,AMOUNTMST,[DUEDATE] FROM (");
            sbSql.AppendLine(" SELECT CASE WHEN convert(char(10),VENDTRANS.DOCUMENTDATE,103)='01/01/1900' THEN '' ELSE CONVERT(char(10),VENDTRANS.DOCUMENTDATE,103) END [InvoiceDate]");
            sbSql.AppendLine(" ,VENDTRANS.INVOICE [Invoice No]");
            sbSql.AppendLine(" ,VENDTRANS.VOUCHER");
            sbSql.AppendLine(" ,VENDTRANS.TransDate [Receiving Date]");         
            sbSql.AppendLine(" ,CASE WHEN CONVERT(CHAR(10),VENDINVOICEINFOTABLE.HOYA_AWBDATE,103)='01/01/1900' THEN '' ELSE CONVERT(CHAR(10),VENDINVOICEINFOTABLE.HOYA_AWBDATE,103) END [AWB Date]");
            sbSql.AppendLine(" ,VENDTRANS.CURRENCYCODE");
            sbSql.AppendLine(" ,CURRENCY.CURRENCYCODEISO");         
            sbSql.AppendLine(" ,VENDTRANSOPEN.AMOUNTCUR*-1 AMOUNTCUR");
            sbSql.AppendLine(" ,VENDTRANS.EXCHRATE/100 EXCHRATE");         
            sbSql.AppendLine(" ,VENDTRANSOPEN.AMOUNTMST*-1 AMOUNTMST");         
            sbSql.AppendLine(" ,CASE WHEN VENDTRANS.TRANSTYPE=0 THEN DATEADD(dd,PAYMTERM.NumOfDays,VENDTRANS.TransDate) ELSE DATEADD(dd,PAYMTERM.NumOfDays,VENDINVOICEINFOTABLE.HOYA_AWBDATE) END [DUEDATE]");
            sbSql.AppendLine(" FROM VENDTABLE");
            sbSql.AppendLine(" INNER JOIN PAYMTERM ON VENDTABLE.PAYMTERMID=PAYMTERM.PAYMTERMID");
            sbSql.AppendLine(" INNER JOIN VendTrans ON VENDTABLE.ACCOUNTNUM=VendTrans.ACCOUNTNUM");
            sbSql.AppendLine(" INNER JOIN CURRENCY ON VENDTRANS.CURRENCYCODE=CURRENCY.CURRENCYCODE");
            sbSql.AppendLine(" INNER JOIN VENDTRANSOPEN ON  VENDTRANS.ACCOUNTNUM= VENDTRANSOPEN.ACCOUNTNUM and VendTrans.RECID=VENDTRANSOPEN.REFRECID");
            sbSql.AppendLine(" INNER JOIN DIRPARTYTABLE ON VENDTABLE.PARTY=DIRPARTYTABLE.RECID");
            sbSql.AppendLine(" LEFT OUTER JOIN VENDINVOICEJOUR ON VENDTRANS.VOUCHER=VENDINVOICEJOUR.LEDGERVOUCHER");
            sbSql.AppendLine("  AND VENDTRANS.ACCOUNTNUM=VENDINVOICEJOUR.INVOICEACCOUNT");
            sbSql.AppendLine("  AND VENDTRANS.TRANSDATE=VENDINVOICEJOUR.INVOICEDATE");
                 sbSql.AppendLine(" LEFT OUTER JOIN (SELECT NUM,INVOICEACCOUNT,VENDINVOICESAVESTATUS,MAX(DOCUMENTDATE) DOCUMENTDATE,MAX(HOYA_AWBDATE) HOYA_AWBDATE,DATAAREAID ");
            sbSql.AppendLine("                  FROM VENDINVOICEINFOTABLE GROUP BY NUM,INVOICEACCOUNT,VENDINVOICESAVESTATUS,DATAAREAID) VENDINVOICEINFOTABLE");
            sbSql.AppendLine("                              ON VENDINVOICEJOUR.INVOICEID=VENDINVOICEINFOTABLE.NUM");
            sbSql.AppendLine("                              AND VENDINVOICEJOUR.INVOICEACCOUNT=VENDINVOICEINFOTABLE.INVOICEACCOUNT");
            sbSql.AppendLine("                              AND VENDINVOICEJOUR.DATAAREAID=VENDINVOICEINFOTABLE.DATAAREAID");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET=VENDTRANS.DefaultDimension");
            sbSql.AppendLine(" INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.RECID=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine(" INNER JOIN DimensionAttribute ON DimensionAttribute.RECID=DimensionAttributeValue.DIMENSIONATTRIBUTE");
            sbSql.AppendLine(" INNER JOIN DIMENSIONFINANCIALTAG ON DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE=DIMENSIONFINANCIALTAG.VALUE");       
            sbSql.AppendLine(" WHERE DimensionAttribute.name='D1_Factory' and convert(char(10),VENDTRANS.CLOSED,103)='01/01/1900'");
            sbSql.AppendLine(" AND (VENDINVOICEINFOTABLE.VENDINVOICESAVESTATUS=0 OR VENDINVOICEINFOTABLE.VENDINVOICESAVESTATUS IS NULL)");

            sbSql.AppendLine("  AND VENDTRANS.TransDate <= CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", APReportOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" AND VendTrans.PROMISSORYNOTESTATUS=6");

            if (APReportOBJ.Factory != "")
            {
                sbSql.AppendLine(" AND ECL_SHORTNAME IN ('" + APReportOBJ.Factory + "')");
            }

            if (APReportOBJ.vendercode != "")
            {
                sbSql.AppendLine(" AND VENDTRANS.ACCOUNTNUM='" + APReportOBJ.vendercode + "'");
            }

            sbSql.AppendLine(" ) AS REPORT ORDER BY REPORT.DUEDATE,REPORT.CURRENCYCODE,REPORT.VOUCHER,REPORT.[Invoice No]");




            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();
            object ret = null;
            rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

            return rs;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

        }


        public ADODB.Recordset getAPReconcile(APReportOBJ APReportOBJ)
        {
            //DateTime dt = new DateTime(APReportOBJ.DateTo.Year, APReportOBJ.DateTo.Month, 1).AddMonths(1);
            StringBuilder sbSql = new StringBuilder();

            try
            {
             
            sbSql.AppendLine(" SELECT VendDirPartyTableView.NAME, VENDTRANS.ACCOUNTNUM");
            sbSql.AppendLine(" ,CURRENCY.CURRENCYCODEISO");
            sbSql.AppendLine(" ,SUM(VENDTRANSOPEN.AMOUNTCUR*-1) amtcurr");
            sbSql.AppendLine(" ,SUM(VENDTRANSOPEN.AMOUNTMST*-1) amtbht");
            sbSql.AppendLine(" FROM VENDTABLE");
            sbSql.AppendLine(" INNER JOIN VendTrans ON VENDTABLE.ACCOUNTNUM=VendTrans.ACCOUNTNUM");
            sbSql.AppendLine(" INNER JOIN CURRENCY ON VENDTRANS.CURRENCYCODE=CURRENCY.CURRENCYCODE");
            sbSql.AppendLine(" INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM ON DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET=VENDTRANS.DefaultDimension");
            sbSql.AppendLine(" INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.RECID=DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine(" INNER JOIN DimensionAttribute ON DimensionAttribute.RECID=DimensionAttributeValue.DIMENSIONATTRIBUTE");
            sbSql.AppendLine(" INNER JOIN DIMENSIONFINANCIALTAG ON DIMENSIONATTRIBUTEVALUESETITEM.DISPLAYVALUE=DIMENSIONFINANCIALTAG.VALUE");
            sbSql.AppendLine(" INNER JOIN VENDTRANSOPEN ON  VENDTRANS.ACCOUNTNUM= VENDTRANSOPEN.ACCOUNTNUM and VendTrans.RECID=VENDTRANSOPEN.REFRECID");
            sbSql.AppendLine(" INNER JOIN VendDirPartyTableView ON VendTable.Party=VendDirPartyTableView.Party");
            sbSql.AppendLine(" ");
            sbSql.AppendLine(" ");
            sbSql.AppendLine(" WHERE DimensionAttribute.name='D1_Factory' and convert(char(10),VENDTRANS.CLOSED,103)='01/01/1900'");

            sbSql.AppendLine("  AND VENDTRANS.TransDate <= CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", APReportOBJ.DateTo) + "',103)");
            sbSql.AppendLine(" AND VendTrans.PROMISSORYNOTESTATUS=6");
            sbSql.AppendLine(" AND VENDGROUP='" + APReportOBJ.venderGroup + "'");


                if (APReportOBJ.Factory != "")
                {
                    sbSql.AppendLine(" AND ECL_SHORTNAME IN ('" + APReportOBJ.Factory + "')");
                }


            sbSql.AppendLine(" GROUP BY VENDTRANS.ACCOUNTNUM,VendDirPartyTableView.NAME,CURRENCY.CURRENCYCODEISO");
            sbSql.AppendLine(" ORDER BY VENDTRANS.ACCOUNTNUM,CURRENCY.CURRENCYCODEISO");



                ADODB.Recordset rs = new ADODB.Recordset();
                ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
                ADODBConnection.Open();
                object ret = null;
                rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

                return rs;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

        }


        public ADODB.Recordset getAPSummary(APReportOBJ APReportOBJ)
        {
            //DateTime dt = new DateTime(APReportOBJ.DateTo.Year, APReportOBJ.DateTo.Month, 1).AddMonths(1);
            StringBuilder sbSql = new StringBuilder();

            try
            {

            sbSql.AppendLine(" SELECT ACCOUNTNUM,NAME,CURRENCYCODE,SUM(AmtCurr)AmtCurr,SUM(VAT)VAT");
            sbSql.AppendLine(" ,SUM(AmtCurr)+SUM(VAT) AmtBht");
            sbSql.AppendLine(" ,SUM(WHT)WHT");
            sbSql.AppendLine(" ,SUM(AmtCurr)+SUM(VAT)+SUM(WHT) NetPay");
            sbSql.AppendLine(" FROM (");
            sbSql.AppendLine("  SELECT VENDINVOICEJOUR.INVOICEACCOUNT ACCOUNTNUM,VENDINVOICETRANS.INVOICEDATE");
            sbSql.AppendLine("  ,DIRPARTYTABLE.NAME,VENDTABLE.VENDGROUP");
            sbSql.AppendLine("  ,VENDINVOICEJOUR.CURRENCYCODE");
            sbSql.AppendLine("  ,SUM(VENDINVOICETRANS.LINEAMOUNT)-SUM(VENDINVOICETRANS.DISCAMOUNT) AmtCurr");
            sbSql.AppendLine("  ,VENDINVOICEJOUR.SUMTAX * VENDINVOICEJOUR.ExchRate / 100 VAT");
            sbSql.AppendLine("  ,ISNULL(VENDTRANSOPEN.ECL_WHTAXAMOUNT,0) WHT");
            sbSql.AppendLine("  ,INVENTDIM.INVENTSITEID");
            sbSql.AppendLine("  FROM VENDINVOICETRANS ");
            sbSql.AppendLine("  INNER JOIN INVENTDIM on INVENTDIM.INVENTDIMID=VENDINVOICETRANS.INVENTDIMID AND INVENTDIM.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine("  INNER JOIN VENDINVOICEJOUR ON VENDINVOICEJOUR.PURCHID=VENDINVOICETRANS.PURCHID");
            sbSql.AppendLine("      AND VENDINVOICEJOUR.INVOICEID=VENDINVOICETRANS.INVOICEID");
            sbSql.AppendLine("      AND VENDINVOICEJOUR.INVOICEDATE=VENDINVOICETRANS.INVOICEDATE");
            sbSql.AppendLine("      AND VENDINVOICEJOUR.NUMBERSEQUENCEGROUP=VENDINVOICETRANS.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("      AND VENDINVOICEJOUR.INTERNALINVOICEID=VENDINVOICETRANS.INTERNALINVOICEID");
            sbSql.AppendLine("      AND VENDINVOICEJOUR.DATAAREAID=VENDINVOICETRANS.DATAAREAID");
            sbSql.AppendLine("  INNER JOIN VENDTRANS ON VENDTRANS.VOUCHER=VENDINVOICEJOUR.LEDGERVOUCHER");
            sbSql.AppendLine("      AND VENDTRANS.ACCOUNTNUM=VENDINVOICEJOUR.INVOICEACCOUNT");
            sbSql.AppendLine("      AND VENDTRANS.TRANSDATE=VENDINVOICEJOUR.INVOICEDATE");
            sbSql.AppendLine("      AND VENDTRANS.DATAAREAID=VENDINVOICEJOUR.DATAAREAID");
            sbSql.AppendLine("  INNER JOIN VENDTABLE on VENDTABLE.ACCOUNTNUM=VENDTRANS.ACCOUNTNUM AND VENDTABLE.DATAAREAID=VENDTRANS.DATAAREAID");
            sbSql.AppendLine("  INNER JOIN VendGroup ON VendGroup.VENDGROUP=VENDTABLE.VENDGROUP AND VendGroup.DATAAREAID=VENDTABLE.DATAAREAID");
            sbSql.AppendLine("  INNER JOIN DIRPARTYTABLE ON DIRPARTYTABLE.RECID=VENDTABLE.PARTY");
            sbSql.AppendLine("  LEFT OUTER JOIN VENDTRANSOPEN ON VENDTRANSOPEN.REFRECID=VENDTRANS.RECID AND VENDTRANSOPEN.DATAAREAID=VENDTRANS.DATAAREAID");
            sbSql.AppendLine("  WHERE VENDINVOICETRANS.DATAAREAID='hoya' AND NOT(VENDINVOICETRANS.NUMBERSEQUENCEGROUP LIKE '%-HO%' OR VENDINVOICETRANS.NUMBERSEQUENCEGROUP LIKE '%-CHO%')");
            sbSql.AppendLine("  GROUP BY VENDINVOICEJOUR.INVOICEACCOUNT,DIRPARTYTABLE.NAME,VENDTABLE.VENDGROUP,VENDINVOICEJOUR.CURRENCYCODE,VENDINVOICEJOUR.SUMTAX,VENDINVOICEJOUR.ExchRate");
            sbSql.AppendLine("  ,VENDINVOICETRANS.PURCHID");
            sbSql.AppendLine("  ,VENDINVOICETRANS.INVOICEID");
            sbSql.AppendLine("  ,VENDINVOICETRANS.INVOICEDATE");
            sbSql.AppendLine("  ,VENDINVOICETRANS.NUMBERSEQUENCEGROUP");
            sbSql.AppendLine("  ,VENDINVOICETRANS.INTERNALINVOICEID");
            sbSql.AppendLine("  ,VENDTRANSOPEN.ECL_WHTAXAMOUNT");
            sbSql.AppendLine("  ,INVENTDIM.INVENTSITEID");
            sbSql.AppendLine(" ) AS Detl");
            sbSql.AppendLine(" WHERE ");

            sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(DATETIME,'" + String.Format("{0:dd/MM/yyyy}", APReportOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine("      AND CONVERT(date,'" + String.Format("{0:dd/MM/yyyy}", APReportOBJ.DateTo) + "',103)");


            if (APReportOBJ.Factory != "")
            {
                sbSql.AppendLine(" AND INVENTSITEID='" + APReportOBJ.Factory + "'");
            }
            if (APReportOBJ.vendercode != "")
            {
                sbSql.AppendLine(" AND ACCOUNTNUM='" + APReportOBJ.vendercode + "'");
            }
            if (APReportOBJ.venderGroup != "")
            {
                sbSql.AppendLine(" AND VENDGROUP IN ('" + APReportOBJ.venderGroup.Replace(",", "','") + "')");
            }

            sbSql.AppendLine(" GROUP BY ACCOUNTNUM,NAME,CURRENCYCODE");

                ADODB.Recordset rs = new ADODB.Recordset();
                ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
                ADODBConnection.Open();
                object ret = null;
                rs = ADODBConnection.Execute(sbSql.ToString(), out ret, 0);

                return rs;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

        }


    }// end class
}
