using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NewVersion.Report.POReamain
{
    class PORemainDAL
    {
        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();


        public ADODB.Recordset getPORemain(PORemainOBJ PORemainOBJ)
        {
            StringBuilder sbSql = new StringBuilder();

            String strFac;
            if (PORemainOBJ.Factory == "GMO")
            {
                 strFac = "MO";
            }
            else
            {
                 strFac = PORemainOBJ.Factory;
            }

            sbSql.AppendLine("select PurchTable.createdDateTime [PO DATE]");
            sbSql.AppendLine(",purchtable.PurchID [Purch ID]");
            sbSql.AppendLine(",Purchline.PurchReqID");
            sbSql.AppendLine(",Purchline.Name");
            sbSql.AppendLine(",D3_DIMENSIONATTRIBUTEVALUESETITEM.DisplayValue [SEC]");
            sbSql.AppendLine(",DimensionAttributeValueSetItem.DisplayValue [Line Factory]");
            sbSql.AppendLine(",Hfactory.DisplayValue [Head Factory]");
            sbSql.AppendLine(",Purchline.ItemID [Item ID]");
            sbSql.AppendLine(",Purchline.Name [Name]");
            sbSql.AppendLine(",Purchline.PurchQty [Qty]");
            sbSql.AppendLine(",Purchline.PurchUnit [Unit]");
            sbSql.AppendLine(",purchtable.CurrencyCode [Cur]");
            sbSql.AppendLine(",Purchline.PurchPrice [Purch Price]");
            sbSql.AppendLine(",Purchline.DeliveryDate [ExpArrv]");
            sbSql.AppendLine(",purchtable.DlvTerm [Inco.Tm]");
            sbSql.AppendLine(",Payment  [Paym.Tm]");
            sbSql.AppendLine(",Purchline.ConfirmedDlv [Confirm]");
            sbSql.AppendLine(",Purchline.PurchQty*RemainPurchPhysical [Qty Rcpt]");
            sbSql.AppendLine(",RemainPurchPhysical [Qty Remain]");
            sbSql.AppendLine(",Purchline.PurchPrice*RemainPurchPhysical [Total Remain]");
            sbSql.AppendLine(",purchtable.ECL_ReingishoNo [Ringi No]");
            sbSql.AppendLine(",purchtable.ECL_QuotationNo  [Quot No]");
            sbSql.AppendLine(",purchtable.ECL_PurchRemark  [Remark]");
            sbSql.AppendLine(",purchtable.FreightZone [Orderer]");
            sbSql.AppendLine(",CASE purchtable.PurchStatus ");
            sbSql.AppendLine("WHEN 0 THEN 'NONE'");
            sbSql.AppendLine("WHEN 1 THEN 'BACKORDER'");
            sbSql.AppendLine("WHEN 2 THEN 'RECEIVED'");
            sbSql.AppendLine("WHEN 3 THEN 'INVOICED'");
            sbSql.AppendLine("WHEN 4 THEN 'CANCELED'");
            sbSql.AppendLine("END  [PURCH STATUS]");

            sbSql.AppendLine("from PurchTable");
            sbSql.AppendLine("INNER JOIN  purchline on purchtable.PurchID = purchline.Purchid ANd purchTable.DATAAREAID = purchline.DATAAREAID");
            sbSql.AppendLine("INNER JOIN  vendtable ON vendtable.AccountNum = PurchTable.OrderAccount");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValueSetItem Hfactory ON Hfactory.DimensionAttributeValueSet = Purchtable.DefaultDimension");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValue Hvalue ON Hvalue.Recid = Hfactory.DimensionAttributeValue");
            sbSql.AppendLine("INNER JOIN DimensionAttribute HAttribute ON HAttribute.Recid = Hvalue.DimensionAttribute");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValueSetItem ON DimensionAttributeValueSetItem.DimensionAttributeValueSet = PurchLine.DefaultDimension");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValue ON DimensionAttributeValue.Recid = DimensionAttributeValueSetItem.DimensionAttributeValue");
            sbSql.AppendLine("INNER JOIN DimensionAttribute ON DimensionAttribute.Recid = DimensionAttributeValue.DimensionAttribute");
            sbSql.AppendLine("INNER JOIN DIMENSIONATTRIBUTEVALUESETITEM D3_DIMENSIONATTRIBUTEVALUESETITEM ON D3_DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUESET=PurchLine.DefaultDimension");
            sbSql.AppendLine("INNER JOIN DimensionAttributeValue D3_DimensionAttributeValue ON D3_DimensionAttributeValue.RECID=D3_DIMENSIONATTRIBUTEVALUESETITEM.DIMENSIONATTRIBUTEVALUE");
            sbSql.AppendLine("INNER JOIN DimensionAttribute D3_DimensionAttribute ON D3_DimensionAttribute.RECID=D3_DimensionAttributeValue.DIMENSIONATTRIBUTE");
            sbSql.AppendLine("INNER JOIN DirpartyTable ON DirpartyTable.Recid = VendTable.Party");
            sbSql.AppendLine("WHERE purchTable.DATAAREAID='hoya' ");
            sbSql.AppendLine("AND purchTable.Inventlocationid='"+PORemainOBJ.Factory+"' ");
            sbSql.AppendLine("AND purchtable.numberSequenceGroup =  '"+strFac+"-"+PORemainOBJ.NumberSequenceGroup+"'");
            sbSql.AppendLine("AND PurchLine.RemainPurchPhysical  > 0 ");
            sbSql.AppendLine("AND DimensionAttribute.name='D1_Factory' AND HAttribute.name = 'D1_Factory' AND D3_DimensionAttribute.name='D3_Subsection' AND purchtable.PurchStatus  != '4' ");
            sbSql.AppendLine(" AND PurchTable.createdDateTime BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", PORemainOBJ.DateFrom) + "',103) ");
            sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", PORemainOBJ.DateTo) + "',103)");

            //sbSql.AppendLine(" AND CUSTGROUP IN ('" + InvoiceOBJ.CustomerGroup + "')");
            //sbSql.AppendLine(" AND INVOICEDATE BETWEEN CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateFrom) + "',103) ");
            //sbSql.AppendLine(" AND CONVERT(datetime,'" + String.Format("{0:dd/MM/yyyy}", InvoiceOBJ.DateTo) + "',103)");
            //sbSql.AppendLine("GROUP BY CURRENCYCODEISO,ECL_SALESCOMERCIAL,INVOICEDATE)salesTotal ");
            //sbSql.AppendLine(" GROUP BY CURRENCYCODEISO ");


            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection ADODBConnection = QueryDAL.GetADODBConnection();
            ADODBConnection.Open();

            rs.Open(sbSql.ToString(), ADODBConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);

            return rs;

        }// End getSaleByCustomerAndCurrency

    }
}
