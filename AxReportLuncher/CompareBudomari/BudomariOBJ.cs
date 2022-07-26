using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NewVersion.CompareBudomari
{
    class BudomariOBJ
    {

        public string strFactory;
        private int strShipmentLoc;
        private string strCustGroup;
        private string numbersequence;
        private string strNumberSequenceGroup;
        private string strCurr;

        private DateTime dt1;
        private DateTime dt2;
        private string strInvoiceAccount;

        private string Sheet1;
        private string Sheet2;
        /*  
         public  ARReconcileOBJ(DataRow dr)
            {
                       strFactory = dr["Factory"].ToString();
                       strInvoiceAccount = dr["InvoiceAccount"].ToString();
                       dt1 = Convert.ToDateTime(dr["dt1"].ToString());
                       dt2 = Convert.ToDateTime(dr["dt2"].ToString());
      
            }
      */
        public string Factory
        {
            get
            {
                return strFactory;
            }
            set
            {
                strFactory = value;
            }
        }

        public int ShipmentLocation
        {
            get
            {
                return strShipmentLoc;
            }
            set
            {
                strShipmentLoc = value;
            }
        }

        public string InvoiceAccount
        {
            get
            {
                return strInvoiceAccount;
            }
            set
            {
                strInvoiceAccount = value;
            }
        }



        public string CurrencyISO
        {
            get
            {
                return strCurr;
            }
            set
            {
                strCurr = value;
            }
        }

        public string CustomerGroup
        {
            get
            {
                return strCustGroup;
            }
            set
            {
                strCustGroup = value;
            }

        }

        public string NumberSequenceGroup
        {
            get
            {
                return strNumberSequenceGroup;
            }
            set
            {
                strNumberSequenceGroup = value;
            }

        }


        public string numbersequence2
        {
            get
            {
                return numbersequence;

            }
            set
            {

                numbersequence = value;

            }
        }



        public DateTime DateFrom
        {
            get
            {
                return dt1;
            }
            set
            {
                dt1 = value;
            }
        }


        public DateTime DateTo
        {
            get
            {
                return dt2;
            }
            set
            {
                dt2 = value;
            }

        }

          public string GetSheet1
        {
            get
            {
                return Sheet1;
            }
            set
            {
                Sheet1 = value;
            }

        }


          public string GetSheet2
        {
            get
            {
                return Sheet2;
            }
            set
            {
                Sheet2 = value;
            }

        }




    }//end Class
}
