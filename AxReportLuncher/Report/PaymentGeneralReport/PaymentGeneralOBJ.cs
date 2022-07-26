using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NewVersion.Report.PaymentGeneralReport
{
    class PaymentGeneralOBJ
    {

        public string strFactory;
        private int strShipmentLoc;
        private string strCustGroup;
        private string numbersequence;
        private string strNumberSequenceGroup;
        private DateTime dt1;
        private DateTime dt2;
        private string StrGroupVoucher_;
        private string StrStartVoucher_;
        private string StrEndVoucher_;


        /*
            public  InvoiceReportOBJ(DataRow dr)
                {
                    strFactory = dr["Factory"].ToString();
                    dt1 = Convert.ToDateTime(dr["dt1"].ToString());
                    dt2 = Convert.ToDateTime(dr["dt2"].ToString());
                }
        */

        public string GroupVoucher
        {
            get
            {
                return StrGroupVoucher_;
            }
            set
            {
                StrGroupVoucher_ = value;
            }
        }

        public string StartVoucher
        {
            get
            {
                return StrStartVoucher_;
            }
            set
            {
                StrStartVoucher_ = value;
            }
        }

        public string EndVoucher
        {
            get
            {
                return StrEndVoucher_;
            }
            set
            {
                StrEndVoucher_ = value;
            }
        }



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




    }
}
