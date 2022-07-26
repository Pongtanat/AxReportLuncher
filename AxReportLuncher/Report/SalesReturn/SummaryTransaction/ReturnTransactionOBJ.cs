using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NewVersion.Report.SalesReturn.SummaryTransaction
{
    class ReturnTransactionOBJ
    {
         public string strFactory;
        private int strShipmentLoc;
        private string strCustGroup;
        private string strNumberSequenceGroup;
        private DateTime dt1;
        private DateTime dt2;
        private bool _ShowWH;
        private string _WH;
        private string strSec;
        private string strItem1;
        private string strItem2;
        private string strVoucher1;
        private string strVoucher2;
        private int intTransType;
        private string Strcat;


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

        public string Section
        {
            get
            {
                return strSec;
            }
            set
            {
                strSec = value;
            }
        }

        public string Category
        {
            get
            {
                return Strcat;
            }
            set
            {
                Strcat = value;
            }
        }

        public bool ShowWH
        {
            get
            {
                return _ShowWH;
            }
            set
            {
                _ShowWH = value;
            }
        }

        public int TransType
        {
            get
            {
                return intTransType;
            }
            set
            {
                intTransType = value;
            }
        }

        public string ItemFrom
        {
            get
            {
                return strItem1;
            }
            set
            {
                strItem1 = value;
            }
        }

        public string VoucherFrom
        {
            get
            {
                return strVoucher1;
            }
            set
            {
                strVoucher1 = value;
            }
        }


        public string VoucherTo
        {
            get
            {
                return strVoucher2;
            }
            set
            {
                strVoucher2 = value;
            }
        }

        public string ItemTo
        {
            get
            {
                return strItem2;
            }
            set
            {
                strItem2 = value;
            }
        }




        public string WareHouse
        {
            get
            {
                return _WH;
            }
            set
            {
                _WH = value;
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

