using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace  NewVersion.Material
{

 public class MaterialOBJ
    {
        public string strFactory;
        private int strShipmentLoc;
        private string strCustGroup;
        private string strNumberSequenceGroup;
        private DateTime dt1;
        private DateTime dt2;
        private string ROLLFAC;

        /*
            public  InvoiceReportOBJ(DataRow dr)
                {
                    strFactory = dr["Factory"].ToString();
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


        public string _ROLLFAC
        {
            get
            {
                return ROLLFAC;
            }
            set
            {
                ROLLFAC = value;
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


    }//end class
}
