using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NewVersion.Report.APReport
{
    class APReportOBJ
    {

        public string strFactory;
        private string strVender;
        private string strVenderGroup;
        private DateTime dt1;
        private DateTime dt2;

     
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



        public string vendercode
        {
            get
            {

                return strVender;
            }
            set
            {
                strVender = value;
            }
        }

        public string venderGroup
        {
            get
            {
                return strVenderGroup;
            }
            set
            {
                strVenderGroup = value;
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

