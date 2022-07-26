using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace NewVersion
{
   
    class AXREportLancherBLL
    {
        AXReportLuncherDAL Adapter = new AXReportLuncherDAL();
        public Array getAllMenu(){
            string strAllMenu = "";
            DataTable dt = Adapter.getAllMenu();

            foreach (DataRow dr in dt.Rows)
            {
                strAllMenu += dr["NOTE"] + ",";
            }

            strAllMenu = strAllMenu.Substring(0, strAllMenu.Length - 1).Replace("", "");
            string[] arrAllMenu = strAllMenu.Split(',');

            string strMenu = "";
            foreach (string str in arrAllMenu)
            {
                if (strMenu.IndexOf(str) == -1)
                {
                    strMenu += str + ",";

                }
            }


            strMenu = strMenu.Substring(0, strMenu.Length - 1);
            string [] arrMenu= strMenu.Split(',');
            Array.Sort(arrMenu);
            return arrMenu;
        }

        public DataTable getMenuByUser(string strUser)
        {
            return Adapter.getMenuByUser(strUser);

        }

        public DataTable getRoleuser(string strUser)
        {
            return Adapter.getRoleuser(strUser);

        }

    }
}
