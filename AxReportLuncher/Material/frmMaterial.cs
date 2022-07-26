using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;





namespace NewVersion.Material
{
    public partial class frmMaterial : Form
    {

        string filePath;
        SQLConnectionDAL QueryDAL = new SQLConnectionDAL();
        DataTable dt = new DataTable();
        Microsoft.Dynamics.BusinessConnectorNet.Axapta Ax = new Microsoft.Dynamics.BusinessConnectorNet.Axapta();
        SQLConnectionDAL conn = new SQLConnectionDAL();
        Material.MaterialOBJ MaterialOBJ =  new NewVersion.Material.MaterialOBJ();




        public frmMaterial()
        {
            InitializeComponent();
        }

       
        private void frmMaterial_Load(object sender, EventArgs e)
        {

            btnGenerate.Enabled = false;
            MaterialOBJ.Factory = "RP";
        
           
        
        }

        private void btImport_Click(object sender, EventArgs e)
        {
            btImport.Enabled = false;

              string strSystemPath = System.IO.Directory.GetCurrentDirectory();
           
              string Folderpath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    OpenFileDialog dlg = new OpenFileDialog();

                     dlg.Filter = "Excel Files|*.xls;*.xlsx";
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                         filePath = dlg.FileName;
                         lblFile.Text = System.IO.Path.GetFileName(filePath);
                         btnGenerate.Enabled = true;
                    }

         } //end Import


        public DataTable READExcel(string path , int sheet)
        {

            //Instance reference for Excel Application
            Microsoft.Office.Interop.Excel.Application objXL = null;
            //Workbook refrence
            Microsoft.Office.Interop.Excel.Workbook objWB = null;
            Excel.Worksheet xlsSheet;

            DataSet ds = new DataSet();
            DataTable dtTemp = new DataTable();
            try
            {
                //Instancing Excel using COM services
                objXL = new Microsoft.Office.Interop.Excel.Application();
                //Adding WorkBook
                objWB = objXL.Workbooks.Open(path);
                xlsSheet = (Excel.Worksheet)objWB.Sheets[sheet];
                //Microsoft.Office.Interop.Excel.Worksheet objSHT = (Excel.Worksheet)objWB.Sheets[sheet];

               // foreach (Microsoft.Office.Interop.Excel.Worksheet objSHT in objWB)
               // {

                    int rows = xlsSheet.UsedRange.Rows.Count;
                    int cols = xlsSheet.UsedRange.Columns.Count;

                    int noofrow = 1;
                    //If 1st Row Contains unique Headers for datatable include this part else remove it
                    //Start
                    for (int c = 1; c <= cols; c++)
                    {
                        string colname = xlsSheet.Cells[1, c].Text;
                        dtTemp.Columns.Add(colname);
                        noofrow = 2;
                    }

                    // Add Column
                    //dtTemp.Columns.Add("Chk");

                    //END
                    for (int r = noofrow; r <= rows; r++)
                    {
                        DataRow dr = dtTemp.NewRow();
                        for (int c = 1; c <= cols; c++)
                        {
                            dr[c - 1] = xlsSheet.Cells[r, c].Text;
                        }
                        dtTemp.Rows.Add(dr);
                    }
                    ds.Tables.Add(dtTemp);
                //}

                //Closing workbook
                objWB.Close();
                //Closing excel application
                objXL.Quit();
            }

            catch (Exception ex)
            {

                objWB.Saved = true;
                //Closing work book
                objWB.Close();
                //Closing excel application
                objXL.Quit();
                //Response.Write("Illegal permission");

            }

            return dtTemp;
        }


        private void btnGenerate_Click(object sender, EventArgs e)
        {
            btnGenerate.Enabled = false;

            if (MaterialOBJ.Factory == "GMO")
            {

                GMO();
            }
            else if (MaterialOBJ.Factory == "RP")
            {
                RP();
            }
            
            btImport.Enabled = true;

        }


    void RP(){

     Excel.Application xlsApp = new Excel.Application();
     xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
     xlsApp.SheetsInNewWorkbook = 1;
     xlsApp.DisplayAlerts = false;
     xlsApp.Visible = false;


     Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(filePath);
     Excel._Worksheet xlsWorkSheet = xlsBookTemplate.Sheets[1];
     Excel.Range xlRange = xlsWorkSheet.UsedRange;

     Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord INVENTJOURNALTRANS, INVENTABLE, INVENTJOURNALTABLE;
     Microsoft.Dynamics.BusinessConnectorNet.AxaptaObject AxInventJournalTable, AxInventJournalTrans;
     System.Data.DataTable RS, RSLOCATION, RSLINE, RSCOST = new System.Data.DataTable();

     MaterialDAL MaterialDAL = new MaterialDAL();
    

     string strJournalId = "";
     string GlassType, SozaiDiv, ItemCD;
     System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();

     int line = 1;
     string strRS = "";


     try
     {
         string strServer = "";
         if (ConfigurationManager.AppSettings["SERVER"] == "LIVE")
         {
             strServer = "LIVE@192.1.87.243:2712";
         }
         else
         {
             strServer = "TEST@192.1.87.237:2712";

         }


         System.Net.NetworkCredential netCredential = new System.Net.NetworkCredential();
         netCredential.Domain = "HOYA";
         netCredential.Password = "P@ssw0rd";
         netCredential.UserName = "hopt-axadmin";
         Ax.LogonAs(netCredential.UserName, netCredential.Domain, netCredential, "hoya", "en-us", strServer, ConfigurationManager.AppSettings["SERVER"]);




     }
     catch (Exception ex)
     {
         Ax.TTSAbort();
         Ax.Logoff();
         MessageBox.Show("Network Connect Failed,Please contact your administrator");

     }


     try
     {
         Ax.TTSBegin();

         for (int sheet = 1; sheet <= xlsBookTemplate.Worksheets.Count; sheet++)
         {
             xlsWorkSheet = xlsBookTemplate.Sheets[sheet];
             Excel.Range xlRangeLine = xlsWorkSheet.UsedRange;
             line = 1;
             strRS = "";
             dt = READExcel(filePath,sheet);
           
             if (dt.Rows.Count > 0)
             {
                 // Group To AX
                 var query = from t in dt.AsEnumerable()
                               where t.Field<string>("DATE") != ""
                             select new
                             {
                                 strDate = t.Field<string>("DATE"),
                                 strItem = t.Field<string>("ITEMCD"),
                                 strGlassType = t.Field<string>("GLASSTYPE"), //ItemRemain
                                 strSozaiDiv = t.Field<string>("SOZAIDIV"),
                                 Qty = t.Field<string>("QUANTITY"),  //RemainQty
                                 strType = t.Field<string>("TYPE")

                             };

                 if (query.ToList().Count > 0)
                 {
                     

                     AxInventJournalTable = null;
                     lblType.Text = xlsWorkSheet.Name;
                     AxInventJournalTable = Ax.CreateAxaptaObject("AxInventJournalTable");

                     if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RETURN") || string.Equals(xlsWorkSheet.Name.ToString(), "USED-RT"))
                     {

                         AxInventJournalTable.Call("parmJournalNameId", "RP-RT-RM");

                     }
                     else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RD")) 
                     {
                         AxInventJournalTable.Call("parmJournalNameId", "RP-IS-NR");

                     }else{

                           AxInventJournalTable.Call("parmJournalNameId", "RP-IS-RM");

                     }

                     AxInventJournalTable.Call("parmDescription", xlsWorkSheet.Name + DateTime.Now.ToString("-yyyyMMdd-H:mm:ss"));
                     AxInventJournalTable.Call("parmInventLocationId", "RP");
                     AxInventJournalTable.Call("parmInventSiteId", "RP");
                     AxInventJournalTable.Call("parmLedgerDimension", 5637144739);


                     if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "SALE"))
                     {
                         RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813619070", "Z1PR", "NN");
                         AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                         AxInventJournalTable.Call("parmLedgerDimension", 5637144730);
                         AxInventJournalTable.Call("parmDifType", 0);
                         strRS = RS.Rows[0][0].ToString();
                     }
                     else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "USED"))
                     {
                         RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813613120", "Z1BR", "NN");
                         AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                         AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                         AxInventJournalTable.Call("parmDifType", 1);
                         strRS = RS.Rows[0][0].ToString();
                     }
                     else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "NG"))
                     {
                         RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813614190", "Z1AC", "NN");
                         AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                         AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                         AxInventJournalTable.Call("parmDifType", 2);
                         strRS = RS.Rows[0][0].ToString();
                     }
                     else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RETURN"))
                     {
                         RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813614190", "Z1AC", "NN");
                         AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                         AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                         AxInventJournalTable.Call("parmDifType", 3);
                         strRS = RS.Rows[0][0].ToString();

                     }
                     else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "DEAD"))
                     {
                         AxInventJournalTable.Call("parmDifType", 4);

                     }
                     else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "GD"))
                     {
                         RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813614190", "Z1AC", "NN");
                         AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                         AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                         AxInventJournalTable.Call("parmDifType", 5);
                         strRS = RS.Rows[0][0].ToString();
                     }
                     else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "NG-RT"))
                     {
                         AxInventJournalTable.Call("parmDifType", 7);

                     }
                     else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "USED-RT"))
                     {
                         RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813614190", "Z1AC", "NN");
                         AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                         AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                         AxInventJournalTable.Call("parmDifType", 6);
                         strRS = RS.Rows[0][0].ToString();
               

                     }

                     else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RD"))
                     {
                         RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813619070", "Z1RD", "NN");
                         AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                         AxInventJournalTable.Call("parmLedgerDimension", 5637146383);
                         AxInventJournalTable.Call("parmDifType", 9);
                         strRS = RS.Rows[0][0].ToString();
                     }


                     // END HEAD ///*******************


                     if (query.ToList().Count > 0)
                     {
                         if (strRS != "")
                         {
                             AxInventJournalTable.Call("parmNumOfLines", query.ToList().Count);
                             AxInventJournalTable.Call("save");

                             query.ToList().ForEach(q =>
                             {
                                 //Console.WriteLine(q.ItemReader + " : " + q.QtyReader);

                                 if (q.strItem != "")
                                 {
                                     RSLINE = MaterialDAL.getDimAttrValueLine(q.strItem, q.strGlassType, q.strSozaiDiv);
                                 }
                                 else
                                 {
                                     RSLINE = MaterialDAL.getDimAttrValueLine("", q.strGlassType, q.strSozaiDiv);
                                 }

                                 strJournalId = AxInventJournalTable.Call("parmJournalId").ToString();
                                 INVENTJOURNALTABLE = (Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord)Ax.CallStaticRecordMethod("InventJournalTable", "find", strJournalId.ToString(), true);
                                 INVENTJOURNALTRANS = Ax.CreateAxaptaRecord("InventJournalTrans");
                                 INVENTABLE = (Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord)Ax.CallStaticRecordMethod("InventTable", "find", RSLINE.Rows[0][0].ToString());

                                 //AxInventJournalTrans = new Microsoft.Dynamics.BusinessConnectorNet.AxaptaObject inventJournalTableOBJ;

                                 INVENTJOURNALTRANS.Clear();
                                 INVENTJOURNALTRANS.InitValue();
                                 INVENTJOURNALTRANS.Call("initFromInventJournalTable", INVENTJOURNALTABLE);
                                 INVENTJOURNALTRANS.Call("initFromInventTable", INVENTABLE);

                                 AxInventJournalTrans = (Microsoft.Dynamics.BusinessConnectorNet.AxaptaObject)Ax.CallStaticClassMethod("AxInventJournalTrans", "newInventJournalTrans", INVENTJOURNALTRANS);
                                 RSLOCATION = MaterialDAL.getAddressByLocation("RP", "RP");
                                 AxInventJournalTrans.Call("parmInventDimId", RSLOCATION.Rows[0][0].ToString());



                                 if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "SALE"))
                                 {
                                     RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813619070", "Z1PR", "NN");
                                     AxInventJournalTrans.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                     AxInventJournalTrans.Call("parmLedgerDimension", 5637144730);
                                 }
                                 else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RD"))
                                 {
                                     RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813619070", "Z1RD", "NN");
                                     AxInventJournalTrans.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                     AxInventJournalTrans.Call("parmLedgerDimension", 5637146383);
                                 }
                                 else
                                 {

                                     if (RSLINE.Rows[0][1].ToString() == "EB")
                                     {
                                         RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813614190", "Z1AC", "NN");
                                         AxInventJournalTrans.Call("parmDefaultDimension", RS.Rows[0][0].ToString());

                                     }
                                     else
                                     {
                                         RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813613120", "Z1BR", "NN");
                                         AxInventJournalTrans.Call("parmDefaultDimension", RS.Rows[0][0].ToString());

                                     }

                                     AxInventJournalTrans.Call("parmLedgerDimension", 5637144739);
                                 }//end if

                                 AxInventJournalTrans.Call("parmTransDate", Convert.ToDateTime(q.strDate));

                                 if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RETURN") || string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "USED-RT"))
                                 {
                                     RSCOST = MaterialDAL.getCost(RSLINE.Rows[0][0].ToString());
                                     AxInventJournalTrans.Call("parmCostPrice", RSCOST.Rows[0][0].ToString());
                                     AxInventJournalTrans.Call("parmQty", Math.Round(Convert.ToDouble(q.Qty.ToString()), 2));
                                 }
                                 else
                                 {
                                     AxInventJournalTrans.Call("parmQty", Math.Round(Convert.ToDouble(q.Qty.ToString()) * -1, 2));

                                 }



                                 lblMessage.Text = "Process  GlassType :" + q.strGlassType + " Sozidive :" + q.strSozaiDiv;
                                 lblMessage.ForeColor = System.Drawing.Color.Green;
                                 lblMessage.Focus();

                                 AxInventJournalTrans.Call("save");
                                 line = line + 1;

                             });
                         } // end query
                     }
                 }// check dt
             }//end if
         
         } //end for sheets

         Ax.TTSCommit();
         Ax.TTSAbort();
         xlsApp.Workbooks.Close();
         xlsApp.Quit();
         lblMessage.Text = "Complete";
         lblMessage.ForeColor = System.Drawing.Color.SlateBlue;
         lblMessage.Focus();


     }
     catch (Exception ex)
     {


         lblMessage.Text = "ERROR Line : " + line + " TYPE:" + lblType.Text;

         lblMessage.ForeColor = System.Drawing.Color.Red;
         lblMessage.Focus();
         xlsApp.Workbooks.Close();
     }

 }

    void GMO()
        {

            Excel.Application xlsApp = new Excel.Application();
            xlsApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
            xlsApp.SheetsInNewWorkbook = 1;
            xlsApp.DisplayAlerts = false;
            xlsApp.Visible = false;


            Excel.Workbook xlsBookTemplate = xlsApp.Workbooks.Open(filePath);
            Excel._Worksheet xlsWorkSheet = xlsBookTemplate.Sheets[1];
            Excel.Range xlRange = xlsWorkSheet.UsedRange;

            Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord INVENTJOURNALTRANS, INVENTABLE, INVENTJOURNALTABLE;
            Microsoft.Dynamics.BusinessConnectorNet.AxaptaObject AxInventJournalTable, AxInventJournalTrans;
            System.Data.DataTable RS, RSLOCATION, RSLINE, RSCOST = new System.Data.DataTable();

            MaterialDAL MaterialDAL = new MaterialDAL();


            string strJournalId = "";
            bool Head = true;
            string GlassType, SozaiDiv, ItemCD;
            System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();

            int line = 1;
            string strRS = "";

            try
            {
                string strServer = "";
                if (ConfigurationManager.AppSettings["SERVER"] == "LIVE")
                {
                    strServer = "LIVE@192.1.87.243:2712";
                }
                else
                {
                    strServer = "TEST@192.1.87.237:2712";

                }


                System.Net.NetworkCredential netCredential = new System.Net.NetworkCredential();
                netCredential.Domain = "HOYA";
                netCredential.Password = "P@ssw0rd";
                netCredential.UserName = "hopt-axadmin";
                Ax.LogonAs(netCredential.UserName, netCredential.Domain, netCredential, "hoya", "en-us", strServer, ConfigurationManager.AppSettings["SERVER"]);




            }
            catch (Exception ex)
            {
                Ax.TTSAbort();
                Ax.Logoff();
                MessageBox.Show("Network Connect Failed,Please contact your administrator");

            }

            try
            {
                Ax.TTSBegin();

                for (int sheet = 1; sheet <= xlsBookTemplate.Worksheets.Count; sheet++)
                {
                    xlsWorkSheet = xlsBookTemplate.Sheets[sheet];
                    Excel.Range xlRangeLine = xlsWorkSheet.UsedRange;
                    line = 1;
                    strRS = "";
                    dt = READExcel(filePath, sheet);

                    if (dt.Rows.Count > 0)
                    {
                        // Group To AX
                        var query = from t in dt.AsEnumerable()
                                    where t.Field<string>("Date") != ""
                                    select new
                                    {
                                        strDate = t.Field<string>("Date"),
                                        strItem = t.Field<string>("Item"),
                                        strGlassType = t.Field<string>("Glass type"), //ItemRemain
                                        strSozaiDiv = t.Field<string>("Sozaidiv"),
                                        Qty = t.Field<string>("Quantity"),  //RemainQty
                                        strType = t.Field<string>("TYPE")

                                    };

                        if (query.ToList().Count > 0)
                        {


                            AxInventJournalTable = null;
                            lblType.Text = xlsWorkSheet.Name;
                            AxInventJournalTable = Ax.CreateAxaptaObject("AxInventJournalTable");

                            if (!string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RETURN") && !string.Equals(xlsWorkSheet.Name.ToString(), "OUT-RT") && !string.Equals(xlsWorkSheet.Name.ToString(), "NG-RT"))
                            {

                                AxInventJournalTable.Call("parmJournalNameId", "MO-IS-RM");

                            }else
                            {

                                AxInventJournalTable.Call("parmJournalNameId", "MO-RT-RM");

                            }

                         

                            AxInventJournalTable.Call("parmInventLocationId", "GMO");
                            AxInventJournalTable.Call("parmInventSiteId", "GMO");

                            if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "SALE"))
                            {
                                AxInventJournalTable.Call("parmDescription", "HRMO" + DateTime.Now.ToString("-yyyyMMdd-H:mm:ss"));
                                RS = MaterialDAL.getDimAttrValSetItemAX("MO", "A812619070", "ZRTHOOP", "NN");
                                AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                AxInventJournalTable.Call("parmLedgerDimension", 5637144730);
                                AxInventJournalTable.Call("parmDifType", 0);
                                strRS = RS.Rows[0][0].ToString();
                            }
                            else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "OUT"))
                            {
                                AxInventJournalTable.Call("parmDescription", "PF" + DateTime.Now.ToString("-yyyyMMdd-H:mm:ss"));
                                RS = MaterialDAL.getDimAttrValSetItemAX("MO", "A812613110", "ZCOKT", "NN");
                                AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                                AxInventJournalTable.Call("parmDifType", 1);
                                strRS = RS.Rows[0][0].ToString();
                            }
                            else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "NG"))
                            {
                                AxInventJournalTable.Call("parmDescription", "PFNG" + DateTime.Now.ToString("-yyyyMMdd-H:mm:ss"));
                                RS = MaterialDAL.getDimAttrValSetItemAX("MO", "A812613110", "ZNGPF", "NN");
                                AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                                AxInventJournalTable.Call("parmDifType", 2);
                                strRS = RS.Rows[0][0].ToString();
                            }
                            else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RETURN"))
                            {
                                AxInventJournalTable.Call("parmDescription", "RTPF" + DateTime.Now.ToString("-yyyyMMdd-H:mm:ss"));
                                RS = MaterialDAL.getDimAttrValSetItemAX("MO", "A812613110", "ZRTLM", "NN");
                                AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                                AxInventJournalTable.Call("parmDifType", 3);
                                strRS = RS.Rows[0][0].ToString();
                            }
                            else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "DEAD"))
                            {

                                AxInventJournalTable.Call("parmDescription", "DEAD" + DateTime.Now.ToString("-yyyyMMdd-H:mm:ss"));
                                RS = MaterialDAL.getDimAttrValSetItemAX("MO", "A812612000", "ZDEAD", "NN");
                                AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                AxInventJournalTable.Call("parmLedgerDimension", 5637194826);
                                AxInventJournalTable.Call("parmDifType", 4);
                                strRS = RS.Rows[0][0].ToString();
                            }
                            else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "GD"))
                            {
                                RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813614190", "Z1AC", "NN");
                                AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                                AxInventJournalTable.Call("parmDifType", 5);
                                strRS = RS.Rows[0][0].ToString();
                            }
                            else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "NG-RT"))
                            {

                                AxInventJournalTable.Call("parmDescription", "RTNG" + DateTime.Now.ToString("-yyyyMMdd-H:mm:ss"));
                                RS = MaterialDAL.getDimAttrValSetItemAX("MO", "A812613110", "ZCOKT", "NN");
                                AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                                AxInventJournalTable.Call("parmDifType", 7);
                                strRS = RS.Rows[0][0].ToString();
                            }
                            else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "OUT-RT"))
                            {
                                AxInventJournalTable.Call("parmDescription", "PF" + DateTime.Now.ToString("-yyyyMMdd-H:mm:ss"));
                                RS = MaterialDAL.getDimAttrValSetItemAX("MO", "A812613110", "ZNGPF", "NN");
                                AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                AxInventJournalTable.Call("parmLedgerDimension", 5637144739);
                                AxInventJournalTable.Call("parmDifType", 6);
                                strRS = RS.Rows[0][0].ToString();
                            }

                            else if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RD"))
                            {
                                RS = MaterialDAL.getDimAttrValSetItemAX("RP", "A813619070", "Z1RD", "NN");
                                AxInventJournalTable.Call("parmDefaultDimension", RS.Rows[0][0].ToString());
                                AxInventJournalTable.Call("parmLedgerDimension", 5637146383);
                                AxInventJournalTable.Call("parmDifType", 9);
                                strRS = RS.Rows[0][0].ToString();
                            }

                            //5637146383
                            AxInventJournalTable.Call("parmNumOfLines", query.ToList().Count);
                            AxInventJournalTable.Call("save");
                            // END HEAD ///*******************

                            lblMessage.ForeColor = System.Drawing.Color.Green;
                            lblMessage.Focus();


                            if (query.ToList().Count > 0)
                            {
                                AxInventJournalTable.Call("parmNumOfLines", query.ToList().Count);
                                //AxInventJournalTable.Call("save");


                               // query.ToList().ForEach(q =>
                                //{
                                    //Console.WriteLine(q.ItemReader + " : " + q.QtyReader);
                                foreach (DataRow dtRow in dt.Rows)
                                {
                                    if (strRS != "" && dtRow["Item"].ToString() != "")
                                    {
                                        RSLINE = MaterialDAL.getDimAttrValueLine(dtRow["Item"].ToString(), "", "");

                                        strJournalId = AxInventJournalTable.Call("parmJournalId").ToString();
                                        INVENTJOURNALTABLE = (Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord)Ax.CallStaticRecordMethod("InventJournalTable", "find", strJournalId.ToString(), true);
                                        INVENTJOURNALTRANS = Ax.CreateAxaptaRecord("InventJournalTrans");
                                        INVENTABLE = (Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord)Ax.CallStaticRecordMethod("InventTable", "find", RSLINE.Rows[0][0].ToString());

                                        //AxInventJournalTrans = new Microsoft.Dynamics.BusinessConnectorNet.AxaptaObject inventJournalTableOBJ;

                                        INVENTJOURNALTRANS.Clear();
                                        INVENTJOURNALTRANS.InitValue();
                                        INVENTJOURNALTRANS.Call("initFromInventJournalTable", INVENTJOURNALTABLE);
                                        INVENTJOURNALTRANS.Call("initFromInventTable", INVENTABLE);

                                        AxInventJournalTrans = (Microsoft.Dynamics.BusinessConnectorNet.AxaptaObject)Ax.CallStaticClassMethod("AxInventJournalTrans", "newInventJournalTrans", INVENTJOURNALTRANS);
                                        RSLOCATION = MaterialDAL.getAddressByLocation("GMO", "GMO");
                                        AxInventJournalTrans.Call("parmInventDimId", RSLOCATION.Rows[0][0].ToString());
                                        AxInventJournalTrans.Call("parmDefaultDimension", strRS);

                                        ///AxInventJournalTrans.Call("parmLedgerDimension", 5637144730);
                                        AxInventJournalTrans.Call("parmTransDate", Convert.ToDateTime(dtRow["Date"]));

                                        if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "RETURN") || string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "OUT-RT") || string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "NG-RT"))
                                        {
                                            RSCOST = MaterialDAL.getCost(RSLINE.Rows[0][0].ToString());
                                            AxInventJournalTrans.Call("parmCostPrice", RSCOST.Rows[0][0].ToString());
                                            AxInventJournalTrans.Call("parmQty", Math.Round(Convert.ToDouble(dtRow["Quantity"].ToString()), 2));
                                        }
                                        else
                                        {
                                            AxInventJournalTrans.Call("parmQty", Math.Round(Convert.ToDouble(dtRow["Quantity"].ToString()) * -1, 2));

                                        }


                                        if (string.Equals(xlsWorkSheet.Name.ToString().ToUpper(), "DEAD"))
                                        {

                                            AxInventJournalTrans.Call("parmLedgerDimension", 5637194826);
                                        }
                                        else
                                        {
                                            AxInventJournalTrans.Call("parmLedgerDimension", 5637144739);

                                        }



                                        lblMessage.Text = "Process   :" + dtRow["Item"].ToString();
                                        AxInventJournalTrans.Call("save");
                                        line = line + 1;
                                    }
                                }// foreach Dt

                               // });





                            } // end query

                        }// check dt
                    }//end if

                } //end for sheets

                Ax.TTSCommit();
                Ax.TTSAbort();
                xlsApp.Workbooks.Close();
                xlsApp.Quit();
                lblMessage.Text = "Complete";
                lblMessage.ForeColor = System.Drawing.Color.SlateBlue;
                lblMessage.Focus();


            }
            catch (Exception ex)
            {


                lblMessage.Text = "ERROR Line : " + line + " TYPE:" + lblType.Text;

                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Focus();
                xlsApp.Workbooks.Close();
            }

        }

 private void gbSection_Validated(object sender, EventArgs e)
 {
     GroupBox g = gbSection as GroupBox;
     var a = from RadioButton r in g.Controls where r.Checked == true select r.Name;
     string strchecked = a.First();

     switch (strchecked)
     {
         case "rdoGMO":
            MaterialOBJ.Factory = "GMO";
             break;
         case "rdoRP":
             MaterialOBJ.Factory = "RP";
             break;

     }

 }



    }//end classs

   
}
