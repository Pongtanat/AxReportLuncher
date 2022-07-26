using System;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace NewVersion
{
public partial 	class SQLConnectionDAL
	{
        

    private SqlConnection objConn;
    private SqlCommand objCmd;
    private SqlTransaction Trans;
    private String strConnString;

    private string _Username = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

    public SQLConnectionDAL()
    {

        /*
        if (_Username == "HOYA\\fongmek" || _Username == "HOYA\\panpanya")
        {
            strConnString = ConfigurationManager.ConnectionStrings["HOAX61TEST"].ConnectionString.ToString();
            FrmMain._Database = "TEST";
        }
        else
        {
            strConnString = ConfigurationManager.ConnectionStrings["HOAX61LIVE"].ConnectionString.ToString();
         }
        */
        if (ConfigurationManager.AppSettings["SERVER"] == "LIVE")
        {
            strConnString = ConfigurationManager.ConnectionStrings["HOAX61LIVE"].ConnectionString.ToString();
            FrmMain._Database = ConfigurationManager.AppSettings["SERVER"];
        }
        else
        {
            strConnString = ConfigurationManager.ConnectionStrings["HOAX61TEST"].ConnectionString.ToString();
      
            FrmMain._Database = ConfigurationManager.AppSettings["SERVER"];
        }
        
    }


    public SqlDataReader QueryDataReader(String strSQL)
    {
        SqlDataReader dtReader;
        objConn = new SqlConnection();
        objConn.ConnectionString = strConnString;
        objConn.Open();

        objCmd = new SqlCommand(strSQL, objConn);
        dtReader = objCmd.ExecuteReader();
        return dtReader; //*** Return DataReader ***//
    }

    public DataSet QueryDataSet(String strSQL)
    {
        DataSet ds = new DataSet();
        SqlDataAdapter dtAdapter = new SqlDataAdapter();
        objConn = new SqlConnection();
        objConn.ConnectionString = strConnString;
        objConn.Open();

        objCmd = new SqlCommand();
        objCmd.Connection = objConn;
        objCmd.CommandText = strSQL;
        objCmd.CommandType = CommandType.Text;

        dtAdapter.SelectCommand = objCmd;
        dtAdapter.Fill(ds);
        return ds;   //*** Return DataSet ***//
    }

    public DataTable QueryDataTable(String strSQL)
    {
        SqlDataAdapter dtAdapter;
        DataTable dt = new DataTable();
        objConn = new SqlConnection();
        objConn.ConnectionString = strConnString;
        objConn.Open();

        dtAdapter = new SqlDataAdapter(strSQL, objConn);
        dtAdapter.Fill(dt);
        return dt; //*** Return DataTable ***//
    }

    public Boolean QueryExecuteNonQuery(String strSQL)
    {
        objConn = new SqlConnection();
        objConn.ConnectionString = strConnString;
        objConn.Open();

        try
        {
            objCmd = new SqlCommand();
            objCmd.Connection = objConn;
            objCmd.CommandType = CommandType.Text;
            objCmd.CommandText = strSQL;

            objCmd.ExecuteNonQuery();
            return true; //*** Return True ***//
        }
        catch (Exception)
        {
            return false; //*** Return False ***//
        }
    }


    public Object QueryExecuteScalar(String strSQL)
    {
        Object obj;
        objConn = new SqlConnection();
        objConn.ConnectionString = strConnString;
        objConn.Open();

        try
        {
            objCmd = new SqlCommand();
            objCmd.Connection = objConn;
            objCmd.CommandType = CommandType.Text;
            objCmd.CommandText = strSQL;

            obj = objCmd.ExecuteScalar();  //*** Return Scalar ***//
            return obj;
        }
        catch (Exception)
        {
            return null; //*** Return Nothing ***//
        }
    }

    public void TransStart()
    {
        objConn = new SqlConnection();
        objConn.ConnectionString = strConnString;
        objConn.Open();
        Trans = objConn.BeginTransaction(IsolationLevel.ReadCommitted);
    }


    public void TransExecute(String strSQL)
    {
        objCmd = new SqlCommand();
        objCmd.Connection = objConn;
        objCmd.Transaction = Trans;
        objCmd.CommandType = CommandType.Text;
        objCmd.CommandText = strSQL;
        objCmd.ExecuteNonQuery();
    }


    public void TransRollBack()
    {
        Trans.Rollback();
    }

    public void TransCommit()
    {
        Trans.Commit();
    }

    public void Close()
    {
        objConn.Close();
        objConn = null;
    }

    public ADODB.Connection GetADODBConnection()
    {
        ADODB.Connection ADODBConnection = new ADODB.Connection();


        /*
        if (_Username == "HOYA\\fongmek" || _Username == "HOYA\\panpanya")
        {
           // strConnString = ConfigurationManager.ConnectionStrings["HOAX61TEST"].ConnectionString.ToString();
            ADODBConnection.ConnectionString = ConfigurationManager.AppSettings["ADO_HOAX61TEST"];
            //FrmMain._Database = "TEST";
        }
        else
        {
            ADODBConnection.ConnectionString = ConfigurationManager.AppSettings["ADO_HOAX61LIVE"];
            FrmMain._Database = "LIVE";
        }
        */
        

        if (ConfigurationManager.AppSettings["SERVER"] == "LIVE")
        {
            ADODBConnection.ConnectionString = ConfigurationManager.AppSettings["ADO_HOAX61LIVE"];
            FrmMain._Database = ConfigurationManager.AppSettings["SERVER"];
        }
        else if(ConfigurationManager.AppSettings["SERVER"] == "TEST")
        {
            ADODBConnection.ConnectionString = ConfigurationManager.AppSettings["ADO_HOAX61TEST"];
            FrmMain._Database = ConfigurationManager.AppSettings["SERVER"];

        }
       
            

        ADODBConnection.CursorLocation = ADODB.CursorLocationEnum.adUseClient;
        ADODBConnection.ConnectionTimeout = 6000;
        ADODBConnection.CommandTimeout = 6000;
      
        return ADODBConnection;


        }

    public ADODB.Connection HOAX244Connect()
    {
        ADODB.Connection ADODBConnection = new ADODB.Connection();

        ADODBConnection.ConnectionString = ConfigurationManager.AppSettings["ADO_OOMLBAX244"];
        
        ADODBConnection.CursorLocation = ADODB.CursorLocationEnum.adUseClient;
        ADODBConnection.ConnectionTimeout = 6000;
        ADODBConnection.CommandTimeout = 6000;

        return ADODBConnection;


    }



    public ADODB.Connection RPContingConnect()
    {
        ADODB.Connection ADODBConnection = new ADODB.Connection();

        ADODBConnection.ConnectionString = ConfigurationManager.AppSettings["ADO_COSTINGRP"];
        FrmMain._Database = ConfigurationManager.AppSettings["SERVER"];

            
        

        ADODBConnection.CursorLocation = ADODB.CursorLocationEnum.adUseClient;
        ADODBConnection.ConnectionTimeout = 6000;
        ADODBConnection.CommandTimeout = 6000;

        return ADODBConnection;


    }



	}
}
