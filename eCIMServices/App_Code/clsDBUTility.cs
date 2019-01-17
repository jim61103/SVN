using System.Data;
using System.Data.OleDb;
using System;
using System.Net;


/// <summary>
/// clsDBUTility 的摘要描述
/// </summary>
public class clsDBUTility
{
    public OleDbConnection conn;
    private string oleDbStr = "";
    private DataRow row;
    private string strSQL = "";

    public clsDBUTility()
    {
        //
        // TODO: 在此加入建構函式的程式碼
        //
    }

    public void New()
    {
        //oleDbStr = "Provider=" & _Settings.DBProvider & ";Data Source=" & _Settings.DBDataSource & ";Persist Security Info=True;Password=" & _Settings.DBPassword & ";User ID=" & _Settings.DBUserID & "";
    }

    public void New(string oleDbSet)
    {
        oleDbStr = oleDbSet;
    }

    public void OpenDBLink()
    {
        if (conn != null)
        {
            if (conn.State == ConnectionState.Open)
            {
                ClosedDBLink();
            }
        }
        conn = new OleDbConnection(oleDbStr);
        conn.Open();
        AlterSchema();
    }

    public void OpenDBLink(string oleDbString)
    {
        try
        {
            if (conn != null)
            {
                if (conn.State == ConnectionState.Open)
                {
                    ClosedDBLink();
                }
            }
            conn = new OleDbConnection(oleDbString);
            conn.Open();
        }
        catch (System.Exception)
        {
            throw;
        }
    }

    public void ClosedDBLink()
    {
        try
        {
            if (conn != null)
            {
                conn.Close();
                conn.Dispose();
                conn = null;
            }
        }
        catch (System.Exception)
        {
            throw;
        }
    }

    private void AlterSchema()
    {
        //' To define the default schema'
       /* if (_Settings.DBSchema.equal("")) {
            strSQL = "alter session set current_schema=" & _Settings.DBSchema;
            OleDbCommand command = new OleDbCommand(strSQL, conn);
            command.ExecuteNonQuery();
            command.Dispose();
        }*/
    }

    public void ExecuteNonQuery(string sSQL)
    {
        try
        {
            if (conn == null)
            {
                this.OpenDBLink();
            }
            else if (conn.State == ConnectionState.Closed)
            {
                OpenDBLink();
            }
            OleDbCommand command = new OleDbCommand(sSQL, conn);
            command.ExecuteNonQuery();
            command.Dispose();
        }
        catch (System.Exception)
        {
            throw;
        }
    }

    public DataSet SQLQuery(string sSQLtxt, bool bTimeCheck)
    {
        DataSet SQLQuary;
        OleDbDataAdapter adapter = null;
        bTimeCheck = true;
        try
        {            
            long lStartTime = 0;
            long lEndTime = 0;
            lStartTime = DateTime.Now.Ticks;
            if (conn == null) {
                OpenDBLink();
            }
            else if (conn.State == ConnectionState.Closed) {
                OpenDBLink();
            }
            
            adapter =new OleDbDataAdapter(sSQLtxt, conn);
            SQLQuary = new DataSet();
            adapter.Fill(SQLQuary, "Query");

            if (SQLQuary.Tables.Count ==0) {
                SQLQuary = null;
            }
            else 
            {
                if (SQLQuary.Tables[0].Rows.Count == 0) {
                    SQLQuary = null;
                }
            }
            

            lEndTime = DateTime.Now.Ticks;
            TimeSpan timeSpan  = new TimeSpan(lEndTime - lStartTime);
            double dTotal = timeSpan.TotalSeconds;
            //if (_Settings.EnableSQLLog) {
            //    RecordLog(timeSpan.ToString(), sSQLtxt);
            //}

            if ((!bTimeCheck) && (dTotal > 20)) {
                bTimeCheck = true;
            }

            
            //if (_Settings.SQLWarningMail.Length > 0 && _Settings.SQLWarningTime > 0 && dTotal > _Settings.SQLWarningTime && bTimeCheck ) {
                try 
	            {	        
		            string sMailBody = "";
                    string sServerIP = "";
                    IPHostEntry ipE = Dns.GetHostEntry(Dns.GetHostName());
                    IPAddress[] IpA = ipE.AddressList;
                    for (int i = 0; i < IpA.GetUpperBound(0); i++)
                    {
			            sServerIP = IpA[i].ToString();
                    }
                }
	            catch (Exception)
                {		
		            throw;
                }

                return SQLQuary;
            //}
        }
        catch (System.Exception)
        {

            throw;
        }
        finally
        {
            if (adapter != null)
            {
                adapter.Dispose();
                adapter = null;
            }
        }
    }

    private void RecordLog(string sProcessTime, string sSQL)
    {
        try 
	    {
            if (1 != 1)
            {

            }
            else
            {

            }
	    }
	    catch (Exception)
	    {
    		
		    throw;
	    }
    }
}
