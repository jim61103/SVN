using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iEMS_Setting
{
    class DBConnect
    {
        OleDbConnection conn;
        OleDbCommand cmd;
        OleDbDataReader dr;
        OleDbDataAdapter adt;

        public DBConnect()
        {
            
        }

        public int insert(string str, OleDbConnection conn)
        {
            int i = 0;
            try
            {
               
                cmd = new OleDbCommand(str, conn);
                i = cmd.ExecuteNonQuery();
                System.Threading.Thread.Sleep(30);
                return i;
                //return dr;
            }
            catch (InvalidOperationException e)
            {
                conn.Close();
                conn.Dispose();
                return 0;
            }
        }

        public OleDbConnection getConnect()
        {
            try
            {
                conn = new OleDbConnection(Global.ConnectDB);
                conn.Open();                
                return conn;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return null;
            }
        }

        public OleDbConnection getConnect(String s)
        {
            try
            {
                conn = new OleDbConnection(s);
                conn.Open();
                return conn;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return null;
            }
        }

        public OleDbDataReader queryforResult(string str, OleDbConnection conn)
        {
            try
            {
                cmd = new OleDbCommand(str, conn);
                dr = cmd.ExecuteReader();
                System.Threading.Thread.Sleep(30);
                return dr;
            }
            catch (InvalidOperationException e)
            {
                conn.Close();
                conn.Dispose();
                return null;
            }
        }

        public OleDbDataAdapter queryforAdapter(string str,OleDbConnection conn)
        {
            try
            {
                cmd = new OleDbCommand(str, conn);
                adt = new OleDbDataAdapter(str, conn);
                System.Threading.Thread.Sleep(30);
                return adt;
            }
            catch (InvalidOperationException e)
            {
                conn.Close();
                conn.Dispose();
                return null;
            }
        }

        public void closeData()
        {
            try
            {
                conn.Close();
                conn.Dispose();
                dr.Close();
            }
            catch (Exception e)
            {

            }
        }

    }
}
