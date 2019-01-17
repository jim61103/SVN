using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iEMS_Setting
{
    static public class Global
    {
        public static string UserID = "";
        public static string Level = "";
        public static string ConnectDB = "";

        //BUMP UAT
        public static string BUMP_UAT_ConnectDB = "Provider=OraOleDB.Oracle;User ID=insitedev;Data Source=DEV_AQA;Password=insitedevaqa;OLEDB.NET=True;";
      
        //BUMP Prod
        public static string BUMPConnectDB = "Provider=OraOleDB.Oracle;User ID=insitedev;Data Source=PROORCL;Password=insite!p;OLEDB.NET=True;";
        
        //DLP
        public static string DLPConnectDB = "Provider=OraOleDB.Oracle;User ID=insitedev;Data Source=DLPPROD;Password=insite!p;OLEDB.NET=True;";

        public static string Env = "";

        //flag for by Pass change DB, always use UAT ,only for 110.1
        public static bool isUAT = false;
    }
}
