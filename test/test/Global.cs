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

        public static string sConnString = "Provider=OraOLEDB.Oracle;Data Source=FDCPROORCL;User ID=att_ecim;Password=attecim#182;OLEDB.NET=True;";

        //BUMP UAT
        public static string BUMP_UAT_ConnectDB = "Provider=OraOleDB.Oracle;User ID=insitedev;Data Source=DEV_AQA;Password=insitedevaqa;OLEDB.NET=True;";
      
        //BUMP Prod
        public static string BUMPConnectDB = "Provider=OraOleDB.Oracle;User ID=insitedev;Data Source=PROORCL;Password=insite!p;OLEDB.NET=True;";
        
        //DLP
        public static string DLPConnectDB = "Provider=OraOleDB.Oracle;User ID=insitedev;Data Source=DLPPROD;Password=insite!p;OLEDB.NET=True;";
        
        public static string Env = "";

        //PROD EFGP
        public static string EFGP_PROD = "http://10.185.30.231:8086/NaNaWeb/services/WorkflowService?wsdl";

        //UAT EFGP
        public static string EFGP_UAT = "http://10.185.30.230:8086/NaNaWeb/services/WorkflowService?wsdl";

        //flag for by Pass change DB, always use UAT ,only for 110.1
        public static bool isUAT = true;

        //flag for by pass EasyFlow approve
        public static bool passEFGP = false;

        //flag for change to FDC 測試完一定要改回來
        public static bool isFDC = false;
    }
}
