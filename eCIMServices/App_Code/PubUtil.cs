using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections; //20161108 add by chuck for sync elton code
using System.Configuration;
using System.Net.Mail;    //20161108 add by chuck for sync elton code
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Web.Configuration;
using System.Web.Services;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Threading;

//******20161108 add by chuck for sync elton code******
/*Json.NET相關的命名空間*/
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
//******20161108 add by chuck for sync elton code******

/// <summary>
/// Summary description for PubUtil
/// 
/// Version: 20140703.001
/// </summary>
public class PubUtil
{
    // For Logon Domain User Only
    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern bool LogonUser(String lpszUsername, String lpszDomain, String lpszPassword,
        int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public extern static bool CloseHandle(IntPtr handle);

    [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public extern static bool DuplicateToken(IntPtr ExistingTokenHandle,
        int SECURITY_IMPERSONATION_LEVEL, ref IntPtr DuplicateTokenHandle);

    private const char returnKey = '\n';
    private static WindowsImpersonationContext impersonatedUser = null;   //mark add 20110308001

    //******20161108 add by chuck for sync elton code******
    private static string _AASReturn = "success";   //20160607 Added by Elton for AAS change return type from "" to "success"

    static public bool setLogonDomain()
    {
        PubUtil.LogonDomain("tw", "admt5", "T5@16847968");
        return true;
    }
    //******20161108 add by chuck for sync elton code******

    //For Single Sign On Domain, add by Elton 20090309
    static public bool setLogonDomain(string Domain, string User, string Pwd)
    {
        PubUtil.LogonDomain(Domain, User, Pwd);
        return true;
    }

    //For Single Sign On Domain, add by Elton 20090309, 0316 Test OK
    static private bool LogonDomain(string Domain, string User, string Pwd)
    {
        try
        {
            IntPtr tokenHandle = new IntPtr(0);
            IntPtr dupeTokenHandle = new IntPtr(0);
            const int LOGON32_PROVIDER_DEFAULT = 0;

            //This parameter causes LogonUser to create a primary token.
            const int LOGON32_LOGON_INTERACTIVE = 9;
            tokenHandle = IntPtr.Zero;

            if (Domain != "" & User != "")//& Pwd != "")
            {
                if (LogonUser(User, Domain, Pwd, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, ref tokenHandle))
                {
                }
                else
                {
                    throw new Exception("Logon failed, pls check config file!");
                }
            }
            else
            {
                User = ConfigurationManager.AppSettings["User"].ToString();//mark add 20090106001
                Pwd = ConfigurationManager.AppSettings["Pwd"].ToString();//mark add 20090106001
                Domain = ConfigurationManager.AppSettings["Domain"].ToString();//mark add 20090106001
                //if (LogonUser("admins", "USTC", "*ej/ vu3@z8 h96*", LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, ref tokenHandle)) //mark by Mark 20090106001
                if (LogonUser(User, Domain, Pwd, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, ref tokenHandle))//mark add 20090106001
                {
                }
                else
                {
                    throw new Exception("Logon failed, pls check config file!");
                }
            }

            WindowsIdentity newId = new WindowsIdentity(tokenHandle);
            //WindowsImpersonationContext impersonatedUser = newId.Impersonate();   //mark by Mark 20110308001
            impersonatedUser = newId.Impersonate(); //mark add 20110308001
            return true;
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "LogonDomain", Domain, User, Pwd);
            throw ex;
        }
    }
    //========= mark add 20110308001 ==========
    static public void ReleaseLogonDomain()
    {
        //impersonatedUser.Undo();
    }

    public static string GetProcessWaferCount(string sLotNo, string sRSC, string sRecipe)
    {
        try
        {
            string sSQL = "";
            string FullPath = "";
            string sProcessWaferCount = "0";

            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton
            sSQL = "Select * from att_ecim.ecimservice_eqpinfo where name= '" + sRSC + "' and TYPE =' '";
            ds = sObjService.AMSDBQuery(sSQL);   //0:EQPName 1:IP 2:Folder 3:LoginName 4:LoginPWD 5: File 6:Case 7: Type

            switch (ds.Tables[0].Rows[0][6].ToString())
            {
                case "WaferCount":
                case "SUSS":    //Coater , Aligner

                    //  if (!isExistValitationTable(sLotNo, sRSC)) { return ""; }

                    FullPath = @"\\" + ds.Tables[0].Rows[0][1].ToString() + @"\" + ds.Tables[0].Rows[0][2].ToString();
                    PubUtil.setLogonDomain(FullPath, ds.Tables[0].Rows[0][3].ToString(), ds.Tables[0].Rows[0][4].ToString());
                    //if (sLotNo.IndexOf('@') > 1)
                    //{
                    //PubUtil.WriteLog("CheckMassOut", sRSC, sLotNo, "Lot include @!!", "return");
                    //    return "";
                    //}

                    //FullPath=@"F:\\Log\\SCB\\"; // For Local test
                    PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Start");


                    if (sRSC.IndexOf("BSCB") > -1)
                    {
                        sProcessWaferCount = T1_SCB_Parser(FullPath, sLotNo, sRSC, sRecipe);
                    }
                    else if (sRSC.IndexOf("BSTE") > -1)
                    {
                        sProcessWaferCount = T1_Step_Parser(FullPath, sLotNo, sRSC, sRecipe);
                    }
                    else
                    {
                        sProcessWaferCount = T5_COA_ALI_Parser(FullPath, sLotNo, sRSC, sRecipe);

                        //add by scott on 20180306
                        try
                        {
                            if (sRSC.IndexOf("T5-ACOA05") > -1 || sRSC.IndexOf("T5-ACOA06") > -1)
                            {
                                if (sRecipe.EndsWith("-2") && int.Parse(sProcessWaferCount) != 0)
                                {
                                    PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Recipe Name is - " + sRecipe + "&& Count1 = " + sProcessWaferCount);
                                    sProcessWaferCount = (int.Parse(sProcessWaferCount) / 2).ToString();
                                    PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Recipe Name is - " + sRecipe + "&& Count2 = " + sProcessWaferCount);
                                }
                                else if (sRecipe.EndsWith("-3") && int.Parse(sProcessWaferCount) != 0)
                                {
                                    PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Recipe Name is - " + sRecipe + "&& Count1 = " + sProcessWaferCount);
                                    sProcessWaferCount = (int.Parse(sProcessWaferCount) / 3).ToString();
                                    PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Recipe Name is - " + sRecipe + "&& Count2 = " + sProcessWaferCount);
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            PubUtil.WriteLog("Exception", "2 Coter or 3 Coater", sRSC, sLotNo, ex.Message);
                            return "";
                        }
                    }
                    break;

                case "TryMax":  //Descum
                    FullPath = @"\\" + ds.Tables[0].Rows[0][1].ToString() + @"\" + ds.Tables[0].Rows[0][2].ToString();
                    PubUtil.setLogonDomain(FullPath, ds.Tables[0].Rows[0][3].ToString(), ds.Tables[0].Rows[0][4].ToString());

                    DataSet TryMaxds = new DataSet();
                    sSQL = "select bumprecipecode from insitedev.amkaas_lot_validate where CARRIER = '" + sLotNo + "' and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
                    TryMaxds = sObjService.AMSDBQuery(sSQL);

                    //if (TryMaxds != null)
                    if (TryMaxds != null && TryMaxds.Tables.Count > 0) //add by chuck for by pass manual lot on 20150714
                    {
                        sRecipe = TryMaxds.Tables[0].Rows[0][0].ToString();
                        PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Recipe Name is - " + sRecipe);
                    }
                    else
                    {
                        PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "get Recipe fail - " + sRecipe);
                        return "";
                    }

                    PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Start");
                    sProcessWaferCount = TryMax_Process_End_Parse(FullPath, sLotNo, sRecipe).ToString();
                    break;

                case "Athlete": //BMBM, BMBR
                    //Get Recipe
                    DataSet Athleteds = new DataSet();
                    sSQL = "select bumprecipecode from insitedev.amkaas_lot_validate where CARRIER = '" + sLotNo + "' and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
                    Athleteds = sObjService.AMSDBQuery(sSQL);

                    //if (Athleteds != null)
                    if (Athleteds != null && Athleteds.Tables.Count > 0)  //add by chuck for ehance on 20150313
                    {
                        sRecipe = Athleteds.Tables[0].Rows[0][0].ToString();
                        PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Recipe Name is - " + sRecipe);
                    }
                    else
                    {
                        PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "get Recipe fail " + sRecipe);
                        return "";
                    }

                    //Copy file from EQP. to T1review02 add by chuck on 20141209
                    FullPath = @"\\" + ds.Tables[0].Rows[0][1].ToString() + @"\" + ds.Tables[0].Rows[0][2].ToString();
                    PubUtil.setLogonDomain(FullPath, ds.Tables[0].Rows[0][3].ToString(), ds.Tables[0].Rows[0][4].ToString());
                    try
                    {
                        CopyFile(sLotNo, sRSC, sRecipe);
                        //CopyFile(sLotNo, sRSC);
                    }
                    catch (Exception ex)
                    {
                        PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "CopyFile exception : " + ex.Message);
                    }

                    FullPath = @"\\" + ds.Tables[0].Rows[0][1].ToString() + @"\" + ds.Tables[0].Rows[0][2].ToString();
                    PubUtil.setLogonDomain(FullPath, ds.Tables[0].Rows[0][3].ToString(), ds.Tables[0].Rows[0][4].ToString());

                    //FullPath = @"D:\Log\Ahtlete\";// +FileName;  //Test Only set path here 


                    //Get Process Wafer count
                    PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Get Process Wafer Count Start");

                    sProcessWaferCount = T1Ahtlete_Parser(FullPath, sLotNo, sRSC, sRecipe);
                    break;

                case "ByPort":

                    DataSet ByPortds = new DataSet();
                    sSQL = "select UOMNAME,CARRIER from insitedev.amkaas_lot_validate where portno = '" + sLotNo + "' and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
                    ByPortds = sObjService.AMSDBQuery(sSQL);

                    //if (ByPortds != null)
                    if (ByPortds != null && ByPortds.Tables.Count > 0) //add by chuck for by pass manual lot on 20150714
                    {
                        sProcessWaferCount = ByPortds.Tables[0].Rows[0][0].ToString();
                        sLotNo = ByPortds.Tables[0].Rows[0][1].ToString();
                        //  if (!isExistValitationTable(sLotNo, sRSC)) { return ""; }
                    }
                    else
                    {
                        PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Get Process Wafer Count fail.");
                        return "";
                    }
                    break;

                case "ByLot":

                    if (!isExistValitationTable(sLotNo, sRSC)) { return ""; } //Add by Stanley 20131001
                    int iWaferFile = 0;
                    DataSet ByLotds = new DataSet();
                    sSQL = "select UOMNAME from insitedev.amkaas_lot_validate where CARRIER = '" + sLotNo + "' and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
                    ByLotds = sObjService.AMSDBQuery(sSQL);

                    if (ByLotds != null)
                    {
                        sProcessWaferCount = ByLotds.Tables[0].Rows[0][0].ToString();
                        if (sRSC.Equals("T5-BSPT04"))
                        {
                            iWaferFile = Get_SPT_WaferFile(sLotNo, sRSC, ds.Tables[0].Rows[0]);
                            if (iWaferFile > int.Parse(sProcessWaferCount)) sProcessWaferCount = iWaferFile.ToString();
                        }

                        //20161208 add by Elton for check BSPT TM Pressure data log
                        try
                        {
                            if (sRSC.Equals("T1_BP_BSPT002") || sRSC.Equals("T1_BP_BSPT004") || sRSC.Equals("T1_BP_BSPT005"))
                            {
                                getTMPressureData(sLotNo, sRSC);
                            }
                        }
                        catch (Exception ex)
                        {
                        }

                        //20181221 add by Elton for UAT check FDC Wafer Count log
                        try
                        {
                            string sFDCwaferCount = "";
                            if (!ds.Tables[0].Rows[0][2].ToString().Equals(""))
                            {
                                sFDCwaferCount = getFDCWaferCount(sLotNo, sRSC, ds.Tables[0].Rows[0][2].ToString());
                            }

                            if (sFDCwaferCount.Equals(sProcessWaferCount))
                            {
                                PubUtil.WriteLog("eCIMWebServiceLog", "GetFDCWafer Count", sRSC, sLotNo, "FDC Wafer Count is same..");
                            }
                            else
                            {
                                PubUtil.WriteLog("Error", "GetFDCWafer Count", sRSC, sLotNo, "FDC Wafer Count is not same!!");
                            }

                        }
                        catch (Exception)
                        {

                        }

                    }
                    else
                    {
                        PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, "Get Process Wafer Count fail.");
                        return "";
                    }
                    break;

                // 20181211 add by Elton for check FDC data for Wafer Count
                case "ByFDC":

                    if (!isExistValitationTable(sLotNo, sRSC)) { return ""; }
                    sProcessWaferCount = getFDCWaferCount(sLotNo, sRSC, ds.Tables[0].Rows[0][2].ToString());
                    break;
            }

            string sDelimeter = "|";
            string sReturnValue = sLotNo + sDelimeter + sRSC + sDelimeter + sRecipe + sDelimeter + sProcessWaferCount;
            PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, sReturnValue);
            return sReturnValue;
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "GetProcessWaferCount", sRSC, sLotNo, ex.Message);
            return "";
        }
    }
    //public static string GetBMBREdcData(string sLotNo, string sRSC)
    public static string GetBMBREdcData(string sLotNo, string sRSC, string sRecipe)
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "GetBMBREdcData", sRSC, sLotNo, "eCIM Get BMBR EDC data Start.");
        try
        {
            string FullPath = "";
            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton
            sSQL = "Select * from att_ecim.ecimservice_eqpinfo where name= '" + sRSC + "' and TYPE =' '";
            ds = sObjService.AMSDBQuery(sSQL);   //0:EQPName 1:IP 2:Folder 3:LoginName 4:LoginPWD 5: File

            FullPath = @"\\" + ds.Tables[0].Rows[0][1].ToString() + @"\" + ds.Tables[0].Rows[0][2].ToString();
            PubUtil.setLogonDomain(FullPath, ds.Tables[0].Rows[0][3].ToString(), ds.Tables[0].Rows[0][4].ToString());

            //FullPath = @"D:\Log\Ahtlete\";// +FileName;  //Test Only set path here 

            return T1Ahtlete_Parser(FullPath, sLotNo, sRSC, sRecipe);
            //return "";

        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "GetBMBREdcData", sRSC, sLotNo, ex.Message);
            throw;
        }
    }
    //private static string T5BDES01_RPTParser(string[] sDataContent,string LOTID,string Recipe)
    private static string T5BDES01_RPTParser(string[] sDataContent)
    {
        try
        {
            bool bStartRead = false;
            int iStartIdx = 0;
            for (int i = 0; i < sDataContent.Length; i++)
            {
                if (sDataContent[i] != null)
                {
                    if (bStartRead && sDataContent[i].IndexOf("NA") < 0)
                    {
                        string[] sData = sDataContent[i].Split('|');
                        if (sData[0] != "")
                        {
                            if (Convert.ToInt16(sData[0]) <= 25)
                            {
                                //if (sData[2].Remove(sData[2].IndexOf('_')).Trim().Equals(LOTID) && sData[3].Trim().ToUpper().Equals(Recipe + ".RCP"))
                                iStartIdx += 1;
                            }
                        }
                    }
                    if (sDataContent[i].ToUpper().IndexOf("SLOT |") > -1)
                    {
                        bStartRead = true;
                    }
                }
            }
            return iStartIdx.ToString();
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "T5BDES01_RPTParser", "", "", ex.Message);
            throw ex;
        }
    }

    private static string T5BDES01_RPTParser(string FullPath, string sLotID, string sRSC, string sRecipe)
    {
        bool bFindFile = false;
        string FileName = "";
        string FileType = "";
        StreamReader fnr;
        string[] sDataContent;
        int iIdx = 0;
        DirectoryInfo dir;
        DateTime dt = DateTime.Now.AddMinutes(-25);
        try
        {
            FileType = "rpt";
            FileName = "*" + sLotID + "." + FileType;
            //FileName = "00-11-L-BU1460365.rpt";

            //FullPath = @"\\Review0\other\MFG\EDC\rpt\";// For test only
            FullPath = FullPath + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("0#") + @"\" + DateTime.Now.Day.ToString("0#") + @"\";// +FileName;

            //FullPath = @"\\Review0\other\MFG\EDC\rpt\";         //For test only
            //PubUtil.setLogonDomain("ustc", "10869", "amkor@t5");//For test only

            dir = new DirectoryInfo(FullPath);

            foreach (FileInfo filename in dir.GetFiles(FileName))
            {
                if (dt < filename.LastWriteTime)
                {
                    dt = filename.LastWriteTime;
                    //FullPath = @"\\Review0\other\MFG\EDC\rpt\\" + filename.Name;    //For test only
                    FullPath += filename.Name;
                }
            }

            //FileName = "BUT2210159@eCIMTest.prt";   // For test only
            //FullPath = @"\\Review0\other\MFG\EDC\rpt\" + FileName;    //For test only

            FullPath = FullPath.Replace("\\", @"\");
            if (File.Exists(FullPath))
            {
                bFindFile = true;
            }

            if (!bFindFile)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "T5BDES01_RPTParser", sRSC, sLotID, "File not Find! " + FullPath);
                return "Can not find file of " + FullPath;
            }

            File.Copy(FullPath, @"D:\EQP_Data\" + sRSC + @"\" + FileName, true);

            FullPath = @"D:\EQP_Data\" + sRSC + @"\" + FileName;

            fnr = new StreamReader(FullPath);
            sDataContent = new string[40];

            while (fnr.Peek() >= 0)
            {
                string stmp = fnr.ReadLine().Replace("\n", "");
                stmp = stmp.Replace("\r\n", "");
                if (stmp != "")
                {
                    sDataContent[iIdx] = stmp;
                    iIdx += 1;
                }
            }
            fnr.Close();

            bool bStartRead = false;
            int iStartIdx = 0;
            for (int i = 0; i < sDataContent.Length; i++)
            {
                if (sDataContent[i] != null)
                {
                    if (bStartRead && sDataContent[i].IndexOf("NA") < 0)
                    {
                        string[] sData = sDataContent[i].Split('|');
                        if (sData[0] != "")
                        {
                            if (Convert.ToInt16(sData[0]) <= 25)
                            {
                                //if (sData[2].Remove(sData[2].IndexOf('_')).Trim().Equals(LOTID) && sData[3].Trim().ToUpper().Equals(Recipe + ".RCP"))
                                iStartIdx += 1;
                            }
                        }
                    }
                    if (sDataContent[i].ToUpper().IndexOf("SLOT |") > -1)
                    {
                        bStartRead = true;
                    }
                }
            }
            return iStartIdx.ToString();
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "T5BDES01_RPTParser", sRSC, sLotID, ex.Message);
            throw ex;
        }
    }

    private static string T5_COA_ALI_Parser(string FullPath, string sLotID, string sRSC, string sRecipe)
    {
        //bool bFindFile = false; //20161108 add by chuck for sync elton code
        string FileName = "";
        string FileType = "";
        StreamReader fnr;
        string[] sDataContent;
        int iIdx = 0;
        DirectoryInfo dir;
        DateTime dt = DateTime.Now.AddMinutes(-30);

        try
        {
            FileHandle(FullPath); //Append by Stanley 20130719 for file handle.
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "FileHandle", sRSC, sLotID, ex.Message);
        }

        try
        {

            FileType = "prt";
            FileName = sLotID + "_CAR*." + FileType;

            //FullPath = @"\\Review0\other\MFG\EDC\rpt\";         //For test only
            //PubUtil.setLogonDomain("ustc", "10869", "amkor@t5");//For test only

            dir = new DirectoryInfo(FullPath);

            //test 
            FileInfo[] aFileInfo = dir.GetFiles(FileName);
            if (aFileInfo.Length == 0)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "T5_COA_ALI_Parser", sRSC, sLotID, "Get 0 file list!" + FullPath + "/" + FileName);
                return "Get 0 file list - " + FullPath;
            }
            if (aFileInfo.Length == 1)
            {
                FullPath += aFileInfo[0].Name;
                FileName = aFileInfo[0].Name;
            }
            else
            {
                string sFileName = "";
                for (int i = 0; i < aFileInfo.Length; i++)//20121115 add by Elton for function enhancement
                {
                    if (dt < aFileInfo[i].LastWriteTime)
                    {
                        dt = aFileInfo[i].LastWriteTime;
                        //FullPath = @"\\Review0\other\MFG\EDC\rpt\\" + filename.Name;    //For test only
                        sFileName = aFileInfo[i].Name;
                    }
                }
                FileName = sFileName;
                FullPath += sFileName;
            }

            FullPath = FullPath.Replace("\\", @"\");

            PubUtil.WriteLog("eCIMWebServiceLog", "T5_COA_ALI_Parser", sRSC, sLotID, FullPath + " File Copy Start!");
            if (!Directory.Exists(@"D:\EQP_Data\" + sRSC + @"\")) Directory.CreateDirectory(@"D:\EQP_Data\" + sRSC + @"\");
            File.Copy(FullPath, @"D:\EQP_Data\" + sRSC + @"\" + FileName, true);
            if (!Directory.Exists(@"D:\EQP_Data\" + sRSC + @"\backup\")) Directory.CreateDirectory(@"D:\EQP_Data\" + sRSC + @"\backup\");
            File.Copy(FullPath, @"D:\EQP_Data\" + sRSC + @"\backup\" + FileName, true);  //Copy file to back up
            PubUtil.WriteLog("eCIMWebServiceLog", "T5_COA_ALI_Parser", sRSC, sLotID, FullPath + " File Copy OK!");

            FullPath = @"D:\EQP_Data\" + sRSC + @"\" + FileName;

            fnr = new StreamReader(FullPath);
            sDataContent = new string[68];

            while (fnr.Peek() >= 0)
            {
                string stmp = fnr.ReadLine().Replace("\n", "");
                stmp = stmp.Replace("\r\n", "");
                if (stmp != "")
                {
                    sDataContent[iIdx] = stmp;
                    iIdx += 1;
                }
            }
            fnr.Close();

            bool bStartRead = false;
            int iStartIdx = 0;
            for (int i = 0; i < sDataContent.Length; i++)
            {
                if (sDataContent[i] != null)
                {
                    if (bStartRead && sDataContent[i].ToUpper().IndexOf("OK") > 0)
                    {
                        if (sDataContent[i].IndexOf(sLotID) > 0)
                        {
                            if (sDataContent[i].IndexOf(sRecipe) > 0)
                            {
                                iStartIdx += 1;
                            }
                        }
                    }
                    if (sDataContent[i].ToUpper().IndexOf("SLOT NAME") > -1)
                    {
                        bStartRead = true;
                    }
                }
            }
            File.Delete(@"D:\EQP_Data\" + sRSC + @"\" + FileName);

            //add by Jim for check 2COA 3COA return wafer count 20181120
            try
            {
                if (sRSC.IndexOf("T5-ACOA05") > -1 || sRSC.IndexOf("T5-ACOA06") > -1)
                {
                    if (ConfigurationManager.AppSettings["2COAT"].ToString().IndexOf(sRecipe) > -1)
                    {
                        string[] COA2 = ConfigurationManager.AppSettings["2COAT"].ToString().Split('/');
                        foreach (string temp in COA2)
                        {
                            if (sRecipe.Equals(temp))
                            {
                                PubUtil.WriteLog("eCIMWebServiceLog", "T5_COA_ALI_Parser", sRSC, sLotID, "Recipe = " + sRecipe + ", is 2COAT");
                                return ((int)(iStartIdx / 2)).ToString();
                            }
                        }
                    }
                    else if (ConfigurationManager.AppSettings["3COAT"].ToString().IndexOf(sRecipe) > -1)
                    {
                        string[] COA3 = ConfigurationManager.AppSettings["3COAT"].ToString().Split('/');
                        foreach (string temp in COA3)
                        {
                            if (sRecipe.Equals(temp))
                            {
                                PubUtil.WriteLog("eCIMWebServiceLog", "T5_COA_ALI_Parser", sRSC, sLotID, "Recipe = " + sRecipe + ", is 3COAT");
                                return ((int)(iStartIdx / 3)).ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PubUtil.WriteLog("Exception", "T5_COA_ALI_Parser check 2CO / 3CO", sRSC, sLotID, ex.Message);
            }

            return iStartIdx.ToString();
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "T5_COA_ALI_Parser", sRSC, sLotID, ex.Message);
            throw ex;
        }
    }

    private static string T1_SCB_Parser(string FullPath, string sLotID, string sRSC, string sRecipe)
    {
        bool rCount = true;
        string FileName = "";
        StreamReader fnr;
        string[] sDataContent;
        int iIdx = 0;
        DirectoryInfo dir;
        DateTime dt = DateTime.Now.AddMinutes(-30);

        try
        {
            FileHandle(FullPath); //Append by Stanley 20130719 for file handle.
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "FileHandle", sRSC, sLotID, ex.Message);
        }

        try
        {
            FileName = "PJ_" + sLotID + "*.csv";
            dir = new DirectoryInfo(FullPath);
            FileInfo[] aFileInfo = dir.GetFiles(FileName);
            bool getFileFlag = false;
            int loopCount = 1;
            while (!getFileFlag) //Append by Stanley for Get file retry function. 20140828
            {
                aFileInfo = dir.GetFiles(FileName);
                if (loopCount > 3)
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "T1_SCB_Parser", sRSC, sLotID, "Get 0 file list! ==>" + FullPath + "/" + FileName);
                    return "Get 0 file list - " + FullPath;
                }
                if (aFileInfo.Length == 0)
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "T1_SCB_Parser", sRSC, sLotID, "Get file retry:" + loopCount.ToString());
                    Thread.Sleep(int.Parse(ConfigurationManager.AppSettings["getFileDelay"].ToString()));
                    loopCount++;
                }
                else
                {
                    getFileFlag = true;
                }
            }

            if (aFileInfo.Length == 1)
            {
                FullPath += aFileInfo[0].Name;
                FileName = aFileInfo[0].Name;
            }
            else
            {
                string sFileName = "";
                for (int i = 0; i < aFileInfo.Length; i++)//20121115 add by Elton for function enhancement
                {
                    if (dt < aFileInfo[i].LastWriteTime)
                    {
                        dt = aFileInfo[i].LastWriteTime;
                        sFileName = aFileInfo[i].Name;
                    }
                }
                FileName = sFileName;
                FullPath += sFileName;
            }

            FullPath = FullPath.Replace("\\", @"\");

            if (!Directory.Exists(@"D:\EQP_Data\" + sRSC + @"\")) Directory.CreateDirectory(@"D:\EQP_Data\" + sRSC + @"\");
            File.Copy(FullPath, @"D:\EQP_Data\" + sRSC + @"\" + FileName, true);
            if (!Directory.Exists(@"D:\EQP_Data\" + sRSC + @"\backup\")) Directory.CreateDirectory(@"D:\EQP_Data\" + sRSC + @"\backup\");
            File.Copy(FullPath, @"D:\EQP_Data\" + sRSC + @"\backup\" + FileName, true);  //Copy file to back up
            PubUtil.WriteLog("eCIMWebServiceLog", "T1_SCB_Parser", sRSC, sLotID, FullPath + " File Copy OK!");

            FullPath = @"D:\EQP_Data\" + sRSC + @"\" + FileName;

            fnr = new StreamReader(FullPath);
            sDataContent = new string[26];

            while (rCount || fnr.Peek() >= 0)
            {
                string stmp = fnr.ReadLine();
                if (stmp != null)
                {
                    stmp = stmp.Replace("\n", "");
                    stmp = stmp.Replace("\r\n", "");
                    if (stmp != "")
                    {
                        Array.Resize(ref sDataContent, iIdx + 1);
                        sDataContent[iIdx] = stmp;
                        iIdx += 1;
                    }
                    if (iIdx > 25) rCount = false;
                }
                else
                {
                    rCount = false;
                }
            }
            fnr.Close();

            int iStartIdx = 0;
            string[] RunData;
            RunData = new string[13]; // total 13 columns in csv file

            for (int i = 1; i < sDataContent.Length; i++)
            {
                if (sDataContent[i] != null || sDataContent[i] != "")
                {
                    RunData = sDataContent[i].Split(',');
                    if (RunData[10] != "") // Check not empty will be wafer count + 1
                    {
                        iStartIdx += 1;
                    }
                }
            }
            File.Delete(@"D:\EQP_Data\" + sRSC + @"\" + FileName);
            return iStartIdx.ToString();
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "T1_SCB_Parser", sRSC, sLotID, ex.Message);
            throw ex;
        }
    }
    //copy from T5BALI_RPTParser 20130204
    private static string T1Ahtlete_Parser(string FullPath, string sLotID, string sRSC, string sRecipe)
    {
        //bool bFindFile = false; //20161108 add by chuck for sync elton code
        string FileName = "";
        StreamReader fnr;
        string[] sDataContent;
        int iIdx = 0;
        DirectoryInfo dir;
        DateTime dt = DateTime.Now.AddMinutes(-60);

        try
        {
            FileHandle(FullPath); //Append by Stanley 20130802 for file handle.
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "FileHandle", sRSC, sLotID, ex.Message);
            return "";
        }

        try
        {
            FileName = sRecipe + "-" + sLotID + ".L";
            dir = new DirectoryInfo(FullPath);

            FileInfo[] aFileInfo = dir.GetFiles(FileName);
            if (aFileInfo.Length == 0)
            {
                FileInfo[] sFiles = dir.GetFiles();
                Boolean bFlag = true; // For find file by LotID. Append by Stanley 20131003
                for (int i = 0; i < sFiles.Length; i++)
                {
                    if (sFiles[i].Name.IndexOf(sLotID) > -1)
                    {
                        Array.Resize(ref aFileInfo, 1);
                        aFileInfo[0] = sFiles[i];
                        sRecipe = sFiles[i].Name.Split('-')[0];
                        bFlag = false;
                        break;
                    }
                }
                if (bFlag)
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "T1Ahtlete_Parser", sRSC, sLotID, "Get 0 file list! ==>" + FullPath + "/" + FileName);
                    return "Get 0 file list - " + FullPath;
                }
            }

            if (aFileInfo.Length == 1)
            {
                FullPath += aFileInfo[0].Name;
                FileName = aFileInfo[0].Name;
            }
            else
            {
                string sFileName = "";
                for (int i = 0; i < aFileInfo.Length; i++)//20121115 add by Elton for function enhancement
                {
                    if (dt < aFileInfo[i].LastWriteTime)
                    {
                        dt = aFileInfo[i].LastWriteTime;
                        sFileName = aFileInfo[i].Name;
                    }
                }
                FileName = sFileName;
                FullPath += sFileName;
            }

            FullPath = FullPath.Replace("\\", @"\");

            fnr = new StreamReader(FullPath);
            //sDataContent = new string[200];
            sDataContent = new string[1000]; //20140918 Roger request modify

            while (fnr.Peek() >= 0)
            {
                string stmp = fnr.ReadLine().Replace("\n", "");
                stmp = stmp.Replace("\r\n", "");
                if (stmp != "")
                {
                    //Array.Resize(ref sDataContent, iIdx + 1);
                    sDataContent[iIdx] = stmp;
                    iIdx += 1;
                }
            }
            fnr.Close();

            int iStartIdx = 0;
            int Total = 0;
            int[] aReturn = new int[9];
            String sParserResult = "";
            int Miss = 0;
            int Extra = 0;
            int Shift = 0;
            int AfterTotal = 0;
            int AfterMiss = 0;
            int AfterExtra = 0;
            int AfterShift = 0;

            if (sRSC.IndexOf("BMBR") > 0) //EDC
            {
                for (int i = 0; i < sDataContent.Length; i++)
                {
                    if (sDataContent[i] != null)
                    {
                        if (sDataContent[i].ToUpper().IndexOf(sLotID.ToUpper()) > 0)
                        {
                            if (sDataContent[i].ToUpper().IndexOf(sRecipe.ToUpper()) > 0)
                            {
                                iStartIdx += 1;
                                string[] aa = sDataContent[i].ToUpper().Split(',');
                                if (aa[6] == "" || aa[2].ToUpper() != sRecipe.ToUpper() || aa[3].ToUpper() != sLotID.ToUpper()) continue;//get edc result data from log file
                            }
                        }

                        string[] a = sDataContent[i].ToUpper().Split(',');
                        Total += int.Parse(a.GetValue(6).ToString());
                        Miss += int.Parse(a.GetValue(7).ToString());
                        Extra += int.Parse(a.GetValue(8).ToString());
                        Shift += int.Parse(a.GetValue(9).ToString());
                        AfterMiss += int.Parse(a.GetValue(11).ToString());
                        AfterExtra += int.Parse(a.GetValue(12).ToString());
                        AfterShift += int.Parse(a.GetValue(13).ToString());
                        AfterTotal = AfterMiss + AfterExtra + AfterShift;
                        aReturn[1] = Total;
                        aReturn[2] = Miss;
                        aReturn[3] = Extra;
                        aReturn[4] = Shift;
                        aReturn[5] = AfterTotal;
                        aReturn[6] = AfterMiss;
                        aReturn[7] = AfterExtra;
                        aReturn[8] = AfterShift;
                    }
                }
            }

            int[] WaferCount = new int[25];
            for (int i = 0; i < sDataContent.Length; i++)
            {
                if (sDataContent[i] != null)
                {
                    if (sDataContent[i].ToUpper().IndexOf(sLotID.ToUpper()) > 0 && sDataContent[i].ToUpper().IndexOf(sRecipe.ToUpper()) > 0)
                    {
                        if (sDataContent[i].ToUpper().IndexOf("MOVE-OUT") > -1)
                        {
                            string[] aa = sDataContent[i].ToUpper().Split(',');

                            //                            int iSlot;
                            try
                            {
                                //                                iSlot = int.Parse(aa[4]); //BMBR Log format rule
                                string[] sTemp = aa[4].Split('-');  // Filter not normal char for Slot.Append by Stanley 20131003
                                if (sTemp.Length > 1) { aa[4] = sTemp[1]; }

                                WaferCount[int.Parse(aa[4]) - 1] = 1; // From 0 start
                            }
                            catch (Exception)
                            {
                                //                               iSlot = int.Parse(aa[6]); //BMBM Log format rule
                                string[] sTemp = aa[6].Split('-');  // Filter not normal char for Slot.Append by Stanley 20131003
                                if (sTemp.Length > 1) { aa[6] = sTemp[1]; }

                                WaferCount[int.Parse(aa[6]) - 1] = 1; // From 0 start
                            }
                            PubUtil.WriteLog("eCIMWebServiceLog", "T1Ahtlete_Parser", sRSC, sLotID, Total.ToString());
                        }
                    }
                }
            }

            for (int i = 0; i < WaferCount.Length; i++)
            {
                if (WaferCount[i] == 1) { aReturn[0] += 1; }
            }

            for (int i = 0; i < aReturn.Length; i++)
            {
                sParserResult += aReturn[i].ToString() + "|";
            }

            PubUtil.WriteLog("eCIMWebServiceLog", "T1Ahtlete_Parser", sRSC, sLotID, sParserResult);
            return sParserResult;
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "T1Ahtlete_Parser", sRSC, sLotID, ex.Message);
            return ex.Message;
        }
    }

    public static void WriteLog(String LogType, string FunctionName, string EQPID, string Info, string msg)
    {
        try
        {
            //PubUtil.setLogonDomain("10.185.110.3", "administrator", "connt1$168");    //T1 using
            //PubUtil.setLogonDomain("10.185.126.4", "administrator", "att5@ecim");       //T5 using

            string Date = System.DateTime.Now.ToString("yyyyMMdd");
            string DateTime = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string subDate = Date.Substring(0, 4);
            string sLogPath = ConfigurationManager.AppSettings["LOG_PATH"].ToString();

            //check log folder
            if (!Directory.Exists(sLogPath + LogType + @"\" + subDate))
                Directory.CreateDirectory(sLogPath + LogType + @"\" + subDate);
            //check log file
            if (!File.Exists(sLogPath + LogType + @"\" + subDate + @"\" + Date + ".log"))
                File.WriteAllText(sLogPath + LogType + @"\" + subDate + @"\" + Date + ".log", DateTime + "|" + FunctionName + "|" + EQPID + "|" + Info + "|" + msg + "\r\n");
            else
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(sLogPath + LogType + @"\" + subDate + @"\" + Date + ".log", true))
                {
                    file.WriteLine(DateTime + "|" + FunctionName + "|" + EQPID + "|" + Info + "|" + msg);
                }
            }

        }
        catch (Exception)
        {
            //throw;
        }
    }
    //public static string GetAlignerEdcData(string sType)
    public static DataTable GetAlignerEdcData(string sRSC, string sType)
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "GetAlignerEdcData", sRSC, sType, "eCIM Get Aligner EDC data Start. ");
        try
        {
            string sSQL = "";
            string sFile = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton
            sSQL = "Select * from att_ecim.ecimservice_eqpinfo where name= '" + sRSC + "' and TYPE= 'EDC'";
            ds = sObjService.AMSDBQuery(sSQL);   //0:EQPName 1:IP 2:Folder 3:LoginName 4:LoginPWD 5: File 6: Select 7: Type
            sFile = ds.Tables[0].Rows[0][5].ToString();
            string FullPath = @"\\" + ds.Tables[0].Rows[0][1].ToString() + @"\" + ds.Tables[0].Rows[0][2].ToString() + DateTime.Now.DayOfWeek + @"\";
            PubUtil.setLogonDomain(FullPath, ds.Tables[0].Rows[0][3].ToString(), ds.Tables[0].Rows[0][4].ToString());

            return Aligner_Parse(FullPath, sType, sFile);

        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "GetAlignerEdcData", sRSC, sType, ex.Message);
            throw;
        }
    }
    //copy from T1BMBR_RPTParser 20130528
    private static DataTable Aligner_Parse(string FullPath, string sType, string sFile)
    {
        //bool bFindFile = false; //20161108 add by chuck for sync elton code
        string FileName = "";
        StreamReader fnr;
        string[] sDataContent;
        double[] dDataContent;
        int iIdx = 0;
        int iNo = 2;
        DirectoryInfo dir;
        try
        {
            FileName = sFile;
            dir = new DirectoryInfo(FullPath);

            FileInfo[] aFileInfo = dir.GetFiles(sFile);
            if (aFileInfo.Length == 0)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "Aligner_Parse", sType, sFile, "Get 0 file list! ==>" + FullPath + "/" + FileName);
            }
            if (aFileInfo.Length == 1)
            {
                FullPath += aFileInfo[0].Name;
                FileName = aFileInfo[0].Name;
            }
            else
            {
                string sFileName = "";
                FileName = sFileName;
                FullPath += sFileName;
            }

            FullPath = FullPath.Replace("\\", @"\");

            int ioutCount = int.Parse(ConfigurationManager.AppSettings["Aligner_Out_Count"]);

            fnr = new StreamReader(FullPath);
            sDataContent = new string[1];
            dDataContent = new double[ioutCount];

            try
            {
                while (fnr.Peek() >= 0)
                {
                    string stmp = fnr.ReadLine().Replace("\n", "");
                    stmp = stmp.Replace("\r\n", "");
                    if (stmp != "") // && stmp.IndexOf("," + iIdx) > -1)
                    {
                        Array.Resize(ref sDataContent, iIdx + 1);
                        sDataContent[iIdx] = stmp;
                        iIdx += 1;
                    }
                }
                fnr.Close();
            }
            catch
            {
                fnr.Close();
            }
            switch (sType)
            {
                case "365"://
                    iNo = 2;
                    break;
                case "405"://
                    iNo = 3;
                    break;
            }

            double dUniformity = 0;
            double dMax = 0;
            double dMin = 0;
            double dType = 0;

            for (int i = 1; i < ioutCount; i++)
            {
                if (sDataContent[i] != null)
                {
                    string[] a = sDataContent[i].Split(',');

                    dType = Convert.ToDouble(a[iNo]);

                    dDataContent[i - 1] = dType;

                    if (dMax == 0)
                    {
                        dMax = dType;
                    }

                    if (dMin == 0)
                    {
                        dMin = dType;
                    }

                    if (dMax < dType)
                    {
                        dMax = dType;
                    }

                    if (dMin > dType)
                    {
                        dMin = dType;
                    }

                }
            }
            dUniformity = Math.Round((dMax - dMin) / (dMax + dMin) * 100, 1);

            dDataContent[ioutCount - 1] = dUniformity;

            return Out_ToTable(dDataContent);
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "Aligner_Parse", sType, sFile, ex.Message);
            return null;
        }
    }
    //public static string GetAlignerEdcData(string sType)
    public static int Get_TryMax_Data(string LotNo, string sRSC, string Recipe)
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "Get_TryMax_Data", sRSC, LotNo, "eCIM Get AMS TryMax Process End data Start!");
        try
        {

            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton
            sSQL = "Select * from att_ecim.ecimservice_eqpinfo where name= '" + sRSC + "' and TYPE =' '";
            ds = sObjService.AMSDBQuery(sSQL);   //0:EQPName 1:IP 2:Folder 3:LoginName 4:LoginPWD 5: File

            string FullPath = @"\\" + ds.Tables[0].Rows[0][1].ToString() + @"\" + ds.Tables[0].Rows[0][2].ToString();
            // FullPath = @"D:\TEST\"; //For Test
            // PubUtil.setLogonDomain(FullPath, ds.Tables[0].Rows[0][3].ToString(), ds.Tables[0].Rows[0][4].ToString());
            int rtValue = TryMax_Process_End_Parse(FullPath, LotNo, Recipe);  //For DataTable

            // string FullPath = @"D:\TEST\"; //For Test

            FileHandle(FullPath); //Append by Stanley 20130719 for file handle. For Test
            PubUtil.WriteLog("eCIMWebServiceLog", "Get_TryMax_Data", sRSC, LotNo, "eCIM Get AMS TryMax Process End data. End!");
            return rtValue;

        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "Get_TryMax_Data", sRSC, LotNo, ex.Message);
            throw;
        }
    }
    //copy from T1Aligner_RPTParser 20130711
    private static int TryMax_Process_End_Parse(string FullPath, string LotNo, string Recipe)
    {
        string[] sDataContent;
        int iIdx = 0;
        int iTotal = 0;
        int iDataBlockLine = int.Parse(ConfigurationManager.AppSettings["DataBlockLine"]); //Setting Read Black Line count by each wafer

        try
        {
            FileHandle(FullPath); //Append by Stanley 20130802 for file handle.
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "FileHandle", "", LotNo, ex.Message);
        }

        try
        {
            DirectoryInfo dir = new DirectoryInfo(FullPath);
            FileInfo[] aFileInfo = dir.GetFiles("*" + LotNo + ".txt");  //Get Path folder all like *"LotNo".txt files
            if (aFileInfo.Length == 0)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "TryMax_Process_End_Parse", LotNo, FullPath, "Get 0 file list! LotID:" + LotNo);
                return 0;
            }
            //Found out lastest datetime of files by LotID for paser
            FileInfo FullPath_Name = null;
            for (int i = 0; i < aFileInfo.Length; i++)
            {
                if (FullPath_Name == null)
                {
                    FullPath_Name = aFileInfo[i];
                }
                else
                {
                    if (FullPath_Name.CreationTime < aFileInfo[i].CreationTime)
                    {
                        FullPath_Name = aFileInfo[i];
                    }
                }
            }

            FullPath += FullPath_Name.Name;

            StreamReader fnr = new StreamReader(FullPath);
            sDataContent = new string[1];
            while (fnr.Peek() >= 0)
            {
                string stmp = fnr.ReadLine().Replace("\n", "");
                stmp = stmp.Replace("\r\n", "");
                if (stmp != "")  //filter not empty line into array object
                {
                    Array.Resize(ref  sDataContent, iIdx + 1); //Dynatic adjust Array size for keep low data line.
                    sDataContent[iIdx] = stmp;
                    iIdx += 1;
                }
            }
            fnr.Close();

            for (int i = 0; i < sDataContent.Length; i++)
            {
                if (sDataContent[i] != null)
                {
                    string[] b = sDataContent[i].Split(':');
                    if (b[0] == "Initial Slot") //Into Black get relation data.
                    {
                        for (int j = 1; j < (iDataBlockLine + 1); j++) //Set Read Black count loop
                        {
                            Boolean bflag = false; // This flag check recipe name is the same with AMS.
                            switch (j)
                            {
                                case 3:
                                    string[] PPID = sDataContent[i + j].Split(':');
                                    if (PPID[0] == "Module recipe")
                                    {
                                        string[] sRCP = PPID[1].Split('(');
                                        if (sRCP[0].Trim() != Recipe)
                                        {
                                            bflag = true;
                                        }
                                    }
                                    break;
                                case 8:
                                    string[] Process = sDataContent[i + j].Split(':');
                                    if (Process[1].Trim() == "Processed")
                                    {
                                        iTotal += 1;
                                    }
                                    else
                                    {
                                        bflag = true;
                                    }
                                    break;
                            }
                            if (bflag)
                            {
                                break; //If true exit Block read.
                            }
                        }
                        i += iDataBlockLine;
                    }
                }
            }
            return iTotal;
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "TryMax_Process_End_Parser", "", "", ex.Message);
            return 0;
        }
    }

    private static DataTable Out_ToTable(double[] data)
    {

        DataTable dtEDC = new DataTable("Table");
        string row_Name = ConfigurationManager.AppSettings["AMSTxn_Function_Row"];  //AMSTxn_Function_Row
        string[] rowName;

        int iAFR = row_Name.Split(',').Length;
        rowName = new string[iAFR];
        rowName = row_Name.Split(',');

        DataColumn Col;
        DataRow row;

        Col = new DataColumn();
        // Col.ColumnName = "No";
        Col.ColumnName = rowName[0];
        dtEDC.Columns.Add(Col);
        Col = new DataColumn();
        //Col.ColumnName ="Data";
        Col.ColumnName = rowName[1];
        dtEDC.Columns.Add(Col);
        Col = new DataColumn();
        //Col.ColumnName ="WID";
        Col.ColumnName = rowName[2];
        dtEDC.Columns.Add(Col);

        DataSet DS = new DataSet();
        DS.Tables.Add(dtEDC);
        DS.DataSetName = ConfigurationManager.AppSettings["AMSTxn_Function"].ToString();  //"AutoEDC";

        for (int i = 0; i < data.Length; i++)
        {
            row = dtEDC.NewRow();
            row[rowName[0]] = i;
            row[rowName[1]] = data[i];
            row[rowName[2]] = "";
            dtEDC.Rows.Add(row);
        }

        // XmlDataDocument xdd = new XmlDataDocument(DS);

        // return xdd.InnerXml;
        return dtEDC;
    }

    private static string Out_ToXML(string[] no, string[] data)
    {

        DataTable dtEDC = new DataTable("Table");
        string row_Name = ConfigurationManager.AppSettings["AMSTxn_Function_Row"];
        string[] rowName;
        DataColumn Col;
        DataRow row;

        Col = new DataColumn();
        Col.ColumnName = "No";
        dtEDC.Columns.Add(Col);
        Col = new DataColumn();
        Col.ColumnName = "Data";
        dtEDC.Columns.Add(Col);
        Col = new DataColumn();
        Col.ColumnName = "WID";
        dtEDC.Columns.Add(Col);

        DataSet DS = new DataSet();
        DS.Tables.Add(dtEDC);
        DS.DataSetName = ConfigurationManager.AppSettings["AMSTxn_Function"].ToString(); //"AutoEDC";

        int iAFR = row_Name.Split(',').Length;
        rowName = new string[iAFR];
        rowName = row_Name.Split(',');

        for (int i = 0; i < data.Length; i++)
        {
            row = dtEDC.NewRow();
            row["No"] = no[i];
            row[rowName[1]] = data[i];
            row[rowName[2]] = "";
            dtEDC.Rows.Add(row);
        }

        XmlDataDocument xdd = new XmlDataDocument(DS);

        return xdd.InnerXml;
    }

    private static string FileHandle(string FullPath)
    {
        try
        {
            //PubUtil.WriteLog("Info", "FileHandle", "Start", "", FullPath);
            // Move not currect day Files into current month folder. Add by Stanley 20130719
            DateTime.Now.ToShortDateString();
            //DateTime dt = DateTime.Now;
            DateTime dt = DateTime.Now.AddDays(int.Parse(ConfigurationManager.AppSettings["FileDay"].ToString()));
            //string sFileMonth = dt.Year + string.Format("{0:00}", fFile.LastWriteTime.Month);
            string sMonth = string.Format("{0:00}", dt.Month);
            // Process Current Folder files.

            foreach (string sFiles in Directory.GetFiles(FullPath))
            {

                //1. Delete *.bak files.
                string sFileName = sFiles.ToUpper();
                string[] sub;
                // sub = new string[3];
                sub = sFileName.Split('.');
                if (sub[sub.Length - 1] == "BAK")
                {
                    File.Delete(sFiles);
                    continue;
                };

                //2 Move not current day of files into current month folder
                FileInfo fFile = new FileInfo(sFiles);
                string sFilename = fFile.Name;
                string Flag = ConfigurationManager.AppSettings["EnableFilter"].ToString();
                //PubUtil.WriteLog("Info", "EnableFilter", "Start", "", Flag);
                if (Flag == "True")
                {
                    //string sFileMonth = dt.Year + string.Format("{0:00}", fFile.LastWriteTime.Month); 
                    string sFileMonth = string.Format("{0:00}", fFile.LastWriteTime.Year) + string.Format("{0:00}", fFile.LastWriteTime.Month); //20140515 Stanley Enhance

                    if (!Directory.Exists(FullPath + sFileMonth))
                    {
                        Directory.CreateDirectory(FullPath + sFileMonth);

                        //PubUtil.WriteLog("Info", "Create Folder", "Done", "", FullPath + dt.Year + sMonth);
                    }

                    if (dt > fFile.LastWriteTime)
                    {
                        fFile.CopyTo((FullPath + sFileMonth + @"\" + sFilename), true);
                        fFile.Delete();
                        // File.MoveTo(FullPath + sFileMonth + @"\" + sFilename);
                        //PubUtil.WriteLog("Info", "File Move", "", "", sFilename);
                    }
                }
                else
                {
                    string[] FilterType = ConfigurationManager.AppSettings["FileType"].ToString().Split(',');
                    string[] subName = sFilename.Split('.');
                    //if (FilterType.Length > 1)
                    //{
                    for (int i = 0; i < FilterType.Length; i++)
                    {
                        string Filter = FilterType[i];
                        if (Filter.ToUpper() == subName[subName.Length - 1].ToUpper())
                        {
                            string sFileMonth = dt.Year + string.Format("{0:00}", fFile.LastWriteTime.Month);

                            if (!Directory.Exists(FullPath + sFileMonth))
                            {
                                Directory.CreateDirectory(FullPath + sFileMonth);
                                //PubUtil.WriteLog("Info", "Create Folder", "Done", "", FullPath + dt.Year + sMonth);
                            }

                            if (dt > fFile.LastWriteTime)
                            {
                                fFile.CopyTo((FullPath + sFileMonth + @"\" + sFilename), true);
                                fFile.Delete();
                                //PubUtil.WriteLog("Info", "File Move", "", "", sFilename);
                            }
                        }
                    }
                }
            }
            //PubUtil.WriteLog("Info", "FileHandle", "Completed", "", "");
            return "";
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "FileHandle", "", "", ex.Message);
            return "";
        }
    }
    //Check Lot Wafer Count with Process Wafer Count, Match: Track Out Lot, Not Match, Hold Lot.
    public static string checkMassOut(string sLotNo, string sRSC, string sRecipe)
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "MassOut Check", sRSC, sLotNo, "Start");
        try
        {
            string sSQL = "";
            string sReturn = "";
            string sXMLdata = "";
            string sLotWaferCount = "0";
            string sProcessWaferCount = "0";
            DataSet ds = new DataSet();

            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton

            if (sLotNo.IndexOf('@') > -1)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Lot include @!!");
                return "";
            }

            ////20170406 add by chuck for SWR
            //if (getIsSWR(sRSC, sLotNo))
            //{
            //    PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "It is SWR..");
            //    return "";
            //} 

            //Get Process Wafer Count 
            string[] aReturnValue = GetProcessWaferCount(sLotNo, sRSC, sRecipe).Split('|');
            if (aReturnValue.Length > 3)
            {
                sLotNo = aReturnValue[0];
                sRSC = aReturnValue[1];
                sRecipe = aReturnValue[2];
                sProcessWaferCount = aReturnValue[3];
                //     0           1                         2           3    4                          5    6    7    8    9    10   11
                //BP3303333|T1_BP_BCOA001|4620-12|23[option(EDC data):|a[0]|a[1]|a[2]|a[3]|a[4]|a[5]|a[6]|a[7]|]
            }

            if (!isExistValitationTable(sLotNo, sRSC))
            {
                //******20161108 add by chuck for sync elton code******
                ////Modify by Stanley 20140717 =====
                //sSQL = "select * from insitedev.amkaas_lot_validate where CONTAINERNAME='" + sLotNo + "' and resourcename like '" + sRSC + "%' and trackinflag <>'3' ";
                //ds=sObjService.AMSDBQuery(sSQL);
                //if (ds.Tables.Count > 0)
                //{
                string HoldCmd = ConfigurationManager.AppSettings["HoldCMD"].ToString();
                sXMLdata = "<Root><TxnType>LotHold</TxnType><Employee>icsadmin</Employee><Container>" + sLotNo + "</Container><HoldReason>SYS HOLD</HoldReason><Comments>" + HoldCmd.Replace("RSC", sRSC) + "</Comments><PortNo></PortNo><ClientIP>10.185.32.190</ClientIP></Root>";
                sReturn = sObjService.AMSTxn(sXMLdata);
                PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Hold Lot!");
                //}    
                //=========================
                return "";
            } //Add by Stanley 20131001

            //Get Lot Wafer Count by Validation Table
            //sSQL = "select SUM(QTYCURRENTWAFERCOUNT) as WaferCount from insitedev.amkaas_lot_validate where CARRIER='" + sLotNo + "' and resourcename like '" + sRSC + "%'";

            //Add by Chuck for avoid get duplicate record on 20150702
            sSQL = "select SUM(QTYCURRENTWAFERCOUNT) as WaferCount from (select distinct (t.containername) containername, t.carrier, t.QTYCURRENTWAFERCOUNT, t.resourcename from AMKAAS_LOT_VALIDATE t) a where a.CARRIER = '" + sLotNo + "' and a.resourcename like '" + sRSC + "%'";

            PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Get Lot Wafer Count Start");
            ds = sObjService.AMSDBQuery(sSQL);

            if (ds != null)
            {
                sLotWaferCount = ds.Tables[0].Rows[0][0].ToString();
                PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Lot Wafer Count is - " + sLotWaferCount);

                //if ( sRSC.IndexOf("BSPT001") > 0 || sRSC.IndexOf("BSPT003") > 0 || sRSC.IndexOf("T5-BSPT") > 0) // 20140516 Stanley modify for T1.
                if (sRSC.IndexOf("T5-BSPT") > 0) // 20140707 Stanley modify for T1.
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "ByPass Wafer Count.");
                    sRecipe = sLotWaferCount;
                }

            }
            else
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Get Lot Wafer Count fail.");
            }

            //Compare Wafer Count, if match call Mass Out, else Hold Lot.
            if (sProcessWaferCount.Equals(sLotWaferCount))
            {
                try
                {
                    string sEDCdata = "";
                    if (sRSC.IndexOf("BMBR") > -1 && aReturnValue.Length > 10)
                    {
                        //     0           1         2     3                    4     5    6    7    8    9    10   11
                        //BP3303333|T1_BP_BCOA001|4620-12|23[option(EDC data):|a[0]|a[1]|a[2]|a[3]|a[4]|a[5]|a[6]|a[7]|]
                        sEDCdata = "<EDC>";
                        for (int i = 4; i < aReturnValue.Length - 1; i++)
                        {
                            sEDCdata += "<Item><WaferID></WaferID><Value>" + aReturnValue[i] + "</Value></Item>";
                        }
                        sEDCdata += "</EDC>";
                    }

                    //Call Move Out Txn
                    PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Call Mass Out - Start");
                    sXMLdata = "<Root><TxnType>MassInOut</TxnType><Container>" + sLotNo + "</Container><Resource>" + sRSC + "</Resource>" + sEDCdata + "<PortNo></PortNo><ProcTxn>2</ProcTxn><ClientIP>10.185.32.190</ClientIP></Root>";

                    sReturn = sObjService.AMSTxn(sXMLdata);
                    //if (sReturn=="")
                    if (sReturn == _AASReturn) //20161108 add by chuck for sync elton code
                    {
                        PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Call Mass Out - Completed. " + sReturn);
                    }
                    else
                    {
                        //20170406 add by chuck for SWR
                        if (getIsSWR(sRSC, sLotNo))
                        {
                            PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "It is SWR..");
                            return "";
                        } 

                        PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Mass out Fail ! ==>" + sReturn);
                        string HoldCmd = ConfigurationManager.AppSettings["HoldCMD"].ToString();
                        sXMLdata = "<Root><TxnType>LotHold</TxnType><Employee>icsadmin</Employee><Container>" + sLotNo + "</Container><HoldReason>SYS HOLD</HoldReason><Comments>" + HoldCmd.Replace("RSC", sRSC) + "</Comments><PortNo></PortNo><ClientIP>10.185.32.190</ClientIP></Root>";
                        sReturn = sObjService.AMSTxn(sXMLdata);
                    }
                }
                catch
                {
                    //Update Process Complete flag, for Rrtry functionality
                    sSQL = "update insitedev.amkaas_lot_validate set uomname='PROCESSCOMPLETED' where CARRIER='" + sLotNo + "' and resourcename like '" + sRSC + "%'";
                    PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Update Process Complete. SQL:" + sSQL);
                    //sSQL = "select UOMNAME from insitedev.amkaas_lot_validate where portno = '" + sLotNo + "' and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
                    ds = sObjService.AMSDBQuery(sSQL);
                }
            }
            else
            {
                //sProcessWaferCount = sRecipe    //assign process wafer count
                // if (sProcessWaferCount.Equals(sLotWaferCount))
                if (sRecipe.Equals(sLotWaferCount))
                {
                    try
                    {
                        //Call Move Out Txn
                        PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Call Mass Out - Start");
                        sXMLdata = "<Root><TxnType>MassInOut</TxnType><Container>" + sLotNo + "</Container><Resource>" + sRSC + "</Resource><PortNo></PortNo><ProcTxn>2</ProcTxn><ClientIP>10.185.110.5</ClientIP></Root>";

                        sReturn = sObjService.AMSTxn(sXMLdata);
                        //if (sReturn == "")
                        if (sReturn == _AASReturn) //20161108 add by chuck for sync elton code
                        {
                            PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Call Mass Out - Completed");
                        }
                        else
                        {
                            //20170406 add by chuck for SWR
                            if (getIsSWR(sRSC, sLotNo))
                            {
                                PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "It is SWR..");
                                return "";
                            } 

                            PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Mass out Fail. " + sReturn);
                            string HoldCmd = ConfigurationManager.AppSettings["HoldCMD"].ToString();
                            sXMLdata = "<Root><TxnType>LotHold</TxnType><Employee>icsadmin</Employee><Container>" + sLotNo + "</Container><HoldReason>SYS HOLD</HoldReason><Comments>" + HoldCmd.Replace("RSC", sRSC) + "</Comments><PortNo></PortNo><ClientIP>10.185.32.190</ClientIP></Root>";
                            sReturn = sObjService.AMSTxn(sXMLdata);
                        }
                    }
                    catch
                    {
                        //Update Process Complete flag, for Rrtry functionality
                        sSQL = "update insitedev.amkaas_lot_validate set uomname='PROCESSCOMPLETED' where CARRIER='" + sLotNo + "' and resourcename like '" + sRSC + "%'";
                        PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Update Process Complete. SQL:" + sSQL);
                        //sSQL = "select UOMNAME from insitedev.amkaas_lot_validate where portno = '" + sLotNo + "' and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
                        sObjService.AMSDBQuery(sSQL);
                    }
                }
                else
                {
                    //Wafer Count Not match, Hold Lot
                    PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Wafer Count Not Match - Hold Lot Start.");
                    String HoldComment = "PE, Wafer Count Not Match! Lot-" + sLotWaferCount + " Eqp-" + sProcessWaferCount + "@" + sRSC + ". Hold by eCIM";		//Hold Comment

                    //"PE1, Wafer Count Not Match! Lot-" + sLotWaferCount + " Eqp-" + sProcessWaferCount + ". Hold by eCIM";		//Hold Comment

                    sXMLdata = "<Root><TxnType>LotHold</TxnType><Employee>icsadmin</Employee><Container>" + sLotNo + "</Container><HoldReason>SYS HOLD</HoldReason><Comments>" + HoldComment + "</Comments><PortNo></PortNo><ClientIP>10.185.32.190</ClientIP></Root>";
                    sReturn = sObjService.AMSTxn(sXMLdata);

                    sSQL = "update insitedev.amkaas_lot_validate set uomname='SYS HOLD' where CARRIER='" + sLotNo + "' and resourcename like '" + sRSC + "%'";
                    PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Wafer Count Not Match - Hold Lot SQL:" + sSQL);
                    sObjService.AMSDBQuery(sSQL);
                    PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "Wafer Count Not Match - Hold Lot End.");
                }
            }

            PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut", sRSC, sLotNo, "MassOut Check Completed.");
            return "";
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "checkMassOut", sRSC, sLotNo, ex.Message);
            throw;
        }
    }

    public static void setWaferProcessCount(string sLotNo, string sRSC)
    {
        try
        {
            //20171024 change new method call to thread
            bool isNeedRun_1 = false;
            bool isNeedRun_2 = false;

            isNeedRun_1 = PubUtil.isExistValitationTable(sLotNo, sRSC);
            if (!isNeedRun_1)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, "No need to setWaferProcessCount");
                return;
            }

            isNeedRun_2 = PubUtil.isSYSHOLD(sLotNo, "", sRSC);
            if (!isNeedRun_2)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, "No need to setWaferProcessCount, because SYS HOLD");
                return;
            }
            //20171024 change new method call to thread

            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton

            //******20161108 add by chuck for sync elton code******
            //20160623 Added by Elton for bypass invlid
            //sSQL = "select UOMNAME from insitedev.amkaas_lot_validate " +
            //        "where CARRIER='" + sLotNo + "' AND CARRIER = CONTAINERNAME";
            //ds = sObjService.AMSDBQuery(sSQL);

            //if (ds != null && ds.Tables.Count == 0) return;
            //******20161108 add by chuck for sync elton code******

            sSQL = "update insitedev.amkaas_lot_validate set uomname=" +
                    "(select NVL(UOMNAME,0) + 1 UOMNAME from insitedev.amkaas_lot_validate " +
                    "where CARRIER='" + sLotNo + "' AND CARRIER = CONTAINERNAME) " +
                    "where CARRIER='" + sLotNo + "'";
            PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, " Update setWaferProcessCount SQL" + sSQL);

            //sSQL = "update insitedev.amkaas_lot_validate set uomname=" + ds.Tables[0].Rows[0][0].ToString() + "where CARRIER='" + sLotNo + "'";

            sObjService.AMSDBQuery(sSQL);

            //PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, "setWaferProcessCount ==>" + ds.Tables[0].Rows[0][0].ToString());
            //Add by Chuck for fix WriteLog exception : can not find table on 20150702
            PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, "setWaferProcessCount - Completed.");
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "setWaferProcessCount", sRSC, sLotNo, ex.Message);
            throw;
        }
    }

    //20161110 add by Elton for New Transaction AmkLotWfrProc
    public static string setAmkLotWfrProc(string sLotNo, string sRSC)
    {
        string sReturn = "Not Completed!!";
        try
        {                                               //Exist Transaction to get BatchGUID Xml Format:
            //sLotNo data sample , BP1234567-10-2       //<Root>
            string[] sa = sLotNo.Split('-');            //  <TxnType>GetBatchGUID</TxnType>
            sLotNo = sa[0];                             //  <Container>BU6180195</Container>
            string sSlotNo = sa[1];                     //  <Resource>T5-BCOA04</Resource>
            string sAction = sa[2];                     //</Root>
            string sSQL = "";                           //check lot id is correct or not, check FUR TOOL, 20170911 add by Jim
            DataSet ValidationTable = new DataSet();    //check lot id is correct or not, check FUR TOOL, 20170911 add by Jim
            DataSet AMSWaferCount = new DataSet();      //check lot id is correct or not, check FUR TOOL, 20170911 add by Jim
            DataSet WaferStartCount = new DataSet();    //check lot id is correct or not, check FUR TOOL, 20170911 add by Jim
            //DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();

            //20171128 add by Elton for reduce AAS TXN                
            //**************
            try
            {
                //get Validation Table Data
                sSQL = "SELECT * FROM GETECIMTOOLDOWN WHERE CONTAINERNAME = '" + sLotNo + "'";
                ValidationTable = sObjService.AMSDBQuery(sSQL);
                if (ValidationTable.Tables.Count > 0)
                {
                    //Do Nothing!
                }
                else
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc", sRSC, "ByPass --- ", sLotNo);
                    return "NoNeed";
                }
            }
            catch (Exception ex)
            {
                PubUtil.WriteLog("Exception", "setAmkLotWfrProc - GETECIMTOOLDOWN", sRSC, sLotNo, ex.Message);
            }
            //**************

            /* Due to Vendor fix EQP. Software, so no need handley by Amkor ECIM      
                        //******************************************
                        try
                        {
                            //check lot id is correct or not, check FUR TOOL, 20170911 add by Jim
                            if (sRSC.Contains("FUR"))
                            {
                                if (sAction.Equals("1"))
                                {
                                    //1.1 remove _3 */
                                    sLotNo = sLotNo.Replace("_3", ""); //Due to Eqp. report slot_3 cause set wafer count fail.
                                    /*Boolean bGetLotID = false;
                                    int point = 0;
                                    string sNewLotID = "";

                                    //get Validation Table Data
                                    sSQL = "select a.carrier from amkaas_lot_validate a where resourcename like '" + sRSC + "%' and a.carrier = a.containername and trackinflag = '3' order by trackindatetime";
                                    ValidationTable = sObjService.AMSDBQuery(sSQL);

                                    while (!bGetLotID && ValidationTable.Tables.Count > 0)
                                    {
                                        sNewLotID = ValidationTable.Tables[0].Rows[point][0].ToString();

                                        //1.2 query lot qty
                                        //get WaferCount from Validation Table(sum)
                                        sSQL = "select SUM(QTYCURRENTWAFERCOUNT) as WaferCount from AMKAAS_LOT_VALIDATE where CARRIER = '" + sNewLotID + "'";
                                        AMSWaferCount = sObjService.AMSDBQuery(sSQL);

                                        //1.3 query wafer start record
                                        // get start count                Add AND RD.Resourcename like sRSC by Jim 20170914
                                        sSQL = "SELECT COUNT(*) AS WaferStart FROM AMKLOTWFRRCD ALWR, RESOURCEDEF  RD, CONTAINER    C, WORKFLOWBASE WFB, WORKFLOWSTEP WFS " +
                                            "WHERE ALWR.RESOURCEID = RD.RESOURCEID AND ALWR.CONTAINERID = C.CONTAINERID AND ALWR.WORKFLOWID = WFB.REVOFRCDID AND ALWR.WORKFLOWSTEPID = WFS.WORKFLOWSTEPID " +
                                            "AND C.CONTAINERNAME = '" + sNewLotID + "'AND RD.Resourcename like '" + sRSC + "%'";

                                        WaferStartCount = sObjService.AMSDBQuery(sSQL);

                                        int iWaferStartCount = Int32.Parse(WaferStartCount.Tables[0].Rows[0][0].ToString());
                                        int iAMSTotalCount = Int32.Parse(AMSWaferCount.Tables[0].Rows[0][0].ToString());

                                        //1.4 check wafer start record count < lot qty
                                        //1.4 else query another lot id and update sLotNo
                                        // wafer start smaller than AMS QTY
                                        if (iWaferStartCount < iAMSTotalCount)
                                        {
                                            if (!sLotNo.Equals(sNewLotID))//edit by Jim 20170914 for change new LotID write log 
                                            {
                                                PubUtil.WriteLog(sRSC, "setAmkLotWfrProc", sRSC, sLotNo, "NewLotID = " + sNewLotID + " Change to new LotID");
                                                sLotNo = sNewLotID;
                                            }
                                            bGetLotID = true;
                                            PubUtil.WriteLog(sRSC, "setAmkLotWfrProc", sRSC, sLotNo, "Wafer Start --- " + sSlotNo + " StartQTY = " + iWaferStartCount + " AMSCount = " + iAMSTotalCount);
                                            //00:11:45	BP7350211.01	/	Wafer Start --- 25
                                        }
                                        else
                                        {
                                            point++;
                                        }

                                        //add by Jim 20170914 for Exception :"There is no row at position 2" list ValidationTable Data
                                        if (ValidationTable.Tables[0].Rows.Count == point) //out of index
                                        {
                                            string msg = "";
                                            for (int i = 0; i < ValidationTable.Tables[0].Rows.Count; i++)
                                            {
                                                msg += " Lot[" + i + "]=" + ValidationTable.Tables[0].Rows[i][0].ToString();
                                            }
                                            PubUtil.WriteLog(sRSC, "setAmkLotWfrProc Exception", sRSC, sLotNo, msg);
                                            bGetLotID = true;
                                        }
                                    }
                                }
                                else
                                {
                                    if (sAction.Equals("2"))
                                    {
                                        PubUtil.WriteLog(sRSC, "setAmkLotWfrProc", sRSC, "Wafer End --- " + sSlotNo, sLotNo);
                                        //00:13:38	Wafer End --- 25	/	BP7350211.01
                                    }
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            PubUtil.WriteLog(sRSC, "setAmkLotWfrProc for Reflow", sRSC, sLotNo, ex.Message);
                        }
                        //*************************************************
             */

            string sService = "<Root><TxnType>GetBatchGUID</TxnType><Container>" + sLotNo + "</Container><Resource>" + sRSC + "</Resource></Root>";

            sReturn = sObjService.AMSTxn(sService);

            //PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc", "GetBatchGUID.", sService, sReturn);

            if (sReturn.IndexOf("<NewDataSet>") > -1)
            {
                // Return Result: <NewDataSet><Query><BATCHGUID>b00b8f49-17bb-46c8-b380-9167e083eef5</BATCHGUID></Query></NewDataSet>
                sReturn = sReturn.Replace("<NewDataSet><Query>", "");
                sReturn = sReturn.Replace("</Query></NewDataSet>", "");

                //Call New Transaction AmkLotWfrProc Xml Format:
                /*<Root>   
                    <TxnType>AmkLotWfrProc</TxnType>
                    <ExecuteAction>1</ExecuteAction>    // 1 Wafer Proecss Start, 2 Wafer Process End, 3 Delete Wafer Process Record, 4 Wafer Process Error
                    <Container>BU6180029</Container>    //set by eCIM
                    <WaferSlotNo>01</WaferSlotNo>       //set by eCIM
                    <BatchGUID>eae4c14a-3d24-46a9-8235-b01bcafc8b6c</BatchGUID> //get from BatchGUID
                    <Employee>90906</Employee>                                  //get from BatchGUID
                    <Resource>T5-BPLT04</Resource>                              //get from BatchGUID
                    <TrackInRsc>T5-BPLT04-A</TrackInRsc>                        //get from BatchGUID
                    <IsMultiLot>1</IsMultiLot>          //1 Batch Lots Process, 0 Single Lot Process    //get from BatchGUID
                </Root>*/

                sService = "<Root><TxnType>AmkLotWfrProc</TxnType><ExecuteAction>" + sAction + "</ExecuteAction><Container>" + sLotNo + "</Container><WaferSlotNo>" + sSlotNo + "</WaferSlotNo>" + sReturn + "</Root>";
                sReturn = sObjService.AMSTxn(sService);

                //PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc", "AmkLotWfrProc.", sService, sReturn);
                //PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc", sRSC, sLotNo, "setAmkLotWfrProc - Completed.");
            }
            else
            {
                PubUtil.WriteLog("Exception", "setAmkLotWfrProc", sRSC, sLotNo, "No Record exist!! - Completed.");
            }
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "setAmkLotWfrProc", sRSC, sLotNo, ex.Message);
            sReturn = ex.Message;
        }

        return sReturn;
    }

    public static void setWaferProcessCountbyPort(string sRSC, string sPort)
    {
        try
        {
            string sSQL = "";
            //string sSQL1 = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton

            //******20161108 add by chuck for sync elton code******
            //sSQL1 = "select NVL(UOMNAME,0) + 1 UOMNAME from insitedev.amkaas_lot_validate " +
            //    "where portno=" + sPort + " and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
            //ds = sObjService.AMSDBQuery(sSQL1);

            //if (ds.Tables.Count == 0) return;
            //******20161108 add by chuck for sync elton code******

            sSQL = "update insitedev.amkaas_lot_validate set uomname=" +
                    "(select NVL(UOMNAME,0) + 1 UOMNAME from insitedev.amkaas_lot_validate " +
                    "where portno=" + sPort + " and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME) " +
                    "where portno=" + sPort + " and resourcename like '" + sRSC + "%' ";

            PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCountbyPort", sRSC, sPort, "setWaferProcessCountbyPort. SQL:" + sSQL);

            //sSQL = "update insitedev.amkaas_lot_validate set uomname=" +
            //        "(select NVL(UOMNAME,0) + 1 UOMNAME from insitedev.amkaas_lot_validate " +
            //        "where portno=" + sPort + " AND CARRIER = CONTAINERNAME) " +
            //        "where portno=" + sPort ;

            sObjService.AMSDBQuery(sSQL);
            //PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCountbyPort", sRSC, sPort, "setWaferProcessCountbyPort ==>" + ds.Tables[0].Rows[0][0].ToString());
            //Add by Chuck for fix WriteLog exception : can not find table on 20150702
            PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCountbyPort", sRSC, sPort, "setWaferProcessCountbyPort - Completed.");
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "setWaferProcessCountbyPort", sRSC, sPort, ex.Message);
            throw;
        }
    }

    //private static Boolean isExistValitationTable(string LotNo, string sRSC) //Add by Stanley 20131001
    public static Boolean isExistValitationTable(string LotNo, string sRSC) //Modify by Chuck 20150702
    {
        try
        {
            string sSQL = "";
            DataSet ds = new DataSet();

            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton
            sSQL = "select * from insitedev.amkaas_lot_validate where CONTAINERNAME='" + LotNo + "' and resourcename like '" + sRSC + "%' and trackinflag ='3' ";
            PubUtil.WriteLog("eCIMWebServiceLog", "isExistValitationTable", sRSC, LotNo, "Check Valitation Table SQL:" + sSQL);
            ds = sObjService.AMSDBQuery(sSQL);

            if (ds.Tables.Count > 0)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "isExistValitationTable", sRSC, LotNo, "Check Valitation Table Done.");
                return true;
            }
            else
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "isExistValitationTable", sRSC, LotNo, "Check Valitation Table is Null!.");
                return false;
            }
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "isExistValitationTable", sRSC, LotNo, ex.Message);
            return false;
        }
    }

    public static string callMassIn(string sLotNo, string sRSC, string sPort)
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "callMassIn", sRSC, sLotNo, "Start");
        try
        {
            string sReturn = "0";
            string sXMLdata = "";

            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton
            try
            {
                //Call Move In Txn
                sXMLdata = "<Root><TxnType>MassInOut</TxnType><Container>" + sLotNo + "</Container><Resource>" + sRSC + "</Resource><PortNo>" + sPort + "</PortNo><ProcTxn>1</ProcTxn><ClientIP>10.185.110.5</ClientIP></Root>";
                sReturn = sObjService.AMSTxn(sXMLdata);
                //if (sReturn == "")
                if (sReturn == _AASReturn) //20161108 add by chuck for sync elton code
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "callMassIn", sRSC, sLotNo, "Call Mass In - Completed. " + sReturn);
                }
                else
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "callMassIn", sRSC, sLotNo, "Mass in Fail ! ==>" + sReturn);
                    string HoldCmd = ConfigurationManager.AppSettings["HoldCMD"].ToString();
                    sXMLdata = "<Root><TxnType>LotHold</TxnType><Employee>icsadmin</Employee><Container>" + sLotNo + "</Container><HoldReason>SYS HOLD</HoldReason><Comments>" + HoldCmd.Replace("RSC", sRSC) + "</Comments><PortNo></PortNo><ClientIP>10.185.32.190</ClientIP></Root>";
                    sReturn = sObjService.AMSTxn(sXMLdata);
                }
            }
            catch
            {
            }

            PubUtil.WriteLog("eCIMWebServiceLog", "callMassIn", sRSC, sLotNo, "End.");
            return "1";
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "callMassIn", sRSC, sLotNo, ex.Message);
            throw;
        }
    }
    //public static string GetAlignerEdcData(string sType)
    public static int Get_SPT_WaferFile(string LotNo, string sRSC, DataRow Info)
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "Get_SPT_WaferFile", sRSC, LotNo, "Get_SPT_WaferFile Start.");
        try
        {
            int iFileCount = 0;
            int iHour = 2;
            string FullPath = @"\\" + Info[1].ToString() + @"\" + Info[2].ToString();
            PubUtil.setLogonDomain(FullPath, Info[3].ToString(), Info[4].ToString());

            foreach (string sFiles in Directory.GetFiles(FullPath))
            {
                FileInfo fFile = new FileInfo(sFiles);
                if (fFile.Name.IndexOf(LotNo) > -1)  //Filter by Lot by Stanley 20140701
                {

                    if (fFile.LastWriteTime.AddHours(iHour) > DateTime.Now)
                    {
                        iFileCount += 1;
                        PubUtil.WriteLog("eCIMWebServiceLog", "Get_SPT_WaferFile", sRSC, LotNo, iFileCount.ToString() + fFile.LastWriteTime.ToString());
                    }
                }
            }
            PubUtil.WriteLog("eCIMWebServiceLog", "Get_SPT_WaferFile", sRSC, LotNo, iFileCount.ToString());
            return iFileCount;
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "Get_SPT_WaferFile", sRSC, LotNo, ex.Message);
            throw;
        }
    }

    private static string T1_Step_Parser(string FullPath, string sLotID, string sRSC, string sRecipe) //Add by Stanley 20140714
    {
        //bool bFindFile = false; //20161108 add by chuck for sync elton code
        string FileName = "";
        StreamReader fnr;
        string[] sDataContent;
        int iIdx = 0;
        DirectoryInfo dir;
        DateTime dt = DateTime.Now.AddMinutes(-30);

        // FullPath=@"D:\log\"; //For test

        try
        {
            //FileHandle(FullPath); //Append by Stanley 20130719 for file handle.
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "T1_Step_Parser", sRSC, sLotID, ex.Message);
        }

        try
        {
            FileName = "CJ_" + sLotID + "*.txt";
            FullPath += DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString("0#") + "_" + DateTime.Now.Day.ToString("0#") + @"\";// +FileName;
            dir = new DirectoryInfo(FullPath);

            FileInfo[] aFileInfo = dir.GetFiles(FileName);
            if (aFileInfo.Length == 0)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "T1_Step_Parser", sRSC, sLotID, "Get 0 file list!" + FullPath + "/" + FileName);
                return "Get 0 file list - " + FullPath;
            }
            if (aFileInfo.Length == 1)
            {
                FullPath += aFileInfo[0].Name;
                FileName = aFileInfo[0].Name;
            }
            else
            {
                string sFileName = "";
                for (int i = 0; i < aFileInfo.Length; i++)
                {
                    if (dt < aFileInfo[i].LastWriteTime)
                    {
                        dt = aFileInfo[i].LastWriteTime;
                        sFileName = aFileInfo[i].Name;
                    }
                }
                FileName = sFileName;
                FullPath += sFileName;
            }

            FullPath = FullPath.Replace("\\", @"\");

            fnr = new StreamReader(FullPath);
            sDataContent = new string[1500];

            while (fnr.Peek() >= 0)
            {
                string stmp = fnr.ReadLine().Replace("\n", "");
                stmp = stmp.Replace("\r\n", "");
                if (stmp != "")
                {
                    sDataContent[iIdx] = stmp;
                    iIdx += 1;
                }
            }
            fnr.Close();

            Array.Resize(ref sDataContent, iIdx);
            Boolean fProcess = false;
            Boolean fRecipe = false;
            int iStartIdx = 0;

            for (int i = 1; i < sDataContent.Length; i++)
            {
                if (sDataContent[i] != "")
                {
                    if (sDataContent[i].IndexOf("Recipe ID:") > -1)
                    {
                        if (sDataContent[i].IndexOf(sRecipe) > -1) fRecipe = true;
                    }
                }
                if (sDataContent[i] != "") if (sDataContent[i].IndexOf("Processed") > -1) fProcess = true;
            }
            if (fRecipe && fProcess) iStartIdx = 1;
            return iStartIdx.ToString();
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "T1_Step_Parser", sRSC, sLotID, ex.Message);
            throw ex;
        }
    }
    //==========mark add for query Lot Info by Machine ID for T3 20410815===========
    public static string FindLotInfobyMachineID(string sRSC)
    {
        string strMsg = "";
        try
        {
            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton

            sSQL = "SELECT /*+RULE*/ C.CONTAINERNAME,MH.TXNDATE,OW.OWNERNAME,RD.RESOURCENAME FROM MOVEHISTORY MH,CONTAINER C,RESOURCEDEF RD,OWNER OW " +
                    "WHERE MH.TXNDATE>SYSDATE-1 AND MH.HISTORYID=C.CONTAINERID AND C.OWNERID=OW.OWNERID " +
                    "AND MH.RESOURCEID=RD.RESOURCEID AND RD.RESOURCENAME = '" + sRSC + "' " +
                    "ORDER BY MH.TXNDATE DESC";
            ds = sObjService.AMSDBQuery(sSQL);
            if (null != ds && ds.Tables[0].Rows.Count > 0)
            {
                if (!ds.Tables[0].Rows[0]["OWNERNAME"].ToString().Trim().Equals("") && ds.Tables[0].Rows[0]["OWNERNAME"].ToString().ToUpper().Trim().Equals("PRODUCTION"))
                {
                    strMsg += ds.Tables[0].Rows[0]["CONTAINERNAME"].ToString().Trim() + "|";
                    strMsg += ds.Tables[0].Rows[0]["TXNDATE"].ToString().Trim() + "|";
                    strMsg += ds.Tables[0].Rows[0]["OWNERNAME"].ToString().Trim();
                }
            }
            return strMsg;
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "FindLotInfobyMachineID", sRSC, "", ex.Message);
            return strMsg;
        }
    }
    //========================================END=====================================
    public static string CopyFile(string sLotNo, string sRSC, string sRecipe)
    {
        //\\t1review02\edcdata\T1_BP_BMBR001
        PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "Start");
        try
        {
            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sSQL = "Select * from att_ecim.ecimservice_eqpinfo where name= '" + sRSC + "' and TYPE =' '";
            ds = sObjService.AMSDBQuery(sSQL);   //0:EQPName 1:IP 2:Folder 3:LoginName 4:LoginPWD 5: File

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string FullPath = @"\\" + ds.Tables[0].Rows[i][1].ToString() + @"\" + ds.Tables[0].Rows[i][2].ToString();
                string DestinationPath = ConfigurationManager.AppSettings["AutoCopyFilePath_BMBX"].ToString() + sRSC + @"\" + ds.Tables[0].Rows[i][2].ToString().Substring(0, ds.Tables[0].Rows[i][2].ToString().Length - 1);


                switch (ds.Tables[0].Rows[i][2].ToString().Substring(0, ds.Tables[0].Rows[i][2].ToString().Length - 1))
                {
                    case "Results":
                        FullPath = @"\\" + ds.Tables[0].Rows[i][1].ToString() + @"\" + ds.Tables[0].Rows[i][2].ToString() + sRecipe + @"\" + sLotNo + @"\";
                        DestinationPath = ConfigurationManager.AppSettings["AutoCopyFilePath_BMBX"].ToString() + sRSC + @"\" + ds.Tables[0].Rows[i][2].ToString().Substring(0, ds.Tables[0].Rows[i][2].ToString().Length - 1) + @"\" + sLotNo;
                        break;
                }

                //Check Eqp. folder
                if (!Directory.Exists(FullPath))
                {
                    PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "Folder not exists in Eqp. FullPath : " + FullPath);
                    return "";
                }

                //prepare local folder
                PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "prepare local folder");
                //String sTempFolder = @"D:\\temp\";
                String sTempFolder = @"D:\\temp\\" + sRSC + @"\";
                if (!Directory.Exists(sTempFolder))
                    Directory.CreateDirectory(sTempFolder);

                //Copy file from EQP. to Local
                PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "Copy file from EQP. ( " + FullPath + "sRecipe" + "-" + sLotNo + ".L ) to Local");
                //For fix share folder connection issue. add by chuck on 20161202
                string[] sourcePath = Directory.GetFiles(FullPath);
                foreach (string sFiles in sourcePath)
                //foreach (string sFiles in Directory.GetFiles(FullPath))
                {
                    FileInfo fFile = new FileInfo(sFiles);
                    string sFilename = fFile.Name.ToUpper();

                    ///******Add by chuck on 20150203******
                    if (ds.Tables[0].Rows[i][2].ToString().Substring(0, ds.Tables[0].Rows[i][2].ToString().Length - 1).Equals("Results"))
                    {
                        //PubUtil.WriteLog("Copy File", "Copy File", sTempFolder, sFilename, "Copy all file from EQP. to (" + sTempFolder + ") Start");
                        fFile.CopyTo((sTempFolder + sFilename), true);
                        //PubUtil.WriteLog("Copy File", "Copy File", sTempFolder, sFilename, "Copy all file from EQP. to (" + sTempFolder + ") End");
                    }
                    //******Add by chuck on 20150203******
                    else if (sFilename.Contains(sLotNo))
                    {
                        //PubUtil.WriteLog("Copy File", "Copy File", sTempFolder, sFilename, "Copy file from EQP. to (" + sTempFolder + ") Start");
                        fFile.CopyTo((sTempFolder + sFilename), true);
                        //PubUtil.WriteLog("Copy File", "Copy File", sTempFolder, sFilename, "Copy file from EQP. to (" + sTempFolder + ") End");
                    }
                }


                string Domain = ConfigurationManager.AppSettings["Domain"].ToString();
                string Account = ConfigurationManager.AppSettings["User"].ToString();
                string Password = ConfigurationManager.AppSettings["Pwd"].ToString();
                //PubUtil.WriteLog("Copy File", "Copy File", Domain, Account, Password); 
                //PubUtil.setLogonDomain("tw", "bumpfab", "Fabbump9");//
                //PubUtil.setLogonDomain("tw", "admt5", "T5@16847968");
                PubUtil.setLogonDomain(Domain, Account, Password);
                if (!Directory.Exists(DestinationPath))
                    Directory.CreateDirectory(DestinationPath);

                //Process DestinationPath data
                try
                {
                    FileHandle(DestinationPath + @"\");
                }
                catch (Exception ex)
                {
                    PubUtil.WriteLog("Exception", "File Backup Function has exception!", sRSC, sLotNo, ex.Message);
                }

                //Copy file from Local to \\t1review02\edcdata\sRSC
                PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "Copy file from Local to " + DestinationPath);
                //For fix share folder connection issue. add by chuck on 20161202
                string[] targetPath = Directory.GetFiles(sTempFolder);
                foreach (string sFiles in targetPath)
                //foreach (string sFiles in Directory.GetFiles(sTempFolder))
                {
                    FileInfo fFile = new FileInfo(sFiles);
                    string sFilename = fFile.Name;

                    //PubUtil.WriteLog("Copy File", "Copy File", sRSC, sFilename, "Copy Start");
                    fFile.CopyTo((DestinationPath + @"\" + sFilename), true);
                    fFile.Delete();
                    //PubUtil.WriteLog("Copy File", "Copy File", sRSC, sFilename, "Copy End");

                }

                //******Add by chuck for next record can log in success on 20150203******
                PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "第" + (i + 1) + "次" + "Completed");
                PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "Set Domain form TW to Eqp.");
                PubUtil.setLogonDomain(FullPath, ds.Tables[0].Rows[0][3].ToString(), ds.Tables[0].Rows[0][4].ToString());
                ///******Add by chuck for next record can log in success on 20150203******
            }
            return "";
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Error", "Exception", "Copy File", sRSC + " - " + sLotNo, ex.Message); //Modify for add Function and sRC, sLotNo by churk
            return "";
        }

    }

    //Chuck add for fix SQL issue on 20150831
    public static Boolean isSYSHOLD(string sLotNo, string sPort, string sRSC)
    {
        try
        {
            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();

            //For by port use
            if (sLotNo.Equals(""))
            {
                sSQL = "select UOMNAME,CARRIER from insitedev.amkaas_lot_validate where portno = '" + sPort + "' and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
            }
            //For by lot use
            else
            {
                sSQL = "select UOMNAME from insitedev.amkaas_lot_validate where CARRIER = '" + sLotNo + "' and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
            }

            PubUtil.WriteLog("eCIMWebServiceLog", "isSYSHOLD", sRSC, sLotNo + sPort, "Check Valitation Table SQL:" + sSQL);
            ds = sObjService.AMSDBQuery(sSQL);

            if (ds != null && ds.Tables.Count > 0)
            {
                String sUOMNAME = ds.Tables[0].Rows[0][0].ToString();
                if (sUOMNAME.Equals("SYS HOLD"))
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "isSYSHOLD", sRSC, sLotNo + sPort, "Check Valitation Table Done. sUOMNAME : " + sUOMNAME);
                    return false;
                }
                else
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "isSYSHOLD", sRSC, sLotNo + sPort, "Check Valitation Table Done. sUOMNAME : " + sUOMNAME);
                    return true;
                }
            }
            else
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "isSYSHOLD", sRSC, sLotNo + sPort, "Check Valitation Table is Null!.");
                return false;
            }
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "isSYSHOLD", sRSC, sLotNo + sPort, ex.Message);
            return false;
        }
    }

    //20161108 add by chuck for sync elton code
    //Test Only--
    public static System.Data.DataTable GetEDCData(int iServiceType, string sPath, string sLotNo, int iReadings, bool bCompare, string sOffSetArith, double dOffSetValue, int iOffSetAfterPoint, string recipeCode)
    {
        try
        {
            PubUtil.setLogonDomain();

            System.Data.DataTable dtEDC = new System.Data.DataTable();
            clsEDCDataParms oEDCDataParms = new clsEDCDataParms();

            oEDCDataParms.MainFolder = ConfigurationManager.AppSettings["AutoEDCPath"];

            oEDCDataParms.Backup = "Backup";
            oEDCDataParms.Path = sPath;
            oEDCDataParms.BPath = sPath + @"\Before";
            oEDCDataParms.APath = sPath + @"\After";
            oEDCDataParms.LotNo = sLotNo;
            oEDCDataParms.Readings = iReadings;
            oEDCDataParms.Compare = bCompare;
            oEDCDataParms.OffSetArith = sOffSetArith;
            oEDCDataParms.OffSetValue = dOffSetValue;
            oEDCDataParms.OffSetAfterPoint = iOffSetAfterPoint;

            dtEDC.TableName = "Table";
            dtEDC.Columns.Add(new DataColumn("No"));
            dtEDC.Columns.Add(new DataColumn("Data"));
            dtEDC.Columns.Add(new DataColumn("WID"));

            if (!bCompare)
            {
                GetDataFromFile(ref oEDCDataParms, ref dtEDC, iServiceType, recipeCode);
            }
            else
            {
                System.Data.DataTable dsBefore = new System.Data.DataTable();
                System.Data.DataTable dsAfter = new System.Data.DataTable();

                dsBefore = dtEDC.Clone();
                oEDCDataParms.Path = oEDCDataParms.BPath;
                GetDataFromFile(ref oEDCDataParms, ref dsBefore, iServiceType, recipeCode);

                dsAfter = dtEDC.Clone();
                oEDCDataParms.Path = oEDCDataParms.APath;
                GetDataFromFile(ref oEDCDataParms, ref dsAfter, iServiceType, recipeCode);

                if (dsBefore != null & dsAfter != null)
                {
                    DataRow dr;
                    for (int i = 0; i < dsBefore.Rows.Count; i++)
                    {
                        double d1 = Convert.ToDouble(dsBefore.Rows[i][1]) - Convert.ToDouble(dsAfter.Rows[i][1]);

                        //add by Elton for SFC , After - Before , 20120723
                        if (iServiceType == 9) d1 = Convert.ToDouble(dsAfter.Rows[i][1]) - Convert.ToDouble(dsBefore.Rows[i][1]);

                        dr = dtEDC.NewRow();
                        dr["No"] = i;
                        dr["Data"] = d1 < 0 ? 0 : d1;
                        dr["WID"] = dsAfter.Rows[i][2];
                        if (iServiceType != 9)
                            dr["WID"] = "1";
                        dtEDC.Rows.Add(dr);
                    }
                    dtEDC.AcceptChanges();
                }
            }

            string sValue = "";
            foreach (DataRow dr in dtEDC.Rows)
            {
                sValue += dr[1].ToString().Trim() + ",";
            }

            return dtEDC;
        }
        catch (Exception)
        {
            throw;
        }
    }

    private static void GetDataFromFile(ref clsEDCDataParms oEDCDataParms, ref System.Data.DataTable GetLotEDCData, int iServiceType, string AMSrecipeCode)
    {
        try
        {
            bool bFindFile = false;
            string FileName = "";
            string backupFileName = "";
            string FullPath = "";
            string FileType = "";
            string BackPath = "";

            switch (iServiceType)
            {
                case 1: //THK
                    FileType = "MAP";
                    FileName = oEDCDataParms.LotNo + "." + FileType;
                    //oEDCDataParms.MainFolder = @"\\review0\LotReport\RVSI_FQC_BH\AMS_test";
                    break;

                case 2: //RSM
                    FileType = "RSM";
                    FileName = oEDCDataParms.LotNo + "." + FileType;
                    //oEDCDataParms.MainFolder = @"\\review0\LotReport\RVSI_FQC_BH\AMS_test";
                    break;

                case 3: //SHER
                    FileType = "XLS";
                    FileName = oEDCDataParms.LotNo + "." + FileType;

                    switch (ConfigurationManager.AppSettings["FactoryCode"])
                    {
                        case "T1":
                            oEDCDataParms.MainFolder = @"\\t1review02\edcdata";
                            break;

                        case "T5":
                            oEDCDataParms.MainFolder = @"\\Review0\LotReport";
                            break;
                    }
                    break;

                case 4: //PULL
                    FileType = "XLS";
                    FileName = oEDCDataParms.LotNo + "." + FileType;

                    switch (ConfigurationManager.AppSettings["FactoryCode"])
                    {
                        case "T1":
                            oEDCDataParms.MainFolder = @"\\t1review02\edcdata";
                            break;

                        case "T5":
                            oEDCDataParms.MainFolder = @"\\Review0\LotReport";
                            break;
                    }
                    break;

                case 5: //XRF
                    FileType = "txt";
                    FileName = oEDCDataParms.LotNo + "." + FileType;
                    //oEDCDataParms.MainFolder = @"\\review0\other";
                    break;

                case 6: //RS
                    FileType = "txt";
                    FileName = oEDCDataParms.LotNo + "." + FileType;
                    //oEDCDataParms.MainFolder = @"\\review0\other";

                    switch (ConfigurationManager.AppSettings["FactoryCode"])
                    {
                        //case "T1":
                        //    oEDCDataParms.MainFolder = @"\\t1review02\edcdata";
                        //    break;

                        case "T5":
                            oEDCDataParms.MainFolder = @"\\Review0\LotReport";
                            break;
                    }
                    break;

                case 7: //THK(TXT)
                    FileType = "txt";
                    FileName = oEDCDataParms.LotNo + "." + FileType;
                    //oEDCDataParms.MainFolder = @"\\review0\other";
                    break;

                case 8: //XRF-a
                    FileType = "txt";
                    FileName = oEDCDataParms.LotNo + "." + FileType;
                    //oEDCDataParms.MainFolder = @"\\review0\other";
                    break;

                case 9: //SFC 
                    FileType = "txt";
                    //FileName = oEDCDataParms.LotNo + "." + FileType;
                    switch (ConfigurationManager.AppSettings["FactoryCode"])
                    {
                        case "T1":
                            oEDCDataParms.MainFolder = @"\\t1review02\edcdata";
                            break;

                        //case "T5":
                        //    oEDCDataParms.MainFolder = @"\\Review0\LotReport";
                        //    break;
                    }
                    //Keep latest file under folder , others move to backup
                    string sPath = oEDCDataParms.MainFolder + "\\" + oEDCDataParms.Path + "\\";

                    //For T1 BETC special handle Start                        
                    string sRSC = oEDCDataParms.LotNo;
                    string sBeforeFolder = "";
                    switch (sRSC)
                    {
                        case "T1_BP_BETC001":
                        case "T1_BP_BETC007":
                            sBeforeFolder = oEDCDataParms.MainFolder + @"\Etch PD monitor\4CPR\Before\";
                            break;

                        case "T1_BP_BETC002":
                        case "T1_BP_BETC006":
                        case "T1_BP_BETC008":
                        case "T1_BP_BETC002_D": //20121212 add by Elton
                        case "T1_BP_BETC006_D": //20121212 add by Elton
                        case "T1_BP_BETC008_D": //20121212 add by Elton
                            sBeforeFolder = oEDCDataParms.MainFolder + @"\Etch PD monitor\4CEtch\Before\";
                            break;

                        case "T1_BP_BETC003":
                        case "T1_BP_BETC005":
                        case "T1_BP_BETC003_A": //20121212 add by Elton
                        case "T1_BP_BETC005_A": //20121212 add by Elton
                        case "T1_BP_BSCB001":
                        case "T1_BP_BSCB002":
                            sBeforeFolder = oEDCDataParms.MainFolder + @"\Etch PD monitor\2C\Before\";
                            break;

                        default:
                            break;
                    }

                    //Copy B folder to A folder Before
                    switch (sRSC)   //20121123-01 Move by Elton.
                    {
                        case "T1_BP_BETC001":
                        case "T1_BP_BETC007":
                        case "T1_BP_BETC002":
                        case "T1_BP_BETC006":
                        case "T1_BP_BETC008":
                        case "T1_BP_BETC002_D": //20121212 add by Elton
                        case "T1_BP_BETC006_D": //20121212 add by Elton
                        case "T1_BP_BETC008_D": //20121212 add by Elton
                        case "T1_BP_BETC003":
                        case "T1_BP_BETC005":
                        case "T1_BP_BETC003_A": //20121212 add by Elton
                        case "T1_BP_BETC005_A": //20121212 add by Elton
                        case "T1_BP_BSCB001":
                        case "T1_BP_BSCB002":
                            if (sPath.ToUpper().Contains("BEFORE"))//Back up function add by Elton 20121115-1
                            {
                                DirectoryInfo dir = new DirectoryInfo(sBeforeFolder);
                                FileInfo[] aFileInfo = dir.GetFiles("*_0.txt");

                                foreach (FileInfo filename in aFileInfo)
                                {
                                    File.Copy(filename.FullName, sPath + filename.Name, true);//20121121-03 Move by Elton.
                                }
                            }
                            else if (sPath.ToUpper().Contains("AFTER"))//Back up function add by Elton 20121115-1
                            {
                                DirectoryInfo dir = new DirectoryInfo(sPath);
                                FileInfo[] aFileInfo = dir.GetFiles("*_0.txt");

                                foreach (FileInfo filename in aFileInfo)
                                {
                                    File.Copy(filename.FullName, sBeforeFolder + filename.Name, true);
                                }
                            }
                            break;

                        default:
                            break;
                    }

                    //For T1 BETC special handle End
                    string[] aFile = Directory.GetFiles(sPath, "*_0.txt");          //20121121-02 Move by Elton.
                    FileName = aFile[0].Substring(aFile[0].LastIndexOf("\\") + 1);  //20121121-01 Move by Elton.

                    if (aFile.Length > 1)
                    {
                        for (int i = 1; i < aFile.Length; i++)
                        {
                            if (String.Compare(aFile[i], aFile[i - 1]) > 0)
                            {
                                try
                                {
                                    File.Move(aFile[i - 1], aFile[i - 1].Replace(sPath, sPath + "backup\\"));
                                }
                                catch (Exception) { }

                                FileName = aFile[i].Substring(aFile[i].LastIndexOf("\\") + 1);
                            }
                            else if (String.Compare(aFile[i], aFile[i - 1]) < 0)
                            {
                                try
                                {
                                    File.Move(aFile[i], aFile[i].Replace(sPath, sPath + "backup\\"));
                                }
                                catch (Exception) { }

                                FileName = aFile[i - 1].Substring(aFile[i - 1].LastIndexOf("\\" + 1));
                            }
                        }
                    }
                    break;

                case 10:   //Leakage Create by Sakai @20130131
                    FileType = "csv";
                    FileName = oEDCDataParms.LotNo + "." + FileType;
                    break;

                case 12:
                case 13:
                    string sType = "";
                    if (iServiceType == 12) sType = "365";
                    if (iServiceType == 13) sType = "405";
                    sRSC = oEDCDataParms.LotNo;
                    //Com.Trg.EData.eCIMServices.eCIMService eCIMWS = new Com.Trg.EData.eCIMServices.eCIMService();
                    //switch (ConfigurationManager.AppSettings["FactoryCode"])
                    //{
                    //    case "T1":
                    //        eCIMWS.Url = "http://10.185.110.3/eCIMServices/eCIMService.asmx";
                    //        break;


                    //    case "T5":
                    //        eCIMWS.Url = "http://10.185.126.4/eCIMServices/eCIMService.asmx";
                    //        break;
                    //}

                    //GetLotEDCData = eCIMWS.GetAlignerEdcData(sRSC, sType);
                    return;

                case 14: //OMI Create by Stanley @20130624
                case 15: //OMI Create by Stanley @20130624
                case 16: //OMI Create by Stanley @20130624
                    FileType = "txt";
                    FileName = oEDCDataParms.LotNo + "." + FileType;
                    //oEDCDataParms.MainFolder = @"D:\EDC_Temp\";
                    break;

                case 17: //
                    switch (ConfigurationManager.AppSettings["FactoryCode"])
                    {
                        case "T1":
                            oEDCDataParms.MainFolder = @"\\t1review01\t1bumping\LotReport";
                            break;

                        case "T5":
                            oEDCDataParms.MainFolder = @"\\Review0\LotReport";
                            //oEDCDataParms.MainFolder = @"D:\EDC_Temp\"; //test only
                            break;
                    }

                    DirectoryInfo dir1 = new DirectoryInfo(oEDCDataParms.MainFolder);

                    FileInfo[] aFileInfo1 = dir1.GetFiles("*" + System.DateTime.Now.ToString("MMdd") + ".txt");

                    DateTime tempDT = DateTime.Now.AddDays(-1);
                    foreach (FileInfo filename in aFileInfo1)
                    {

                        if (DateTime.Compare(filename.LastWriteTime, tempDT) == 1)
                        {
                            tempDT = filename.LastWriteTime;
                            FileName = filename.Name;
                        }
                    }
                    //FileName = oEDCDataParms.LotNo + "*.txt" ; //test only
                    break;

                case 19:      //20141128 add for Reflow Max Temp.
                    switch (ConfigurationManager.AppSettings["FactoryCode"])
                    {
                        case "T1":
                            oEDCDataParms.MainFolder = @"\\t1review02\edcdata";
                            break;

                        case "T5":
                            oEDCDataParms.MainFolder = @"\\Review0\LotReport";
                            //oEDCDataParms.MainFolder = @"D:\EDC_Temp\"; //test only
                            break;
                    }
                    FileType = "csv";
                    FileName = oEDCDataParms.LotNo + "." + FileType;
                    break;

                default:
                    break;
            }

            backupFileName = oEDCDataParms.LotNo + "_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + "." + FileType;

            if (!Directory.Exists(oEDCDataParms.MainFolder))
            {
                throw new Exception("Main Folder not existing on file server!");
            }

            //if (!Directory.Exists((oEDCDataParms.MainFolder + "\\" + sPath).Replace("\\", @"\")))
            //{
            //    throw new Exception("Sub Folder not existing on file server!");
            //}

            BackPath = oEDCDataParms.MainFolder + "\\" + oEDCDataParms.Path + "\\" + oEDCDataParms.Backup;
            FullPath = oEDCDataParms.MainFolder + "\\" + oEDCDataParms.Path + "\\" + FileName;
            FullPath = FullPath.Replace("\\", @"\");

            if (File.Exists(FullPath))
            {
                bFindFile = true;
            }

            if (!bFindFile)
            {
                throw new Exception(oEDCDataParms.LotNo + "." + FileType + "rsm file not found!");
            }

            if (!Directory.Exists(BackPath))
            {
                Directory.CreateDirectory(BackPath);
            }
            StreamReader fnr;
            string[] sDataContent;
            int iIdx = 0;

            switch (iServiceType)
            {
                case 1:
                    fnr = new StreamReader(FullPath);
                    sDataContent = new string[62];

                    while (fnr.Peek() >= 0)
                    {
                        string stmp = fnr.ReadLine().Replace("\n", "");
                        stmp = stmp.Replace("\r\n", "");
                        if (stmp != "")
                        {
                            sDataContent[iIdx] = stmp;
                            iIdx += 1;
                        }
                    }
                    fnr.Close();
                    //MAPParser(sDataContent, ref GetLotEDCData, oEDCDataParms.Readings);
                    break;

                case 2:
                    fnr = new StreamReader(FullPath);
                    sDataContent = new string[65];

                    while (fnr.Peek() >= 0)
                    {
                        string stmp = fnr.ReadLine().Replace("\n", "");
                        stmp = stmp.Replace("\r\n", "");
                        if (stmp != "")
                        {
                            sDataContent[iIdx] = stmp;
                            iIdx += 1;
                        }
                    }
                    fnr.Close();
                    //RSMParser(sDataContent, ref GetLotEDCData, oEDCDataParms.Readings);
                    break;

                case 3:
                case 4:
                    //SHERParser(iServiceType, FullPath, ref GetLotEDCData, AMSrecipeCode);
                    break;

                case 5:
                    fnr = new StreamReader(FullPath);
                    sDataContent = new string[62];

                    while (fnr.Peek() >= 0)
                    {
                        string stmp = fnr.ReadLine().Replace("\n", "");
                        stmp = stmp.Replace("\r\n", "");
                        if (stmp != "")
                        {
                            sDataContent[iIdx] = stmp;
                            iIdx += 1;
                        }
                    }
                    fnr.Close();
                    //XRFParser(sDataContent, ref GetLotEDCData);
                    break;

                case 6: //RS
                    fnr = new StreamReader(FullPath);
                    sDataContent = new string[5];

                    while (fnr.Peek() >= 0)
                    {
                        string stmp = fnr.ReadLine().Replace("\n", "");
                        stmp = stmp.Replace("\r\n", "");
                        if (stmp != "")
                        {
                            sDataContent[iIdx] = stmp;
                            iIdx += 1;
                        }
                    }
                    fnr.Close();
                    //RSParser(sDataContent, ref GetLotEDCData);
                    break;

                case 7: //THK(txt)
                    fnr = new StreamReader(FullPath);
                    sDataContent = new string[20];

                    while (fnr.Peek() >= 0)
                    {
                        string stmp = fnr.ReadLine().Replace("\n", "");
                        stmp = stmp.Replace("\r\n", "");
                        if (stmp != "")
                        {
                            sDataContent[iIdx] = stmp;
                            iIdx += 1;
                        }
                    }
                    fnr.Close();
                    //THK_TXTParser(sDataContent, ref GetLotEDCData);
                    break;

                case 9: //SFC(rpd)
                    fnr = new StreamReader(FullPath);
                    sDataContent = new string[25];
                    //int i = 0;
                    while (fnr.Peek() >= 0)
                    {
                        string stmp = fnr.ReadLine().Replace("\n", "");
                        stmp = stmp.Replace("\r\n", "");
                        if (stmp != "")
                        {
                            sDataContent[iIdx] = stmp;
                            iIdx += 1;
                        }
                    }
                    fnr.Close();
                    //PD_TXTParser(sDataContent, ref oEDCDataParms, ref GetLotEDCData);
                    break;

                case 10:
                    //LKParser(FullPath, ref GetLotEDCData);
                    break;

                case 14:
                case 15:
                case 16:
                    if (iServiceType == 16)
                    {
                        if (FullPath.IndexOf("_ID") > -1) iServiceType = 1;  //For ID use
                        if (FullPath.IndexOf("_OD") > -1) iServiceType = 2; //For OD use
                    }

                    if (FullPath.IndexOf("-X") > -1) iServiceType = 4;    //For AAC-X
                    if (FullPath.IndexOf("-Y") > -1) iServiceType = 5;   //For AAC-Y

                    fnr = new StreamReader(FullPath);
                    sDataContent = new string[1];
                    while (fnr.Peek() >= 0)
                    {
                        string stmp = fnr.ReadLine().Replace("\r\n", "");
                        stmp = stmp.Replace("\n", "");
                        if (stmp.IndexOf('=') > -1)
                        {
                            string[] a = stmp.Split('=');
                            Array.Resize(ref sDataContent, iIdx + 1);
                            sDataContent[iIdx] = a[1].Trim();
                            iIdx += 1;
                        }
                    }
                    fnr.Close();
                    //OMI_RPTParser(sDataContent, ref GetLotEDCData, iServiceType, oEDCDataParms.Readings);
                    break;

                case 17:

                    fnr = new StreamReader(FullPath);
                    sDataContent = new string[1];
                    while (fnr.Peek() >= 0)
                    {
                        string stmp = fnr.ReadLine().Replace("\n", "");
                        stmp = stmp.Replace("\r\n", "");
                        if (stmp != "")
                        {
                            Array.Resize(ref sDataContent, iIdx + 1);
                            sDataContent[iIdx] = stmp;
                            iIdx += 1;
                        }
                    }
                    fnr.Close();

                    //Ni_Bubble_Parser(sDataContent, ref GetLotEDCData);
                    break;

                case 19:    //20141128 add for Reflow Max Temp.
                    fnr = new StreamReader(FullPath);
                    sDataContent = new string[1];
                    while (fnr.Peek() >= 0)
                    {
                        string stmp = fnr.ReadLine().Replace("\n", "");
                        stmp = stmp.Replace("\r\n", "");
                        if (stmp != "")
                        {
                            Array.Resize(ref sDataContent, iIdx + 1);
                            sDataContent[iIdx] = stmp;
                            iIdx += 1;
                        }
                    }
                    fnr.Close();

                    ReflowMaxTempParser(sDataContent, ref GetLotEDCData);
                    break;

                default:
                    break;
            }

            string sKeepReadings = "";
            if (oEDCDataParms.Readings == 19)
            {
                sKeepReadings = clsEDCDataParms.Readings19;
            }
            else if (oEDCDataParms.Readings == 9)
            {
                sKeepReadings = clsEDCDataParms.Readings9;
            }

            if (sKeepReadings.Length > 0 && GetLotEDCData.Rows.Count == 49)
            {
                foreach (DataRow oRow in GetLotEDCData.Rows)
                {
                    if (sKeepReadings.IndexOf("," + oRow["No"].ToString() + ",") == -1)
                    {
                        oRow.Delete();
                    }
                }
            }
            GetLotEDCData.AcceptChanges();

            //  'Re-Index
            iIdx = 0;
            foreach (DataRow oRow in GetLotEDCData.Rows)
            {
                oRow["No"] = iIdx;
                iIdx += 1;
            }

            GetLotEDCData.AcceptChanges();

            // Average
            double Row_Avg = 0;
            double Row_sum = 0;
            foreach (DataRow oRow in GetLotEDCData.Rows)
            {
                Row_sum += Convert.ToDouble(oRow["Data"].ToString());
            }
            Row_Avg = Row_sum / GetLotEDCData.Rows.Count;

            //  'OffSet
            iIdx = 0;
            if (oEDCDataParms.OffSetArith.Length > 0)
            {
                foreach (DataRow oRow in GetLotEDCData.Rows)
                {
                    iIdx += 1;
                    if (iIdx >= oEDCDataParms.OffSetAfterPoint)
                    {
                        double tem = 0;
                        switch (oEDCDataParms.OffSetArith)
                        {
                            case "+":
                                oRow["Data"] = Convert.ToDouble(oRow["Data"].ToString()) + Convert.ToDouble(oEDCDataParms.OffSetValue.ToString("#0.0000"));
                                //oRow["Data"] = oRow["Data"].ToString();                                    
                                break;

                            case "-":
                                oRow["Data"] = Convert.ToDouble(oRow["Data"].ToString()) - Convert.ToDouble(oEDCDataParms.OffSetValue.ToString("#0.0000"));
                                //oRow["Data"] = oRow["Data"].ToString();
                                break;

                            case "*":
                                oRow["Data"] = Convert.ToDouble(oRow["Data"].ToString()) * Convert.ToDouble(oEDCDataParms.OffSetValue.ToString("#0.0000"));
                                //oRow["Data"] = oRow["Data"].ToString();
                                break;

                            case "/":
                                oRow["Data"] = Convert.ToDouble(oRow["Data"].ToString()) / Convert.ToDouble(oEDCDataParms.OffSetValue.ToString("#0.0000"));
                                //oRow["Data"] = oRow["Data"].ToString();
                                break;

                            case "R:+":
                            case "R:-":
                            case "R:*":
                            case "R:/":
                                tem = (Row_Avg + ((Convert.ToDouble(oRow["Data"].ToString()) - Row_Avg) / Convert.ToDouble(oEDCDataParms.OffSetValue.ToString("#0.0000")))) * 10000;
                                oRow["Data"] = tem.ToString("#0.0000");
                                break;

                            default:
                                break;
                        }

                    }
                }
            }
            GetLotEDCData.AcceptChanges();

            if (BackPath.Length > 0)
            {
                if (iServiceType == 3 || iServiceType == 4 || iServiceType == 9 || iServiceType == 12 || iServiceType == 139)
                    return;
                File.Copy(FullPath, BackPath + "\\" + backupFileName);

                File.Delete(FullPath);
                //File.Copy(FullPath, BackPath + "\\" + FileName.Replace(oEDCDataParms.LotNo, oEDCDataParms.LotNo + "_" + System.DateTime.Now.ToString("yyyyMMddHHmmss")));
            }
        }
        catch (Exception)
        {
            throw;
        }
    }

    private static void ReflowMaxTempParser(string[] sDataContent, ref System.Data.DataTable GetLotEDCData)
    {
        try
        {
            bool bStartRead = false;
            int iStartIdx = 0;
            double dMax = 0;
            DataRow dr;
            for (int i = 0; i < sDataContent.Length - 1; i++)
            {
                if (sDataContent[i] != null)
                {
                    string[] b = sDataContent[i].Split('\t');

                    if (bStartRead)
                    {
                        dMax = Math.Max(dMax, Convert.ToDouble(b[2]));
                        dMax = Math.Max(dMax, Convert.ToDouble(b[4]));
                        dMax = Math.Max(dMax, Convert.ToDouble(b[6]));
                        dMax = Math.Max(dMax, Convert.ToDouble(b[8]));
                        dMax = Math.Max(dMax, Convert.ToDouble(b[10]));

                    }

                    if (b[0] == "Scan") //Into Black get relation data.
                    {
                        bStartRead = true;
                    }
                }
            }

            dr = GetLotEDCData.NewRow();
            dr["No"] = iStartIdx;
            dr["WID"] = "1";
            dr["Data"] = dMax.ToString();
            GetLotEDCData.Rows.Add(dr);
            iStartIdx += 1;
            GetLotEDCData.AcceptChanges();
        }
        catch (Exception)
        {
            throw;
        }
    }

    private static void MAPParser(string[] sDataContent, ref System.Data.DataTable GetLotEDCData, int iReadings)
    {
        try
        {
            bool bStartRead = false;
            string[] aData;
            int iStartIdx = 0;
            DataRow dr;
            for (int i = 0; i < sDataContent.Length - 1; i++)
            {
                if (sDataContent[i] != null)
                {
                    if (sDataContent[i].IndexOf("|") > 0 & sDataContent[i].Substring(0, 1) == "0")
                    {
                        bStartRead = true;
                    }

                    if (bStartRead & sDataContent[i].IndexOf("|") > 0)
                    {
                        aData = sDataContent[i].Split('|');
                        dr = GetLotEDCData.NewRow();
                        dr["No"] = iStartIdx;
                        dr["WID"] = "1";
                        dr["Data"] = aData[3].ToString();
                        GetLotEDCData.Rows.Add(dr);
                        iStartIdx += 1;
                    }
                }
            }
            GetLotEDCData.AcceptChanges();
        }
        catch (Exception)
        {
            throw;
        }
    }

    private static void RSMParser(string[] sDataContent, ref System.Data.DataTable GetLotEDCData, int iReadings)
    {
        try
        {
            bool bStartRead = false;
            string sData = "";
            int iStartIdx = 0;
            DataRow dr;
            for (int i = 0; i < sDataContent.Length; i++)
            {
                if (sDataContent[i] != null)
                {
                    if (bStartRead)
                    {
                        sData = sDataContent[i].Substring(19, 12);
                        dr = GetLotEDCData.NewRow();
                        dr["No"] = iStartIdx;
                        dr["WID"] = "1";
                        dr["Data"] = sData;
                        GetLotEDCData.Rows.Add(dr);
                        iStartIdx += 1;
                    }

                    if (sDataContent[i].IndexOf("CassSlotSelected") > 0)
                    {
                        bStartRead = true;
                    }
                }
            }
            GetLotEDCData.AcceptChanges();
        }
        catch (Exception)
        {
            throw;
        }
    }

    class clsEDCDataParms
    {
        public string Path;
        public string BPath;
        public string APath;
        public string LotNo;
        public int Readings;
        public bool Compare;
        public string OffSetArith;
        public double OffSetValue;
        public int OffSetAfterPoint;
        public string MainFolder;
        public string Backup;

        public const string Readings19 = ",0,9,10,12,13,14,16,25,26,28,29,30,32,33,34,36,37,38,40,";
        public const string Readings9 = ",7,11,15,17,24,31,33,37,41,";
    }

    static string CopyFile(string sLotNo, string sRSC)
    {
        //\\t1review02\edcdata\T1_BP_BMBR001
        PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "Start");
        try
        {
            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sSQL = "Select * from att_ecim.ecimservice_eqpinfo where name= '" + sRSC + "' and TYPE =' '";
            ds = sObjService.AMSDBQuery(sSQL);   //0:EQPName 1:IP 2:Folder 3:LoginName 4:LoginPWD 5: File

            string FullPath = @"\\" + ds.Tables[0].Rows[0][1].ToString() + @"\" + ds.Tables[0].Rows[0][2].ToString();
            string DestinationPath = @"\\t1review02\\edcdata\\" + sRSC;



            //For SIT use     
            //FullPath = "D:\\EQP_Report\\";
            //DestinationPath = "D:\\" + sRSC + @"\";
            //prepare local folder
            PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "prepare local folder");
            //String sTempFolder = @"D:\\temp\";
            String sTempFolder = @"D:\\temp\\" + sRSC + @"\";
            if (!Directory.Exists(sTempFolder))
                Directory.CreateDirectory(sTempFolder);

            //Copy file from EQP. to Local
            PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "Copy file from EQP. ( " + FullPath + "sRecipe" + "-" + sLotNo + ".L ) to Local");
            foreach (string sFiles in Directory.GetFiles(FullPath))
            {
                FileInfo fFile = new FileInfo(sFiles);
                string sFilename = fFile.Name.ToUpper();
                //string Flag = ConfigurationManager.AppSettings["EnableFilter"].ToString(); 
                //if (sFiles.Contains(sLotNo))
                if (sFilename.Contains(sLotNo))
                {
                    PubUtil.WriteLog("Copy File", "Copy File", sTempFolder, sFilename, "Copy file from EQP. to (" + sTempFolder + ") Start");
                    fFile.CopyTo((sTempFolder + sFilename), true);
                    PubUtil.WriteLog("Copy File", "Copy File", sTempFolder, sFilename, "Copy file from EQP. to (" + sTempFolder + ") End");
                }
            }

            PubUtil.setLogonDomain("tw", "bumpfab", "Fabbump9");//
            //PubUtil.setLogonDomain("tw", "admt5", "T5@16847968");
            if (!Directory.Exists(DestinationPath))
                Directory.CreateDirectory(DestinationPath);

            //Process DestinationPath data
            try
            {
                FileHandle(DestinationPath + @"\");
            }
            catch (Exception ex)
            {
                PubUtil.WriteLog("Exception", "File Backup Function has exception!", sRSC, sLotNo, ex.Message);
            }

            //Copy file from Local to \\t1review02\edcdata\sRSC
            PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "Copy file from Local to " + DestinationPath);
            foreach (string sFiles in Directory.GetFiles(sTempFolder))
            {
                FileInfo fFile = new FileInfo(sFiles);
                string sFilename = fFile.Name;
                //string Flag = ConfigurationManager.AppSettings["EnableFilter"].ToString();
                if (sFiles.Contains(sLotNo))
                {
                    PubUtil.WriteLog("Copy File", "Copy File", sRSC, sFilename, "Copy Start");
                    fFile.CopyTo((DestinationPath + @"\" + sFilename), true);
                    fFile.Delete();
                    PubUtil.WriteLog("Copy File", "Copy File", sRSC, sFilename, "Copy End");
                }
            }
            PubUtil.WriteLog("Copy File", "Copy File", sRSC, sLotNo, "Completed");
            return "";
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Error", "Exception", "", "", ex.Message);
            return "";
        }

    }

    //20160428 add by Elton for Purge Big Table Data
    public static string PurgeBigTable(string sTableName, string sCondition, int iRowNum)
    {
        string strMsg = "";
        try
        {
            string sSQL = "";
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();

            int i = 0;

            while (true)
            {
                DateTime start = DateTime.Now;
                DataSet ds = new DataSet();

                sSQL = "delete from " + sTableName + " where " + sCondition + " and rownum <= " + iRowNum;
                System.Diagnostics.Debug.WriteLine(sSQL);
                PubUtil.WriteLog("eCIMWebServiceLog", "PurgeBigTable", "sSQL ==", "", sSQL);

                ds = sObjService.AMSDBQuery(sSQL);
                i++;

                double duringTime = (DateTime.Now - start).TotalSeconds;
                System.Diagnostics.Debug.WriteLine("Run Times -- " + i + " , " + duringTime.ToString());
                PubUtil.WriteLog("eCIMWebServiceLog", "PurgeBigTable", "Times--", i.ToString(), duringTime.ToString());

                if (duringTime <= 1)
                {
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "PurgeBigTable", "", "", ex.Message);
            return ex.Message;
        }

        return strMsg;
    }

    //20160528 add by Elton for Connector Server health checking
    public static string HealthChecking(string sPordLine)
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "HealthChecking", "*****", "*****", "Start");
        string sHighLightMsg = "";
        string sLowLightMsg = "\r\n\r\nNormal List\r\n==================================================";
        string strMsg = "";
        Hashtable hIP = new Hashtable();
        try
        {
            string[] sLine = ConfigurationManager.AppSettings[sPordLine + "HealthChecking"].ToString().Split(',');

            for (int i = 0; i < sLine.Length; i++)
            {
                string sPath = ConfigurationManager.AppSettings[sLine[i]].ToString();
                System.Diagnostics.Debug.WriteLine("File Path == " + sPath);

                ////Check IP's C$ and D$ Free Space
                //string sIP = sPath.Substring(0, sPath.IndexOf('$') - 1);
                //if (!hIP.ContainsKey(sIP))
                //{
                //    System.Diagnostics.Debug.WriteLine("IP == " + sIP);
                //    hIP.Add(sIP, sIP);
                //    // Make a reference to a directory.
                //    DirectoryInfo di = new DirectoryInfo(sIP + "C$");

                //    System.IO.f

                //    // Get a reference to each directory in that directory.
                //    DirectoryInfo[] diArr = di.GetDirectories();

                //    // Display the names of the directories.
                //    foreach (DirectoryInfo dri in diArr)
                //        Console.WriteLine(dri.Name);
                //}


                //Check ErrLog_yyyymmdd.txt file size
                foreach (string sFiles in Directory.GetFiles(sPath + @"data\"))
                {
                    FileInfo fFile = new FileInfo(sFiles);
                    string sFilename = fFile.Name.ToUpper();
                    if (sFilename.Contains("ERRLOG_"))
                    {
                        strMsg = "";
                        long l = fFile.Length / 1024;
                        //l = l * 1024;   //test only
                        strMsg += "\r\nFile Name == " + fFile.FullName + ", File KB Size == " + l.ToString();
                        if (l >= 10240)
                        {
                            sHighLightMsg += strMsg + ", >10 MB Need to Check!!";
                        }
                        else
                        {
                            sLowLightMsg += strMsg;
                        }
                    }
                }

                //Check System.err.txt
                foreach (string sFiles in Directory.GetFiles(sPath + @"error\"))
                {
                    FileInfo fFile = new FileInfo(sFiles);
                    string sFilename = fFile.Name.ToUpper();
                    if (sFilename.Contains("SYSTEM"))
                    {
                        strMsg = "";
                        long l = fFile.Length / 1024;
                        //l = l * 1024;   //test only
                        strMsg += "\r\nFile Name == " + fFile.FullName + ", File KB Size == " + l.ToString();
                        if (l >= 3072)
                        {
                            sHighLightMsg += strMsg + ", >3 MB Need to Check!!";
                        }
                        else
                        {
                            sLowLightMsg += strMsg;
                        }
                    }
                }
            }

            if (sHighLightMsg.Length > 0)
            {
                sHighLightMsg = "\r\nSpecial HighLight\r\n==================================================" + sHighLightMsg;
            }
            strMsg = sHighLightMsg + sLowLightMsg;

            System.Diagnostics.Debug.WriteLine("Checking Msg == " + strMsg);
            PubUtil.WriteLog("eCIMWebServiceLog", "HealthChecking", "Checking Msg == ", "", strMsg);
            PubUtil.WriteLog("eCIMWebServiceLog", "HealthChecking", "", "", "Completed\r\n");

            //Send Mail to Admin
            string to = ConfigurationManager.AppSettings[sPordLine + "MailTo"];
            string from = ConfigurationManager.AppSettings["MailAdmin"];
            MailMessage message = new MailMessage(from, to);
            message.Subject = "Health Checking for Connector";
            message.Body = sHighLightMsg + "\r\n" + sLowLightMsg;
            message.CC.Add(from);
            SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["smtpHost"]);
            //client.UseDefaultCredentials = true;
            //client.Port = Int32.Parse(ConfigurationManager.AppSettings["smtpPort"]);

            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                PubUtil.WriteLog("Exception", "HealthChecking", "Send Mail to Admin", "=====", ex.Message);
            }
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "HealthChecking", "=====", "=====", ex.Message);
            return ex.Message;
        }

        return sHighLightMsg + "\r\n" + sLowLightMsg;
    }

    /*
    public static void setWaferProcessTime(string sValue)
    {
        try
        {
            sValue = "Insert|T5-BPLT05|BUM6320164|1|2011/09/17 15:33:03|||";    //wafer Start test only
            //sValue = "Update|T5-BPLT05|BUM6320164|1||2011/09/17 15:33:03|||";   //wafer  End  test only
            //sValue = "Insert|T5-BPLT05|BUM6320164|1|2011/09/17 15:33:03|2011/09/17 15:50:03||";    //wafer Start/End test only

            string[] saValue = sValue.Split('|');
            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton

            //
            
            sSQL = "insert into " + FILE + " values (?,?,?,?,?, ?,?,?,?,?, ?,?)";

            ds = sObjService.AMSDBQuery(sSQL);

            //if (ds != null && ds.Tables.Count == 0) return;

            //sSQL = "update insitedev.amkaas_lot_validate set uomname=" +
            //        "(select NVL(UOMNAME,0) + 1 UOMNAME from insitedev.amkaas_lot_validate " +
            //        "where CARRIER='" + sLotNo + "' AND CARRIER = CONTAINERNAME) " +
            //        "where CARRIER='" + sLotNo + "'";
            //PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, " Update setWaferProcessCount SQL" + sSQL);

            ////sSQL = "update insitedev.amkaas_lot_validate set uomname=" + ds.Tables[0].Rows[0][0].ToString() + "where CARRIER='" + sLotNo + "'";

            //sObjService.AMSDBQuery(sSQL);

            //PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, "setWaferProcessCount ==>" + ds.Tables[0].Rows[0][0].ToString());
        }
        catch (Exception ex)
        {
            //PubUtil.WriteLog("Exception", "setWaferProcessTime", sRSC, sLotNo, ex.Message);
            throw;
        }
    }*/

    //把DataTable轉成JSON字串
    public static string getDateSet2Json(string sSQL)
    {
        string sReturnValue = "Not Found!!";
        try
        {
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton

            //得到一個DataTable物件
            //sSQL = "select * from amkaas_lot_validate where carrier = 'BP6330631' and resourcename = 'T1_BP_BOVN012'"; //Test Only

            ds = sObjService.AMSDBQuery(sSQL);

            //將DataTable轉成JSON字串
            if (ds.Tables.Count > 0)
            {
                string str_json = JsonConvert.SerializeObject(ds.Tables[0]);
                //str_json = JsonConvert.SerializeObject(ds.Tables[0], Newtonsoft.Json.Formatting.Indented);

                //JSON字串顯示在畫面上
                sReturnValue = str_json;
            }

            //把JSON字串轉成DataTable或Newtonsoft.Json.Linq.JArray
            //Newtonsoft.Json.Linq.JArray jArray = 
            //    JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JArray>(li_showData.Text.Trim());
            //或
            //DataTable dt = JsonConvert.DeserializeObject<DataTable>(sReturnValue);

            //PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotNo, sReturnValue);
            return sReturnValue;
        }
        catch (Exception ex)
        {
            //PubUtil.WriteLog("Exception", "GetProcessWaferCount", sRSC, sLotNo, ex.Message);
            return "Exception :: " + ex.Message;
        }
    }

    //Get TM Pressure From Data Log , 20161123 Added by Elton
    public static string getTMPressureData(string sLotNo, string sRSC)
    {
        string sReturnValue = "Not Found!!";
        try
        {
            bool bFindFile = false;
            string TrackInTime = "";
            DateTime dtTrackIn;

            //Get Track In Time
            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            //sSQL = "select trackindatetime from amkaas_lot_validate where resourcename like '%" + sRSC.ToUpper() + "%'";    //test only
            sSQL = "select trackindatetime from insitedev.amkaas_lot_validate where CARRIER = '" + sLotNo + "' and resourcename like '" + sRSC + "%' AND CARRIER = CONTAINERNAME ";
            ds = sObjService.AMSDBQuery(sSQL);   //0:EQPName 1:IP 2:Folder 3:LoginName 4:LoginPWD 5: File

            if (ds.Tables.Count > 0)
            {
                TrackInTime = ds.Tables[0].Rows[0][0].ToString();
                //TrackInTime = "2016/11/30 15:16:31"; //test only

                dtTrackIn = Convert.ToDateTime(TrackInTime);
                //getTMPressureDelayMin setting add by Elton 20170110
                dtTrackIn = dtTrackIn.AddMinutes(Convert.ToInt16(ConfigurationManager.AppSettings["getTMPressureDelayMinutes"].ToString()));

                //Is need to open another data log file
                Double d1 = Convert.ToDouble(System.DateTime.Now.ToString("yyyyMMdd"));
                Double d2 = Convert.ToDouble(dtTrackIn.ToString("yyyyMMdd"));
                if (d1 - d2 > 0) bFindFile = true;
            }
            else
            {
                PubUtil.WriteLog("Exception", "getTMPressureData", sRSC, sLotNo, "E0002:Can't find Validation record!");
                return "E0002:Can't find Validation record!";
            }
            //Open TM Pressure data log 1st
            if (!openTMPressureData(sLotNo, sRSC, dtTrackIn))
            {
                PubUtil.WriteLog("Exception", "getTMPressureData", sRSC, sLotNo, "Open File 1 fail!!");
                return "Open File 1 fail!!";
            }

            //Open TM Pressure data log 2nd
            if (bFindFile)
            {
                dtTrackIn = dtTrackIn.AddDays(1);
                if (!openTMPressureData(sLotNo, sRSC, dtTrackIn))
                {
                    PubUtil.WriteLog("Exception", "getTMPressureData", sRSC, sLotNo, "Open File 2 fail!!");
                    return "Open File 2 fail!!";
                }
            }
            PubUtil.WriteLog("eCIMWebServiceLog", "getTMPressureData", sRSC, sLotNo, "success");
            sReturnValue = "success";
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "getTMPressureData", sRSC, sLotNo, ex.Message);
            return ex.Message;
        }

        return sReturnValue;
    }

    private static bool openTMPressureData(string sLotNo, string sRSC, DateTime dtTrackIn)
    {
        string sReturnValue = "Not Found-1!!";
        try
        {
            string FullPath = "";
            string sFileName = "";
            string FileType = "";
            string sSQL = "";
            //Double dSpec = Convert.ToDouble("9e-06");
            //20170112 add by Elton for dynamic get TM Pressure Spec 
            string sSpec = "9.0E-6";

            string sTmp = "";
            try
            {
                sTmp = ConfigurationManager.AppSettings["getTMPressureSpec"].ToString();
            }
            catch (Exception)
            {
            }
            if (!sTmp.Equals(""))
            {
                sSpec = sTmp;
            }
            Double dSpec = Convert.ToDouble(sSpec);

            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton
            sSQL = "Select * from att_ecim.ecimservice_eqpinfo where name= '" + sRSC + "' and TYPE ='TMP'";
            ds = sObjService.AMSDBQuery(sSQL);   //0:EQPName 1:IP 2:Folder 3:LoginName 4:LoginPWD 5: File

            if (ds.Tables.Count == 0)
            {
                PubUtil.WriteLog("Exception", "openTMPressureData", sRSC, sLotNo, "Can't find the setting data");
                return false;
            }

            FullPath = @"\\" + ds.Tables[0].Rows[0][1].ToString() + @"\" + ds.Tables[0].Rows[0][2].ToString();
            PubUtil.setLogonDomain(FullPath, ds.Tables[0].Rows[0][3].ToString(), ds.Tables[0].Rows[0][4].ToString());

            FileHandle(FullPath);

            //TM pressure data log file name sample, 001_Pressure_BSPT05_TM_20161001
            FileType = "his";
            sFileName = "*" + dtTrackIn.ToString("yyyyMMdd") + "." + FileType;
            //sFileName = "001_Pressure_BSPT05_TM_20161130.HIS";    //test only

            //Find the File List
            DirectoryInfo dir1 = new DirectoryInfo(FullPath);
            FileInfo[] aFileInfo1 = dir1.GetFiles(sFileName);
            sFileName = "";
            foreach (FileInfo filename in aFileInfo1)
            {
                sFileName = filename.Name;
                FullPath = filename.FullName;
            }

            if (sFileName.Equals("")) return false;


            //Open File get Data log for Track In/Out Tiem
            StreamReader fnr = new StreamReader(FullPath);

            while (fnr.Peek() >= 0)
            {
                string stmp = fnr.ReadLine().Replace("\n", "");
                stmp = stmp.Replace("\r\n", "");
                if (stmp != "")
                {
                    string[] sa = stmp.Split('\t');

                    if (sa.Length > 3)
                    {
                        //Check Time Line > Track In Time
                        DateTime dtLine = Convert.ToDateTime(sa[0]);

                        if (dtLine > dtTrackIn)
                        {
                            //Check Spec > 9e-6, if not Hold Lot
                            if (Convert.ToDouble(sa[2]) > dSpec)
                            {
                                string sSubject = sa[0] + " -- " + sa[2] + " > TM Spec 9E-06, Please ask EE check " + sRSC;
                                System.Diagnostics.Debug.WriteLine(sSubject);
                                PubUtil.WriteLog("Exception", "openTMPressureData", sRSC, sLotNo, sa[0] + " -- " + Convert.ToDouble(sa[2]).ToString() + " > TM Spec 9E-06, Please ask EE check " + sRSC);

                                //> TM Spec 9E-06, Hold Lot
                                String HoldComment = "PE, TM Spec 9E-06! Lot-" + sLotNo + ", Eqp@" + sRSC + ". Hold by eCIM"; //Hold Comment

                                string sXMLdata = "<Root><TxnType>LotHold</TxnType><Employee>icsadmin</Employee><Container>" + sLotNo + "</Container><HoldReason>SYS HOLD</HoldReason><Comments>" + HoldComment + "</Comments><PortNo></PortNo><ClientIP>10.185.32.190</ClientIP></Root>";
                                //sObjService.AMSTxn(sXMLdata);//Need to Hold Lot

                                //20170103 add by Elton , Send Alert Mail first
                                PubUtil.sendAlert2Admin(sSubject, sXMLdata, "att1-bumping-ee1@amkor.com");
                                break;
                            }
                            //Write Output Stream to another file for Lot Base
                        }
                    }
                }
            }
            fnr.Close();
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "openTMPressureData", sRSC, sLotNo, ex.Message);
            return false;
        }

        return true;
    }

    //20170103 add by Elton for send alert mail
    public static void sendAlert2Admin(string sSubject, string sMessage, string sTo)
    {
        string sSMTP_IP = "";
        int iSMTP_Port = 0;
        string sFrom = "";// "ATT_eCIM_Admin@amkor.com";
        string sCC = "";

        sFrom = ConfigurationManager.AppSettings["SMTP_FROM"].ToString();
        sCC = ConfigurationManager.AppSettings["SMTP_CC"].ToString();
        sSMTP_IP = ConfigurationManager.AppSettings["SMTP_IP"].ToString();
        iSMTP_Port = Int16.Parse(ConfigurationManager.AppSettings["SMTP_PORT"].ToString());

        MailMessage message = new MailMessage(sFrom, sTo);
        message.Subject = sSubject;
        message.Body = sMessage;
        message.BodyEncoding = UTF8Encoding.UTF8;
        message.CC.Add(sCC);

        SmtpClient client = new SmtpClient(sSMTP_IP);
        client.Port = iSMTP_Port;
        client.DeliveryMethod = SmtpDeliveryMethod.Network;
        // Credentials are necessary if the server requires the client 
        // to authenticate before it will send e-mail on the client's behalf.
        client.UseDefaultCredentials = false;

        try
        {
            client.Send(message);
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "openTMPressureData", sSubject, sTo, ex.Message);
        }
    }

    // public static String getIsSWR(String sRSC, String sLotNo)
    public static Boolean getIsSWR(String sRSC, String sLotNo)
    {

        Boolean isSWR = false;
        string sSQL = "";
        try
        {
            PubUtil.WriteLog("eCIMWebServiceLog", "getIsSWR", sRSC, sLotNo, "Start");
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();
            sSQL = "select t.opercode from amkaas_lot_validate t where t.resourcename like '" + sRSC + "%'  and t.containername = '" + sLotNo + "' and t.containername = t.carrier ";
            ds = sObjService.AMSDBQuery(sSQL);
            if (ds != null && ds.Tables.Count > 0)
            {
                String sSWR = ds.Tables[0].Rows[0][0].ToString();
                if (sSWR.Equals("SWR"))
                {
                    isSWR = true;
                }
            }
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "getIsSWR", sRSC, sLotNo, ex.Message);
            //return Convert.ToString(isSWR);
            return isSWR;
        }
        PubUtil.WriteLog("eCIMWebServiceLog", "getIsSWR", sRSC, sLotNo, "Completed. isSWR:" + isSWR);
        //return Convert.ToString(isSWR);
        return isSWR;
    }

    //add by Jim for T3 LM get SEQ data 20180628
    public static string Get2DSerialRecordData(string sRSC, string sLotNo)
    {
        //sRSC = DP_LME012 
        PubUtil.WriteLog("eCIMWebServiceLog", "Get2DSerialRecordData", sRSC, sLotNo, "eCIM Get seq data Start.");
        try
        {
            string FullPath = "";
            string sSQL = "";
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton
            sSQL = "Select * from att_ecim.ecimservice_eqpinfo where name= '" + sRSC + "' and TYPE ='LMSEQ'";
            ds = sObjService.AMSDBQuery(sSQL);   //0:EQPName 1:IP 2:Folder 3:LoginName 4:LoginPWD 5: File

            FullPath = @"\\" + ds.Tables[0].Rows[0][2].ToString();
            return T3LM_Parser(FullPath, sLotNo);

        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "Get2DSerialRecordData", sRSC, sLotNo, ex.Message);
            throw;
        }
    }

    //add by Jim for T3 LM get SEQ data 20180628
    private static string T3LM_Parser(string FullPath, string sLotID)
    {
        string FileName = "";
        StreamReader fnr;
        string[] sDataContent;
        DirectoryInfo dir;
        string stmp = "";
        PubUtil.WriteLog("eCIMWebServiceLog", "T3LM_Parser", FullPath, sLotID, "T3LM_Parser Start.");
        try
        {
            FileName = sLotID + ".seq";
            if (!File.Exists(FullPath + FileName))
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "T3LM_Parser", FullPath, sLotID, "Can not get File from." + FullPath + FileName);

                return "Please check File Path " + FullPath + FileName;
            }
            else
            {
                dir = new DirectoryInfo(FullPath);
            }

            FileInfo[] aFileInfo = dir.GetFiles(FileName);

            if (aFileInfo.Length == 1)
            {
                FullPath += aFileInfo[0].Name;
                FileName = aFileInfo[0].Name;
            }
            else
            {
                string sFileName = "";
                FileName = sFileName;
                FullPath += sFileName;
            }

            FullPath = FullPath.Replace("\\", @"\");
            fnr = new StreamReader(FullPath);
            sDataContent = new string[1000];

            while (fnr.Peek() >= 0)
            {
                stmp = fnr.ReadToEnd().Replace("\n", "");
            }
            fnr.Close();
            PubUtil.WriteLog("eCIMWebServiceLog", "T3LM_Parser", FullPath, sLotID, "Get file complete. " + stmp);
            return stmp;
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "T3LM_Parser", FullPath, sLotID, ex.Message);
            return ex.ToString();
        }
    }

    //add by Jim for add new method 20180628 
    public static string eCIMTxn(string type, string value)
    {
        string result = "";
        PubUtil.WriteLog("eCIMWebServiceLog", "eCIMTxn", type, value, "eCIMTxn Start.");
        switch (type)
        {
            case "0001":
                string[] aValue = value.Split(',');
                if (aValue.Length < 2) return "E0001: Input Data not enought";
                return Get2DSerialRecordData(aValue[0], aValue[1]);
            default:
                break;
        }

        return result;
    }


    //add by Jim on 20181019 for when run lot, edit/upload recipe, check FDC status
    //add by Jim on 20181224 for when run lot, enhance edit/upload recipe, check FDC status
    public static string GetFDC_Master(string _Line, string _RSC, string _RCP, string _Equal)
    {
        try
        {
            PubUtil.WriteLog("FDC_Check_Log", "GetFDC_Master", _RSC, _RCP, "Flag = " + _Equal + " Start.");
            bool bEqual = false;//default use * rule
            try
            {
                bEqual = Convert.ToBoolean(_Equal);
            }
            catch (Exception ex)
            {
            }

            string sRCP = _RCP.ToUpper();
            string sSQL = "select * from ATT_ECIM.LimitDefinitionMaster t where t.Division = '" + _Line + "' and t.machid = '" + _RSC + "' and t.status = 'Approved'";
            PubUtil.WriteLog("FDC_Check_Log", "GetFDC_Master", _RSC, _RCP, "SQL = " + sSQL + ";");
            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();  //20140910 add by Elton
            ds = sObjService.AMSDBQuery(sSQL);

            bool bFind = false;

            if (ds.Tables.Count > 0)//only found record will go throught by Jim 20190109
            {
                //0:SN,1:Rev,2:Status,3:Div,4:oper,5:model,6:MC ID,14:RCP
                foreach (DataRow dtRow in ds.Tables[0].Rows)
                {
                    string rcp = dtRow[14].ToString();

                    if (!bEqual)
                    {
                        if (rcp.Equals("*") || rcp.Trim().Length == 0)// Recipe Name = * or ""
                        {
                            bFind = true;
                        }
                        else if (rcp.IndexOf("*") > -1)//12DE-*
                        {
                            string temp = rcp;
                            bool bStartWith = false;
                            bool bBothStart = false;
                            if (temp.StartsWith("*") && temp.EndsWith("*"))//*CO*
                            {
                                bBothStart = true;
                            }
                            else if (temp.StartsWith("*"))//*CO
                            {
                                bStartWith = true;
                            }

                            temp = temp.Replace("*", "");

                            if (bBothStart) //*CO*
                            {
                                if (_RCP.IndexOf(temp) > -1)
                                {
                                    bFind = true;
                                }
                            }
                            else if (bStartWith) // *-CO
                            {
                                if (_RCP.EndsWith(temp))
                                {
                                    bFind = true;
                                }
                            }
                            else // 12DE-*
                            {
                                if (_RCP.StartsWith(temp))// start with 12DE or ...etc
                                {
                                    bFind = true;
                                }
                            }
                        }
                        else if (rcp.Equals(_RCP))
                        {
                            bFind = true;
                        }
                    }
                    else
                    {
                        if (rcp.Equals(_RCP))
                        {
                            bFind = true;
                        }
                    }
                    if (bFind)
                    {
                        string msg = "Success;" + dtRow[0] + "/" + dtRow[1] + "/" + dtRow[2] + "/" + dtRow[3] + "/" + dtRow[4] + "/" + dtRow[5] + "/" + dtRow[6] + "/" + dtRow[14];
                        PubUtil.WriteLog("FDC_Check_Log", "GetFDC_Master", _RSC, _RCP, "Return = " + msg);
                        return msg;
                    }
                }
                PubUtil.WriteLog("FDC_Check_Log", "GetFDC_Master", _RSC, _RCP, "Return = ");
                return "";
            }
            else
            {//can not find record
                PubUtil.WriteLog("FDC_Check_Log", "GetFDC_Master", _RSC, _RCP, "Get record 0 ! ");
                return "";
            }
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("FDC_Check_Log", "GetFDC_Master", "Exception", "", ex.ToString());
            return "";
        }
    }

    //20181211 add by Elton 
    public static string getFDCWaferCount(string sLotID, string sRSC, string sParameterName)
    {
        string sReturnValue = "Not Found!!";
        string sSQL = "";
        //sRSC = "T1_BP_BSPT004"; //Test Only

        try
        {
            //get the Move In time -15 mins
            string dtMoveIn = "";

            DataSet ds = new DataSet();
            AMSTxn.AMSTxnService sObjService = new AMSTxn.AMSTxnService();
            sObjService.Url = ConfigurationManager.AppSettings["AMSTxn.AMSTxnService"].ToString();
            sSQL = "Select distinct t.createdate from amkaas_lot_validate t where t.containername ='" + sLotID + "' and t.resourcename='" + sRSC + "'";
            ds = sObjService.AMSDBQuery(sSQL);

            dtMoveIn = ds.Tables[0].Rows[0][0].ToString();
            DateTime parsedDate = DateTime.Parse(dtMoveIn);
            parsedDate = parsedDate.AddMinutes(-15);
            dtMoveIn = parsedDate.ToString("yyyy/MM/dd HH:mm:ss");

            //get the Move out time as now
            string dtMoveOut = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            //prepare Query SQL
            clsDBUTility dbUtility = new clsDBUTility();

            string sConnString = "Provider=OraOLEDB.Oracle;Data Source=PROORCLFDC;User ID=att_ecim;Password=attecim#182";
            dbUtility.New(sConnString);
            dbUtility.OpenDBLink();

            // Bypass null parameter name
            if (sParameterName.Equals("")) return "Empty Parameter Name!";

            sSQL = "  select DISTINCT t.stringvalue from ATT_ECIM.FDCDetectionHistory t" +
                    " where t.machid ='" + sRSC + "'" +
                    " and t.createdtime >= to_date('" + dtMoveIn.ToString() + "','YYYY/MM/DD HH24:MI:SS')" +
                    " and t.createdtime <= to_date('" + dtMoveOut.ToString() + "','YYYY/MM/DD HH24:MI:SS')" +
                    " AND T.PARAMNAME LIKE '" + sParameterName + "%'" +
                    " and t.stringvalue like '" + sLotID + "%'";

            ds = dbUtility.SQLQuery(sSQL, true);

            if (ds.Tables.Count > 0)
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "GetFDCWaferCount", sRSC, sLotID, ds.Tables[0].Rows.Count.ToString());

                //write to Log file
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount", sRSC, sLotID, row[0].ToString());
                }

                sReturnValue = ds.Tables[0].Rows.Count.ToString();
            }
        }
        catch (Exception ex)
        {
            return "Exception :: " + ex.Message;
        }
        return sReturnValue;
    }
}
