using System;
using System.Web;
using System.Collections;
using System.Threading;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Data;
using System.Data.SqlClient;

/// <summary>
/// Summary description for eCIMService
/// </summary>
[WebService(Namespace = "http://localhost/eCIMServices/")]
//[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)] 
[WebServiceBinding(ConformsTo = WsiProfiles.None)]  //若是EFGP呼叫 則需要改用這個. add by chuck on 20180403
public class eCIMService : System.Web.Services.WebService {
    public eCIMService () {

        //Uncomment the following line if using designed components 
        //InitializeComponent(); 
    }

      [WebMethod]    //2012/5/23 add by Elton for Equipment process wafer count
    public string GetProcessWaferCount(string sLotNo, string sRSC, string sRecipe)
    {
        try
        {
           // sLotNo = "BPE4280322";
            //sRecipe = "1072-A3231-06A1";
          //  sRSC = "T1_BP_BSTE001";
            string sValue = PubUtil.GetProcessWaferCount(sLotNo, sRSC, sRecipe);
            return sValue;
        }
        catch (Exception)
        {
            return "";
        }
    }

    [WebMethod]    //2012/10/03 add by Mark for T3 FC/DPS Laser Marking Data
    //public string getMarkingData(string sLotNo, string sRSC, string RunEVN)
    public string getMarkingData(string sLotNo,  string RunEVN)
    {
        string sValue = null;
        try
        {
            //sValue =GetMarkingData.getMarkingData(sLotNo, sRSC, RunEVN);
            sValue = GetMarkingData.getMarkingData(sLotNo, RunEVN);
            return sValue;
        }
        catch (Exception)
        {
            return sValue;
        }
    }

    [WebMethod]    //2012/11/16 add by Mark for T3 DPS Laser Marking recipe
    public string getDPS_LM_Recipe(string sLotNo, string RunEVN)
    {
        string sValue = null;
        try
        {
            //sValue =GetMarkingData.getMarkingData(sLotNo, sRSC, RunEVN);
            sValue = GetMarkingData.Get_DPS_LM_Recipe(sLotNo, RunEVN);
            return sValue;
        }
        catch (Exception)
        {
            return sValue;
        }
    }



    [WebMethod]    //2013/01/29 add by Roger getBP_BMBR_Edc
    public string GetBMBREdcData(string sLotNo, string sRSC, string sRecipe)
    {
        try
        {
            string sValue = PubUtil.GetBMBREdcData(sLotNo, sRSC, sRecipe);
            return sValue;
        }
        catch (Exception)
        {
            return "";
        }
    }

    [WebMethod]    //2013/05/28 add by Stanley get_Plating_EDC
    public DataTable GetAlignerEdcData(string sRSC, string sType)
    {
        DataTable sValue;

        sValue = PubUtil.GetAlignerEdcData(sRSC, sType);
        return sValue;
     
    }
    
    [WebMethod]    //2013/07/11 add by Stanley get_OMI_EDC
    public int Get_TryMax_ProcessEndData(string LotNo, string RSC, string Recipe)
    {
        try
        {
           //string sValue = PubUtil.GetOMIEdcData(sType);
           // LotNo = "BP3260182";//For test
            //LotNo = "BPE3430080";
           // Recipe = "PBO-FORM-12";
            //Recipe = "Oxide-loss";
            int sValue = PubUtil.Get_TryMax_Data(LotNo, RSC, Recipe);
            return sValue;
        }
        catch (Exception)
        {
            return 0; 
        }
    }

    [WebMethod]    //2013/08/08 Add by Elton for Real Time check Wafer Count & Track Out Lot
    public int checkMassOut(string sLotNo, string sRSC, string sRecipe)
    {
        try
        {
            //sLotNo = "BU4270075";
            //sRSC = "T5-BCOA04";
            //sRecipe = "8003-12CO-5PI41";
            //Call Function only , directly return 1.
            //PubUtil.checkMassOut(sLotNo, sRSC,sRecipe);

            //20170810 new method call for thread
            //spawn out a new thread
            
            PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut Web Method", sLotNo, sRSC, "Start!!");
            ThreadProcessor_checkMassOut tp = new ThreadProcessor_checkMassOut(sLotNo, sRSC, sRecipe);

            Thread th = new Thread(new ThreadStart(tp.service));

            //mark it as a non-background thread
            th.IsBackground = false;

            //start
            th.Start();
            PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut Web Method", sLotNo, sRSC, "Completed.");
             
            return 1;
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("Exception", "checkMassOut Web Method", sRSC, sLotNo, "Exception:" + ex.Message);
            return 0;
        }
    }

    [WebMethod]    //2013/10/21 Add by Stanley for Real Time Track In Lot
    public int callMassIn(string sLotNo, string sRSC, string sPort)
    {
        try
        {
            //sLotNo = "BP3380206";
            //sRSC = "T1_BP_BPLT004";
            //Call Function only , directly return 1.
            PubUtil.callMassIn(sLotNo, sRSC, sPort);
            return 1;
        }
        catch (Exception)
        {
            return 0;
        }
    }

    [WebMethod]    //2014/06/25 Add by Chuck for T1 SCB
    public int SCB(string sLotNo, string sRSC)
    {
        try
        {
            //sLotNo = "BP4330450";
            //sRSC = "T1_BP_BSCB003";
            //sLotNo = "BP4290565";
            //sRSC = "T1_BP_BSCB003";
            //Call Function only , directly return 1.
            PubUtil.GetProcessWaferCount(sLotNo, sRSC,"");
            return 1;
        }
        catch (Exception)
        {
            return 0;
        }
    }

    [WebMethod]    //2013/08/08 Add by Elton for Real Time update Process Wafer Count by Lot + RSC
    public int setWaferProcessCount(string sLotNo, string sRSC)
    {
        try
        {
            //sLotNo = "BP4330450";
            //sRSC = "T1_BP_BSCB003";
            //Call Function only , directly return 1.
            //PubUtil.setWaferProcessCount(sLotNo, sRSC);
            //return 1;

            //bool isNeedRun = false;
            //isNeedRun = PubUtil.isExistValitationTable(sLotNo, sRSC);

            //if (isNeedRun)
            //{
            //    PubUtil.setWaferProcessCount(sLotNo, sRSC);
            //}
            //else {
            //    PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, "No need to setWaferProcessCount");
            //}

            //20161110 add by Elton for New AMS Wafer Start/End Txn Handle
            string sNewLotNo = sLotNo;
            if (sLotNo.IndexOf('-') > -1)
            {
                sLotNo = sLotNo.Substring(0, sLotNo.IndexOf('-'));

                // If Wafer Start, no need to call setWaferProcessCount
                if (sNewLotNo.Substring(sNewLotNo.Length - 2).Equals("-1"))
                {
                    //PubUtil.setAmkLotWfrProc(sNewLotNo, sRSC);//20171024 mark by Elton
                    //20171024 change new method call to thread
                    try
                    {
                        PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc-1 Web Method", sNewLotNo, sRSC, "Start!!");
                        ThreadProcessor_setAmkLotWfrProc tp = new ThreadProcessor_setAmkLotWfrProc(sNewLotNo, sRSC);

                        Thread th = new Thread(new ThreadStart(tp.service));

                        //mark it as a non-background thread
                        th.IsBackground = false;
                        //start
                        th.Start();
                        PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc-1 Web Method", sNewLotNo, sRSC, "Completed.");
                    }
                    catch (Exception ex)
                    {
                        PubUtil.WriteLog("Exception", "setAmkLotWfrProc-1 Web Method", sRSC, sNewLotNo, "Exception:" + ex.Message);
                        return 0;
                    }
                    return 1;
                }
            }

            //bool isNeedRun_1 = false;
            //bool isNeedRun_2 = false;
            //String sPort = "";

            //isNeedRun_1 = PubUtil.isExistValitationTable(sLotNo, sRSC);

            //if (isNeedRun_1)
            //{
            //    isNeedRun_2 = PubUtil.isSYSHOLD(sLotNo, sPort, sRSC);
            //    if (isNeedRun_2)
            //    {
                    //PubUtil.setWaferProcessCount(sLotNo, sRSC);//20171024 mark by Elton
                    //20171024 change new method call to thread
                    try
                    {
                        PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount Web Method", sLotNo, sRSC, "Start!!");
                        ThreadProcessor_setWaferProcessCount tp = new ThreadProcessor_setWaferProcessCount(sLotNo, sRSC);

                        Thread th = new Thread(new ThreadStart(tp.service));

                        //mark it as a non-background thread
                        th.IsBackground = false;
                        //start
                        th.Start();
                        PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount Web Method", sLotNo, sRSC, "Completed.");
                    }
                    catch (Exception ex)
                    {
                        PubUtil.WriteLog("Exception", "setWaferProcessCount Web Method", sRSC, sLotNo, "Exception:" + ex.Message);
                        return 0;
                    }
            //    }
            //    else
            //    {
            //        PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, "No need to setWaferProcessCount, because SYS HOLD");
            //        return 1;
            //    }
            //}
            //else
            //{
            //    PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, "No need to setWaferProcessCount");
            //    return 1;
            //}

            //20161110 add by Elton for New AMS Wafer Start/End Txn Handle
            if (sNewLotNo.IndexOf('-') > -1)
            {
                //PubUtil.setAmkLotWfrProc(sNewLotNo, sRSC);//20171024 mark by Elton
                //20171024 change new method call to thread
                try
                {
                    PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc-2 Web Method", sNewLotNo, sRSC, "Start!!");
                    ThreadProcessor_setAmkLotWfrProc tp = new ThreadProcessor_setAmkLotWfrProc(sNewLotNo, sRSC);

                    Thread th = new Thread(new ThreadStart(tp.service));

                    //mark it as a non-background thread
                    th.IsBackground = false;
                    //start
                    th.Start();
                    PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc-2 Web Method", sNewLotNo, sRSC, "Completed.");
                }
                catch (Exception ex)
                {
                    PubUtil.WriteLog("Exception", "setAmkLotWfrProc-2 Web Method", sRSC, sNewLotNo, "Exception:" + ex.Message);
                    return 0;
                }
            }

            return 1;
        }
        catch (Exception)
        {
            return 0;
        }
    }

    [WebMethod]    //2013/08/08 Add by Elton for Real Time update Process Wafer Count by RSC + Port
    public int setWaferProcessCountbyPort(string sRSC, string sPort)
    {
        //try
        //{
        //    //Call Function only , directly return 1.
        //    PubUtil.setWaferProcessCountbyPort(sRSC, sPort);
        //    return 1;
        //}
        //catch (Exception)
        //{
        //    return 0;
        //}

        //Chuck add for fix SQL issue on 20150831
        String sLotNo = "";
        try
        {
            bool isNeedRun = false;
            isNeedRun = PubUtil.isSYSHOLD(sLotNo, sPort, sRSC);

            if (isNeedRun)
            {
                PubUtil.setWaferProcessCountbyPort(sRSC, sPort);
            }
            else
            {
                PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount", sRSC, sLotNo, "No need to setWaferProcessCountbyPort, because SYS HOLD or Valitation Table is Null");
            }
            return 1;
        }
        catch (Exception)
        {
            return 0;
        }
    }

    [WebMethod]    //2013/08/2 copy by Roger getBP_BMBR_Edc
    public string GetAhtleteMassOut(string sLotNo, string sRSC, string sRecipe)
    {
        try
        {
            //sLotNo = "BP3380206";
            //sRSC = "T1_BP_BMBR001";
            //sRecipe = "A3524_SAC405_25";
            //        FileName = sRecipe + "-" + sLotID + "." + FileType;
            //A3524_SAC405_25-bp3320288.L
            string sValue = PubUtil.GetProcessWaferCount(sLotNo, sRSC, sRecipe);
            return sValue;
        }
        catch (Exception)
        {
            return "";
        }
    }

     [WebMethod]    //2013/10/09 add by Mark for T3 DPS PPM recipe
    public string getDPS_PPM_Recipe(string sLotNo, string RunEVN)
    {
        string sValue = null;
        try
        {
            //sValue =GetMarkingData.getMarkingData(sLotNo, sRSC, RunEVN);
            sValue = GetMarkingData.Get_DPS_PPM_Recipe(sLotNo, RunEVN);
            return sValue;
        }
        catch (Exception)
        {
            return sValue;
        }
    }
    [WebMethod]    //20140407 add by Mark for T3 DPS PnP Besi source device
    public string getDPS_Besi_PnP_SRDevice(string sWaferID)
    {
        string sResult = "";
        try
        {
            System.Data.DataSet ds = new System.Data.DataSet();

            AMSWebReference.AMSWebService ws = new AMSWebReference.AMSWebService();
            ds = ws.AMSDBQuery2("select distinct p.srcdevice from waferdetails wd,container c,product p where wd.waferid='" + sWaferID + "' and wd.baseobjectid=c.containerid and c.productid=p.productid and substr(c.containername,1,2)='DP'", "ecimwebservice");

            if (null == ds || ds.Tables.Count != 1 || ds.Tables[0].Rows.Count < 1)
                sResult = "";
            else
            {
                sResult = ds.Tables[0].Rows[0]["srcdevice"].ToString();
            }
        }
        catch (Exception ex)
        {
            sResult = "";
        }
        return sResult;

    }
    [WebMethod]    //20140407 add by Mark for T3 update MC status for Production Index system
    public string updateMcStatusToProdIndex(string sMachineID, string Status)
    {
        string sResult = "";
        try
        {
            System.Data.DataSet ds = new System.Data.DataSet();
            AMSWebReference.AMSWebService ws = new AMSWebReference.AMSWebService();
            ds = ws.AMSDBQuery2("update legacydev.pn_testers set ecim_status='" + Status + "',ecim_last_update=sysdate where tester='" + sMachineID + "'", "ecimwebservice");

            //if (null == ds || ds.Tables.Count != 1 || ds.Tables[0].Rows.Count < 1)
            //    sResult = "";
            //else
            //{
            //    sResult = ds.Tables[0].Rows[0]["srcdevice"].ToString();
            //}
        }
        catch (Exception ex)
        {
            sResult = "";
        }
        return sResult;

    }
    [WebMethod]    //20140812 add by Mark for T3 query LOT info by Machine ID 20140815
    public string QueryLotInfoByMchineID(string sMachineID)
    {
        string sResult = "";
        try
        {
            sResult = PubUtil.FindLotInfobyMachineID(sMachineID);
            return sResult;
        }
        catch (Exception ex)
        {
            sResult = "";
        }
        return sResult;

    }
    [WebMethod]    //20150408 add by Mark for any SQL update/delete statement
    public string AMSUpdateDeleteSQLStatement(string SQL)
    {
        string sResult = "";
        try
        {
            System.Data.DataSet ds = new System.Data.DataSet();
            AMSWebReference.AMSWebService ws = new AMSWebReference.AMSWebService();
            ds = ws.AMSDBQuery2(SQL, "ecimwebservice");

            if (null == ds || ds.Tables.Count != 1 || ds.Tables[0].Rows.Count < 1)
                sResult = "";
            else
            {
                sResult = ds.Tables[0].Rows[0]["srcdevice"].ToString();
            }
        }
        catch (Exception ex)
        {
            sResult = ex.ToString();
        }
        return sResult;

    }

    [WebMethod]    //20101110 add by Elton for Auto EDC process 
    public string GetEDCData(int iServiceType, string sPath, string sLotNo, int iReadings, bool bCompare, string sOffSetArith, double dOffSetValue, int iOffSetAfterPoint, string recipeCode)
    {
        try
        {
            System.Data.DataTable dt = PubUtil.GetEDCData(iServiceType, sPath, sLotNo, iReadings, bCompare, sOffSetArith, dOffSetValue, iOffSetAfterPoint, recipeCode);
            System.Data.DataSet ds = new System.Data.DataSet();

            if (dt != null)
            {
                ds.Tables.Add(dt);
                ds.DataSetName = "AutoEDC";
                System.Xml.XmlDataDocument xdd = new System.Xml.XmlDataDocument(ds);
                return xdd.InnerXml;
            }
            return null;
        }
        catch (Exception)
        {
            return null;
        }
    }
    [WebMethod]    //20160428 add by Elton for Purge Big Data by Row Num.
    public int PurgeBigTable(string sTable, string sCondition, int iRowNum)
    {
        try
        {
            PubUtil.PurgeBigTable(sTable, sCondition, iRowNum);
            return 0;
        }
        catch (Exception)
        {
            return 1;
        }
    }

    [WebMethod]    //20160528 add by Elton for Connector Server health checking
    public string HealthChecking(string sLine)
    {
        try
        {
            return PubUtil.HealthChecking(sLine);
        }
        catch (Exception)
        {
            return "Fail!!";
        }
    }

    [WebMethod]    //20160808 add by Elton for Log Wafer Start and Wafer End process time 
    public string callTesting(string sValue)
    {
        try
        {
            return PubUtil.getDateSet2Json(sValue);
        }
        catch (Exception)
        {
            return "Fail!!";
        }
    }

    [WebMethod] //For SWR Mass out test
    public Boolean getIsSWR(string resourceName, String containerName)
    {
        try
        {
            return PubUtil.getIsSWR(resourceName, containerName);
        }
        catch (Exception)
        {
            return false;
        }
    }


    //For Easy Flow UAT. add by chuck on 20180403
    [WebMethod, System.Web.Services.Protocols.SoapRpcMethod]
    public string AMSDML(string SQL, string sRCPGroup)
    {
        string sResult = "";
        PubUtil.WriteLog("EasyFlowLog", "AMSDML", SQL, sRCPGroup, "Start!!");
        try
        {
            System.Data.DataSet ds = new System.Data.DataSet();
            AMSWebReference.AMSWebService ws = new AMSWebReference.AMSWebService();
            ds = ws.AMSDBQuery2(SQL, "ecimwebservice");
            PubUtil.WriteLog("EasyFlowLog", "AMSDML", "Call AMSDBQuery", SQL, "Complete!!");

            if (sRCPGroup.Equals("True"))
            {//do Recipe Group
                PubUtil.WriteLog("EasyFlowLog", "AMSDML", "Recipe Group = ", sRCPGroup, "Start!!");

                PubUtil.WriteLog("EasyFlowLog", "AMSDML", "Recipe Group = ", sRCPGroup, "Complete!!");
            }
        }
        catch (Exception ex)
        {
            sResult = "";
            PubUtil.WriteLog("Exception", "AMSDML", SQL, sRCPGroup, "Exception:" + ex.Message);
        }
        PubUtil.WriteLog("EasyFlowLog", "AMSDML", SQL, sRCPGroup, "Complete!!" + sResult);
        return sResult;
    }

    //add by Jim for add new method 20180628 
    [WebMethod]
    public string eCIMTxn(string type, string value)
    {
        string result = "";
        PubUtil.WriteLog("eCIMWebServiceLog", "eCIMTxn", type, value, "eCIMTxn Start.");
        switch (type)
        {
            case "0001"://method objective
                {
                    string[] aValue = value.Split(',');
                    if (aValue.Length < 2) return "E0001: Input Data not enought";
                    return PubUtil.Get2DSerialRecordData(aValue[0], aValue[1]);
                }
            case "0002": //add by Jim on 20181019 for when run lot, edit/upload recipe, check FDC status
                {
                    string[] aValue = value.Split(',');
                    if (aValue.Length < 4) return "E0002: Input Data not enought! (Line,RSC,RCP,Flag)";
                    return PubUtil.GetFDC_Master(aValue[0], aValue[1], aValue[2], aValue[3]);
                    //Line,RSC,RCP,true
                }

            default:
                break;
        }

        return result;
    }
}

public class ThreadProcessor_checkMassOut
{
    string _LotNo = "";
    string _RSC = "";
    string _Recipe = "";

    public ThreadProcessor_checkMassOut(string sLotNo, string sRSC, string sRecipe)
    {
        _LotNo = sLotNo;
        _RSC = sRSC;
        _Recipe = sRecipe;
    }
    public void service()
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut ThreadProcessor", _LotNo, _RSC, "Start!!");
        PubUtil.checkMassOut(_LotNo, _RSC, _Recipe);
        PubUtil.WriteLog("eCIMWebServiceLog", "checkMassOut ThreadProcessor", _LotNo, _RSC, "Completed.");
    }
}

public class ThreadProcessor_GetProcessWaferCount
{
    string _LotNo = "";
    string _RSC = "";
    string _Recipe = "";

    public ThreadProcessor_GetProcessWaferCount(string sLotNo, string sRSC, string sRecipe)
    {
        _LotNo = sLotNo;
        _RSC = sRSC;
        _Recipe = sRecipe;
    }
    public void service()
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount ThreadProcessor", _LotNo, _RSC, "Start!!");
        PubUtil.GetProcessWaferCount(_LotNo, _RSC, _Recipe);
        PubUtil.WriteLog("eCIMWebServiceLog", "GetProcessWaferCount ThreadProcessor", _LotNo, _RSC, "Completed.");
    }
}

//setWaferProcessCount
public class ThreadProcessor_setWaferProcessCount
{
    string _LotNo = "";
    string _RSC = "";

    public ThreadProcessor_setWaferProcessCount(string sLotNo, string sRSC)
    {
        _LotNo = sLotNo;
        _RSC = sRSC;
    }
    public void service()
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount ThreadProcessor", _LotNo, _RSC, "Start!!");
        PubUtil.setWaferProcessCount(_LotNo, _RSC);
        PubUtil.WriteLog("eCIMWebServiceLog", "setWaferProcessCount ThreadProcessor", _LotNo, _RSC, "Completed.");
    }
}

public class ThreadProcessor_setAmkLotWfrProc
{
    string _LotNo = "";
    string _RSC = "";

    public ThreadProcessor_setAmkLotWfrProc(string sLotNo, string sRSC)
    {
        _LotNo = sLotNo;
        _RSC = sRSC;
    }
    public void service()
    {
        PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc ThreadProcessor", _LotNo, _RSC, "Start!!");
        PubUtil.setAmkLotWfrProc(_LotNo, _RSC);
        PubUtil.WriteLog("eCIMWebServiceLog", "setAmkLotWfrProc ThreadProcessor", _LotNo, _RSC, "Completed.");
    }
}