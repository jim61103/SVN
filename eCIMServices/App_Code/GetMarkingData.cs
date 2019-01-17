using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections.Generic;

/// <summary>
/// GetMarkingData 的摘要描述
/// </summary>
public class GetMarkingData
{
    public GetMarkingData()
    {
        //
        // TODO: 在此加入建構函式的程式碼
        //

    }
    //public static string getMarkingData(string sLotNo, string sRSC, string RunEVN)
    public static string getMarkingData(string sLotNo, string RunEVN)
    {
        string sResult = string.Empty;
        string tempMarkingData = string.Empty;
        string BinNO = string.Empty;
        bool SingleBin = false;
        Dictionary<string, string> WaferList = new Dictionary<string, string>();
        try
        {
            DataSet ds = new DataSet();
            GetMarkingDataWebService.GetMarkingDataV2 ws = new GetMarkingDataWebService.GetMarkingDataV2();
            //ds = ws.GetMarkingDataByInSiteMO_V2(sLotNo, RunEVN);
            ds = ws.GetMarkingDataByInSiteLot_Generator(sLotNo, RunEVN);
            if (null == ds || ds.Tables.Count != 1)
                throw new Exception();
            else if (ds.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows[0]["BinNo"].ToString().Trim().Equals(string.Empty)) SingleBin = true;

                //if (sRSC.Substring(0, 2).Equals("DP") && SingleBin)
                //{
                //    WaferList = GetWaferListbyUstLotId(sLotNo);
                //    if (WaferList.Count < 0) throw new Exception();
                //}

                //if (WaferList.Count > 0)
                //{
                //    foreach (KeyValuePair<string, string> kvp in WaferList)
                //    {
                //        sResult += kvp.Value + ":" + tempMarkingData + "|";
                //    }
                //}
                //else
                //{
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    if (!row["message"].ToString().Trim().ToUpper().Equals("SUCCESS"))
                    { //mark add 20150731 for Marking Data enhance
                        string errormessage = row["message"].ToString();
                        throw new Exception(errormessage);
                    }
                    for (int i = 1; i <= 10; i++)
                    {
                        tempMarkingData += row["MarkingLine" + i + "Top"] + ";";
                    }
                    if (SingleBin)
                        sResult += "BIN1:" + tempMarkingData.Substring(0, tempMarkingData.Length - 1);
                    else
                    {
                        BinNO = Convert.ToInt16(row["BinNo"].ToString().Replace("BIN", "")).ToString();
                        //BinNO = String.Format("{0:00}", Convert.ToInt16(row["BinNo"]));
                        if (!BinNO.Equals(string.Empty)) BinNO = "_BIN" + BinNO;
                        sResult += row["WaferID"] + BinNO + ":" + tempMarkingData.Substring(0, tempMarkingData.Length - 1) + "|";
                        BinNO = string.Empty;
                    }
                    tempMarkingData = string.Empty;
                }
                if (SingleBin)
                    sResult = "N@" + sResult;
                else
                    sResult = "Y@" + sResult.Substring(0, sResult.Length - 1);
                //}
                PubUtil.WriteLog("GetMarkingData", sLotNo, sResult, string.Empty, string.Empty);
                SingleBin = false;
                //WaferList.Clear();
            }
        }
        catch (Exception ex)
        {
            PubUtil.WriteLog("getMarkingData", "", "", "", ex.Message);
            sResult = "";
        }
        return sResult;
    }

    //public static string Get_DPS_BG_Recipe(string sLotNo, string RunEVN)
    //{
    //    string sResult="";
    //    try
    //    {
    //        DataSet ds = new DataSet();
    //        GetMarkingDataWebService.GetMarkingDataV2 ws = new GetMarkingDataWebService.GetMarkingDataV2();
    //        //ds = ws.GetMarkingDataByInSiteMO_V2(sLotNo, RunEVN);
    //        ds = ws.GetMarkingDataByInSiteLot(sLotNo, RunEVN);
    //        if (null == ds || ds.Tables.Count != 1)
    //            throw new Exception();

    //    }
    //    catch (Exception ex)
    //    {
    //        sResult="";
    //    }
    //    return sResult;
    //}
    //==========mark add 20121116001 for DPS LM recipe name==========
    public static string Get_DPS_LM_Recipe(string sLotNo, string RunEVN)
    {
        string sResult = "";
        try
        {
            DataSet ds = new DataSet();
            GetMarkingDataWebService.GetMarkingDataV2 ws = new GetMarkingDataWebService.GetMarkingDataV2();
            //ds = ws.GetMarkingDataByInSiteMO_V2(sLotNo, RunEVN);
            ds = ws.GetMarkingDataByInSiteLot(sLotNo, RunEVN);
            if (null == ds || ds.Tables.Count != 1 || ds.Tables[0].Rows.Count < 1)
                sResult = "";
            else
            {
                sResult = ds.Tables[0].Rows[0][1].ToString();
            }
        }
        catch (Exception)
        {
            sResult = "";
        }
        return sResult;
    }
    //============================ END ==============================
    //==========mark add 20131009 for PPM recipe name==========
    public static string Get_DPS_PPM_Recipe(string sLotNo, string RunEVN)
    {
        string sResult = "";
        try
        {
            DataSet ds = new DataSet();
            GetMarkingDataWebService.GetMarkingDataV2 ws = new GetMarkingDataWebService.GetMarkingDataV2();
            //ds = ws.GetMarkingDataByInSiteMO_V2(sLotNo, RunEVN);
            ds = ws.NewGetMarkingDataByInSiteLot(sLotNo, RunEVN, "Query");
            if (null == ds || ds.Tables.Count != 1 || ds.Tables[0].Rows.Count < 1)
                sResult = "";
            else
            {
                sResult = ds.Tables[0].Rows[0][1].ToString();
            }
        }
        catch (Exception ex)
        {
            sResult = "";
        }
        return sResult;
    }
    //============================ END ==============================
}
