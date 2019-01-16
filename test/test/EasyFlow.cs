using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iEMS_Setting
{    
    class EasyFlow
    {
        private WebReference.WorkflowServiceService EFGP = new WebReference.WorkflowServiceService();

        public void SetEFGPUrl(string URL)
        {
            this.EFGP.Url = URL;
        }

        public string getFormOID(string sFormID)
        {
            return EFGP.findFormOIDsOfProcess(sFormID);
        }

        public string getUnitId(string sUserID)
        {
            string XML = EFGP.fetchOrgUnitOfUserId(sUserID);
            XML = XML.Substring(XML.IndexOf("<id>") + 4, XML.IndexOf("</id>") - 4 - XML.IndexOf("<id>")).Trim();
            return XML;
        }

        public string getFormFieldTemplate(string sFormID)
        {
           return EFGP.getFormFieldTemplate(getFormOID(sFormID));
        }

        public string invokeProcess(string sProcessID,string sUserID,string sUnitID,string sFormID, string sFormFieldValue,string sSubject)
        {            
            return EFGP.invokeProcess(sProcessID, sUserID, sUnitID, sFormID, sFormFieldValue, sSubject);
        }

        public string CombineFDCApproval(string sService,string SQL,FDCMater vo,string sUnitID)
        {

            // //申請人
            sService = setXMLData(sService, "txtUserID", Global.UserID);
            //sService = sService.Replace("defaultValue</txtUserID>", Global.UserID + "</txtUserID>");
            // //發文單位
            sService = setXMLData(sService, "txtUnitID", sUnitID);
            //sService = sService.Replace("defaultValue</txtUnitID>", sUnitID + "</txtUnitID>");
            // //發文日期
            sService = setXMLData(sService, "Textbox_RequestDate", DateTime.UtcNow.AddHours(8).ToString("yyyy-MM-dd HH:mm:ss"));
            //sService = sService.Replace("defaultValue</Textbox_RequestDate>", DateTime.UtcNow.AddHours(8).ToString("yyyy-MM-dd HH:mm:ss") + "</Textbox_RequestDate>");
            // //Prod Line
            sService = setXMLData(sService, "txtProdLine", vo.Division);
            //sService = sService.Replace("defaultValue</txtProdLine>", vo.Division + "</txtProdLine>");
            // //Operation
            sService = setXMLData(sService, "txtOperation", vo.OperCode);
            //sService = sService.Replace("defaultValue</txtOperation>", vo.OperCode + "</txtOperation>");
            // //Equipment Model
            sService = setXMLData(sService, "txtEquipModel", vo.MachModel);
            //sService = sService.Replace("defaultValue</txtEquipModel>",  vo.MachModel+ "</txtEquipModel>");
            // //M/C No
            sService = setXMLData(sService, "txtMachineID", vo.MachId);
            //sService = sService.Replace("defaultValue</txtMachineID>",  vo.MachId+ "</txtMachineID>");
            // //PKG
            sService = setXMLData(sService, "txtPKG", vo.Package);
            //sService = sService.Replace("defaultValue</txtPKG>", vo.Package + "</txtPKG>");
            // //Recipe ID
            sService = setXMLData(sService, "txtRecipeID", vo.RecipeName);
            //sService = sService.Replace("defaultValue</txtRecipeID>", vo.RecipeName + "</txtRecipeID>");
            // //hidProd
            sService = setXMLData(sService, "hidProd", vo.Division.Replace(" ", "_"));
            //sService = sService.Replace("defaultValue</hidProd>", vo.Division.Replace(" ","_") + "</hidProd>");
            // //hdnSQL
            sService = setXMLData(sService, "hdnSQL", SQL);
            //sService = sService.Replace("defaultValue</hdnSQL>", SQL + "</hdnSQL>");
            // hdnSubmitID
            sService = setXMLData(sService, "hdnSubmitID", vo.EditedBy);
            if (vo.EditedBy.Equals("")) sService = setXMLData(sService, "hdnSubmitID", vo.SubmittedBy);

            return sService;
        }

        public string setXMLData(string sService,string field,string Data)
        {
            string TagEnd = "</" + field + ">";
            string Tag = "perDataProId=\"\">";
            int startIndex = sService.IndexOf(field);
            int endIndex = sService.IndexOf(TagEnd ) + TagEnd.Length;

            string OriginalTarget = sService.Substring(startIndex, endIndex - startIndex);
            //<txtPKG id="txtPKG" dataType="java.lang.String" perDataProId=""> </txtPKG>

            string ReplaceData = OriginalTarget.Substring(OriginalTarget.IndexOf(Tag));
            string NewData = Tag + Data + TagEnd;

            sService = sService.Replace(ReplaceData, NewData);

            return sService;
        }

        public string getProcessStatus(string SN)
        {
            //return EFGP.getFormFieldTemplate(getFormOID(sFormID));
            string XML = EFGP.fetchFullProcInstanceWithSerialNo(SN).ToString();
            string[] Msg = XML.Split('\n');
            string state = "";

            if (Msg.Length>0)
            {
                List<Active> act = new List<Active>();
                bool bFlag = false;

                for (int i = 0; i < Msg.Length; i++)
                {
                    if (Msg[i].IndexOf("<com.dsc.nana.services.webservice.ActInstanceInfo>") > -1)
                    {
                        bFlag = true;
                    }
                    else if (Msg[i].IndexOf("</com.dsc.nana.services.webservice.ActInstanceInfo>") > -1)
                    {
                        bFlag = false;
                    }

                    if (bFlag)
                    {
                        Active oAct = new Active();
                        for (int j = i; j < Msg.Length; j++)
                        {
                            Msg[j] = Msg[j].Replace(" ","");
                            if (Msg[j].IndexOf("activityName") > -1)
                            {                                
                                oAct.ActiveName = Msg[j].Substring(Msg[j].IndexOf(">") + 1, Msg[j].IndexOf("</")- (Msg[j].IndexOf(">")+1));
                            }
                            else if (Msg[j].IndexOf("state") > -1)
                            {
                                oAct.ActiveStatus = Msg[j].Substring(Msg[j].IndexOf(">") + 1, Msg[j].IndexOf("</") - (Msg[j].IndexOf(">") + 1));
                            }
                            else if (Msg[j].IndexOf("performerName") > -1)
                            {
                                oAct.PerformName = Msg[j].Substring(Msg[j].IndexOf(">") + 1, Msg[j].IndexOf("</") - (Msg[j].IndexOf(">") + 1));
                                act.Add(oAct);
                                i = j;
                                bFlag = false;
                                break;
                            }
                        }
                    }

                    if (Msg[i].IndexOf("state") > -1)
                    {
                        state = Msg[i].Substring(Msg[i].IndexOf(">") + 1, Msg[i].IndexOf("</")-(Msg[i].IndexOf(">") + 1));
                    }
                }

                for (int i = 0; i < act.Count; i++)
                {
                    if (act[i].ActiveStatus.IndexOf("open") > -1)
                    {
                        return act[i].ActiveName + ": " + act[i].PerformName + " " + act[i].ActiveStatus;
                    }                    
                }
            }

            return state;
        }

        public class Active
        {
            public string ActiveName { get; set; }
            public string ActiveStatus { get; set; }
            public string PerformName { get; set; }

            Active(String n, String s, String p)
            {
                ActiveName = n;
                ActiveStatus = s;
                PerformName = p;
            }

            public Active()
            {

            }

        }

        public void sendTerminateProcessWithSN(string SN,string sUserID,string Comment)
        {
             EFGP.terminatedProcessForSerialNo(SN, sUserID, Comment);
            
        }
    }
}
