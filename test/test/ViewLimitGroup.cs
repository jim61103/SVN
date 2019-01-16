using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace iEMS_Setting
{
    public partial class ViewLimitGroup : Form
    {
        public bool bCopy = false;
        public static string submitted = "Submitted";
        public static string draft = "Draft";
        public static string copy = "Copy";
        public static string register = "Register";
        public static string edit = "Edit";
        public static string Delete = "Deleted";
        public static string Reject = "Rejected";
        public static string Approved = "Approved";

        protected iEMS parent;
        public Dictionary<string, string> MailDic = new Dictionary<string, string>();
        public bool EqpNoSelect = false;
        public bool bSelectCup = false;
        public string status = "";
        public bool view = true;
        public string Division { get; set; }
        public string Operation { get; set; }
        public string Model { get; set; }
        public string Eqpno { get; set; }
        public string Recipe { get; set; }
        public string Package { get; set; }
        public string Lead { get; set; }
        public int Serialno { get; set; }
        public int Revision { get; set; }
        public List<FDCMater> lisFDCMaster = new List<FDCMater>();
        public List<ParameterLimitGroup> lsParameterLG = new List<ParameterLimitGroup>();
        public List<FDCMater> lsRegisterFDCMaster = new List<FDCMater>();
        public List<ParameterLimitGroup> lsRegisterParameter = new List<ParameterLimitGroup>();

        private static List<FDCMater> lsFDCMasterHistory = new List<FDCMater>();
        private static List<ParameterLimitGroup> lsParameterHistory = new List<ParameterLimitGroup>();

        public List<Parameter> lsParameter = new List<Parameter>();
        public List<Parameter> lsCollect = new List<Parameter>();
        public List<string> lsEqpNo = new List<string>();
        public List<eOCAPTemplete> lseCOAPTemplete = new List<eOCAPTemplete>();

        OleDbDataReader reader;
        string str = "";
        public string sMailList = "";


        public ViewLimitGroup()
        {
            InitializeComponent();
        }

        private void ViewLimitGroup_Load(object sender, EventArgs e)
        {
           
        }
        public ViewLimitGroup(List<FDCMater> ls, List<Parameter> lsCollect, string Div, string Operation, string Model, string EqpNo, string Recipe, string pkg, string status, List<string> lsEqpNo, bool view, Form iems)
        {
            this.lisFDCMaster = ls;
            this.Division = Div;
            this.Operation = Operation;
            this.Model = Model;
            this.Eqpno = EqpNo;
            this.Recipe = Recipe;
            this.Package = pkg;
            this.view = view;
            this.lsEqpNo = lsEqpNo;
            this.status = status;
            this.lsCollect = lsCollect;
            InitializeComponent();
            this.Tag = iems;
            Run();
        }

        public ViewLimitGroup(List<FDCMater> ls,bool copy, string Div, string Operation, string Model, string EqpNo, string Recipe, string pkg, string status, List<string> lsEqpNo, bool view, Form iems)
        {
            this.lisFDCMaster = ls;
            this.bCopy = copy;
            this.Division = Div;
            this.Operation = Operation;
            this.Model = Model;
            this.Eqpno = EqpNo;
            this.Recipe = Recipe;
            this.Package = pkg;
            this.lsEqpNo = lsEqpNo;
            this.view = view;
            this.status = status;
            InitializeComponent();
            this.Tag = iems;
            Run();
        }

        public ViewLimitGroup(List<FDCMater> ls, bool copy, string Div, string Operation, string Model, string EqpNo, string Recipe, string pkg, string status, List<string> lsEqpNo, bool view, List<Parameter> lsCollect, Form iems)
        {
            this.lisFDCMaster = ls;
            this.bCopy = copy;
            this.Division = Div;
            this.Operation = Operation;
            this.Model = Model;
            this.Eqpno = EqpNo;
            this.Recipe = Recipe;
            this.Package = pkg;
            this.lsEqpNo = lsEqpNo;
            this.view = view;
            this.lsCollect = lsCollect;
            this.status = status;
            InitializeComponent();
            this.Tag = iems;
            Run();
        }

        public ViewLimitGroup(List<FDCMater> ls, string Div, string Operation, string Model, string EqpNo, string Recipe ,string pkg,string status,bool view, Form iems)
        {//submit approve
            this.lisFDCMaster = ls;
            this.Division = Div;
            this.Operation = Operation;
            this.Model = Model;
            this.Eqpno = EqpNo;
            this.Recipe = Recipe;
            this.Package = pkg;
            this.view = view;
            this.status = status;
            InitializeComponent();
            this.Tag = iems;
            Run();
        }

        public ViewLimitGroup(List<FDCMater> ls, bool Copy,string sDivision, string sOperation, string sEqpModel, string sEqpNo, string status, List<string> lsEqpNo, List<Parameter> lsCollect, Form iems)
        {
            this.lisFDCMaster = ls;
            this.bCopy = Copy;
            this.Division = sDivision;
            this.Operation = sOperation;
            this.Model = sEqpModel;
            this.Eqpno = sEqpNo;
            this.status = status;
            this.lsEqpNo = lsEqpNo;
            this.lsCollect = lsCollect;
            InitializeComponent();
            this.Tag = iems;
            Run();
        }

        public ViewLimitGroup(string sDivision, string sOperation, string sEqpModel, string sEqpNo, string status,List<string> lsEqpNo,List<Parameter> lsCollect, Form iems)
        {
            this.Division = sDivision;
            this.Operation = sOperation;
            this.Model = sEqpModel;
            this.Eqpno = sEqpNo;
            this.status = status;
            this.lsEqpNo = lsEqpNo;
            this.lsCollect = lsCollect;
            InitializeComponent();
            this.Tag = iems;
            Run();
        }

        public void Run()
        {
            try
            {
                switch (status)
                {
                    case "Register":
                        ViewLimit.ReadOnly = false;
                        ViewLimit.AllowUserToAddRows = true;

                        btnRegisterSubmit.Visible = true;
                        btnSaveAsDraft.Visible = true;
                        btnApprove.Visible = false;
                        btnReject.Visible = false;
                        btnDelete.Visible = false;
                        btnCompare.Visible = false;
                        ckUse.Visible = true;
                        ckSecondFlag.Visible = true;
                        ckLotHold.Visible = true;
                        ckMCHold.Visible = true;
                        ckeMail.Visible = true;
                        ckSetMailList.Visible = true;

                        getEocapSettingList();

                        break;

                    case "Submitted":
                        btnApprove.Visible = true;
                        btnReject.Visible = true;
                        btnRegisterSubmit.Visible = false;
                        btnDelete.Visible = false;
                        btnCompare.Visible = true;
                        break;
                    case "Draft":
                        btnApprove.Visible = false;
                        btnReject.Visible = false;
                        btnRegisterSubmit.Visible = true;
                        btnDelete.Visible = false;
                        btnSaveAsDraft.Visible = false;
                        btnCompare.Visible = false;
                        break;
                    case "Deleted":
                        btnApprove.Visible = false;
                        btnReject.Visible = false;
                        btnRegisterSubmit.Visible = false;
                        btnCompare.Visible = false;
                        if (view)
                        {
                            btnDelete.Visible = false;
                        }
                        else
                        {
                            btnDelete.Visible = true;
                        }
                        btnSaveAsDraft.Visible = false;
                        break;
                    case "Edit":                        
                        btnSaveAsDraft.Visible = true;
                        btnRegisterSubmit.Visible = true;
                        btnCopy.Visible = true;
                        ViewLimit.ReadOnly = false;
                        ViewLimit.AllowUserToAddRows = true;
                        ckUse.Visible = true;
                        ckUse.Checked = false;
                        ckSecondFlag.Visible = true;
                        ckSecondFlag.Checked = false;
                        ckLotHold.Visible = true;
                        ckLotHold.Checked = false;
                        ckMCHold.Visible = true;
                        ckMCHold.Checked = false;
                        ckeMail.Visible = true;
                        ckeMail.Checked = false;
                        ckSetMailList.Visible = true;
                        ckSetMailList.Checked = false;
                        btnCompare.Visible = false;
                        getEocapSettingList();
                        break;
                    default:
                        //ViewLimit.ReadOnly = false;//////////////test/////////////
                        btnApprove.Visible = false;
                        btnReject.Visible = false;
                        btnRegisterSubmit.Visible = false;
                        btnDelete.Visible = false;
                        btnSaveAsDraft.Visible = false;
                        break;
                }

                //Marked for don't add space to eqpNo 20161228
                //if (Eqpno != "")
                //{
                //    cboEqpNo.Items.Add("");
                //}

                cboEqpNo.Items.Add(Eqpno);
                cboEqpNo.Text = Eqpno;
                txtRecipe.Text = Recipe;


                foreach (var n in lsEqpNo)
                {
                    if (!cboEqpNo.Items.Contains(n))
                    {
                        cboEqpNo.Items.Add(n);
                    }
                }

                var tmp = from n in lisFDCMaster
                          where n.Division.Equals(Division) && Operation.IndexOf(n.OperCode) >-1 && n.MachModel.Equals(Model) && n.MachId.Equals(Eqpno) && n.RecipeName.Equals(Recipe) && n.Status.Equals(status)
                          select n;

                string s = "";
                if (lisFDCMaster.Count == 1)
                {
                    s = lisFDCMaster[0].Status;
                }
                else//for register
                {
                    s = status;
                }
                
               
                if (status.Equals("Edit") || status.Equals(Delete) || bCopy || s.Equals(Reject))
                {
                    tmp = from n in lisFDCMaster
                          where n.Division.Equals(Division) && Operation.IndexOf(n.OperCode) >-1 && n.MachModel.Equals(Model) && n.MachId.Equals(Eqpno) && n.RecipeName.Equals(Recipe) && n.Package.Equals(Package)
                          select n;
                }

                foreach (var n in tmp)
                {
                    txtPackage.Text = n.Package;
                    txtLead.Text = n.Lead;
                    this.Serialno = n.Serialno;
                    this.Revision = n.Revision; //20180621 add by Jim for getParameter use
                    break;
                }
                //if (status == "Deleted")
                //{
                //if (lisFDCMaster.Count > 0)
                //{
                //    this.Serialno = lisFDCMaster[0].Serialno;
                //}

                //}


                if (Operation.IndexOf("PLAT") > -1 && Model.ToUpper().StartsWith("RAIDER"))
                {
                    lbCup.Visible = true;
                    cboCup.Visible = true;
                }

                if (Eqpno.Equals(""))
                {
                    lbState.Text = "Coverge　　　　　Owner : " + Division + " / " + Operation + " / " + Model + " / NULL";
                }
                else
                {
                    lbState.Text = "Coverge　　　　　Owner : " + Division + " / " + Operation + " / " + Model + " / " + Eqpno;
                }

                

                if (status.Equals("Register") || status.Equals("Edit"))//Register used
                {

                    cboCup.Enabled = true;
                    cboEqpNo.Enabled = true;
                    txtPackage.Enabled = true;
                    txtRecipe.Enabled = true;
                    txtLead.Enabled = true;
                    //if (status.Equals("Register"))
                    //{
                    //    fillGridView(lsCollect);
                    //}
                    //else
                    //{
                    getParameter();
                    if (Operation.IndexOf("PLAT") > -1 && Model.ToUpper().StartsWith("RAIDER"))
                    {

                        List<string> lscup = (from n in lsParameterLG
                                              where n.ParameterName.StartsWith("PS ")
                                              select n.ParameterName).ToList();
                        cboCup.Items.Add(" ");
                        foreach (var n in lscup)
                        {

                            if (!cboCup.Items.Contains(n.Substring(0, 5)))
                            {
                                cboCup.Items.Add(n.Substring(0, 5));
                            }
                        }

                        lscup = (from n in lsCollect
                                 where n.ParamName.StartsWith("PS ")
                                 select n.ParamName).ToList();

                        foreach (var n in lscup)
                        {
                            if (!cboCup.Items.Contains(n.Substring(0, 5)))
                            {
                                cboCup.Items.Add(n.Substring(0, 5));
                            }
                        }
                    }
                    if (status.Equals("Register"))
                    {
                        if (bCopy)
                        {
                            fillGridView(lsCollect, lsParameterLG);
                            //bCopy = false;
                        }
                        else
                        {
                            fillGridView(lsCollect);
                        }

                    }
                    else
                    {
                        bool CanRequest = true;
                        if (Division.IndexOf("T1")>-1)
                        {
                            if (!Global.passEFGP)//by pass EFGP
                            {
                                EasyFlow EF = new EasyFlow();
                                //20180214 for send EF     
                                if (Global.isUAT)//change EFGP environment 20180213
                                {
                                    EF.SetEFGPUrl(Global.EFGP_UAT);
                                }
                                else
                                {
                                    EF.SetEFGPUrl(Global.EFGP_PROD);
                                }

                                if (lisFDCMaster[0].EFGPSN != null)//(EF.getProcessStatus(vo.EFGPSN).IndexOf("close") > -1)//(vo.EFGPSN.Trim().Equals(""))//needs request EFGP
                                {//already request                                
                                    if (!lisFDCMaster[0].EFGPSN.Equals(""))
                                    {
                                        CanRequest = false;
                                        if ((EF.getProcessStatus(lisFDCMaster[0].EFGPSN).IndexOf("close") > -1))
                                        {
                                            CanRequest = true;
                                        }
                                    }
                                }
                            }                            
                        }

                        if (CanRequest)
                        {
                            //Edit 20180214 deep copy list     
                            lsFDCMasterHistory = getMasterHistory(lisFDCMaster);
                            lsParameterHistory = getParameterHistory(lsParameterLG);
                            fillGridView(lsCollect, lsParameterLG);
                        }
                        else
                        {
                            //show message for wait approval
                            MessageBox.Show("This Setting already send to Approval,\n Please wait for Approval Finish", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {//Other used
                 //lbState.Text = "Owner : " + Division + "/" + Operation + "/" + Model + "/" + Eqpno;
                    cboCup.Enabled = false;
                    cboEqpNo.Enabled = false;
                    txtPackage.Enabled = false;
                    txtLead.Enabled = false;

                    getParameter();

                    //add for get his data 
                    lsFDCMasterHistory = GetMasterHis();
                    lsParameterHistory = GetParameterHis();

                    fillGridView();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private List<ParameterLimitGroup> getParameterHistory(List<ParameterLimitGroup> lsParameterLG)
        {
            List<ParameterLimitGroup> lsParam = new List<ParameterLimitGroup>();
            for (int i = 0; i < lsParameterLG.Count; i++)
            {
                ParameterLimitGroup Param = (ParameterLimitGroup)lsParameterLG[i].Clone();
                lsParam.Add(Param);
            }
            return lsParam;
        }

        private List<FDCMater> getMasterHistory(List<FDCMater> lisFDCMaster)
        {
            List<FDCMater> lsMaster = new List<FDCMater>();
            for (int i = 0; i < lisFDCMaster.Count; i++)
            {
                FDCMater Master = (FDCMater)lisFDCMaster[i].Clone();

                lsMaster.Add(Master);
            }
            return lsMaster;
        }

        private void fillGridView(List<Parameter> lsCollect, List<ParameterLimitGroup> lsParameterLG)
        {
            MailDic.Clear();
            
            
            DataTable View = new DataTable("View");
            View.Columns.Add("Parameter name", typeof(string));
            View.Columns.Add("Use", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Second flag", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Step name", typeof(string));
            View.Columns.Add("Step value", typeof(string));
            View.Columns.Add("Event id", typeof(int));
            View.Columns.Add("Warm up time", typeof(int));
            View.Columns.Add("Monitor time", typeof(int));
            View.Columns.Add("Start detect value", typeof(double));
            View.Columns.Add("Unit", typeof(string));
            View.Columns.Add("Data type", typeof(string));
            View.Columns.Add("Use lower limit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Permit equal lower", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Lower limit", typeof(double));
            View.Columns.Add("Use upper limit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Permit equal upper", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Upper limit", typeof(double));
            View.Columns.Add("Rchart flag", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Use Rchart lower limit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Rchart lower limit", typeof(double));
            View.Columns.Add("Use Rchart upper limit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Rchart upper limit", typeof(double));
            View.Columns.Add("Terminal msg", System.Type.GetType("System.Boolean"));
            View.Columns.Add("eMail", System.Type.GetType("System.Boolean"));
            View.Columns.Add("M/C inhibit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("M/C hold", System.Type.GetType("System.Boolean"));
            View.Columns.Add("eOCAP", System.Type.GetType("System.Boolean"));
            //View.Columns.Add("eOCAP Templete",typeof(string));
            View.Columns.Add("Lot hold", System.Type.GetType("System.Boolean"));

            if (!cboCup.Text.Trim().Equals(""))
            {
                bSelectCup = true;
            }
            else
            {
                bSelectCup = false;
            }

            foreach (var n in lsCollect)
            {
                
                bool Find = false;
                for (int i = 0; i < lsParameterLG.Count; i++)
                {
                    if (lsParameterLG[i].ParameterName.Equals(n.ParamName))
                    {
                        Find = true;
                        break;
                    }                   
                }

                if (!Find)
                {
                    if (bSelectCup)
                    {
                        if (n.ParamName.StartsWith(cboCup.Text))
                        {
                            View.Rows.Add(new object[] { n.ParamName, false, false, "", "", null, null,null, null, n.ParamUnit, n.ParamType, false, false, null, false, false, null, false, false, null, false, null, false, false, false, false, false, false });
                        }
                    }
                    else
                    {
                        View.Rows.Add(new object[] { n.ParamName, false, false, "", "", null, null,null, null, n.ParamUnit, n.ParamType, false, false, null, false, false, null, false, false, null, false, null, false, false, false, false, false, false });
                    }                    
                }
            }

            foreach (var n in lsParameterLG)
            {
                bool secondflag = false;
                bool checkmin = false;
                bool AllowMinEqual = false;
                bool checkmax = false;
                bool AllowMaxEqual = false;
                bool RChartFlag = false;
                bool CheckRChartMin = false;
                bool CheckRChartMax = false;
                bool TerminalMsg = false;
                bool Email = false;
                bool McInhibit = false;
                bool McHold = false;
                bool Eocap = false;
                bool LotHold = false;

                if (n.SecondFlag == 1) secondflag = true;
                if (n.CheckMin == 1) checkmin = true;
                if (n.AllowMinEqual == 1) AllowMinEqual = true;
                if (n.CheckMax == 1) checkmax = true;
                if (n.AllowMaxEqual == 1) AllowMaxEqual = true;
                if (n.RChartFlag == 1) RChartFlag = true;
                if (n.CheckRMin == 1) CheckRChartMin = true;
                if (n.CheckRMax == 1) CheckRChartMax = true;
                if (n.TerminalMsg == 1) TerminalMsg = true;
                if (n.eMail == 1) Email = true;
                if (n.MCInhibit == 1) McInhibit = true;
                if (n.MCHold == 1) McHold = true;
                if (n.eOCAP == 1) Eocap = true;
                if (n.LotHold == 1) LotHold = true;
                MailDic.Add(n.ParameterName, n.eMailAddrs);
                if (bSelectCup)
                {
                    if (n.ParameterName.StartsWith(cboCup.Text))
                    {
                        View.Rows.Add(new object[] { n.ParameterName, true, secondflag, n.StepName, n.StepValue, n.EventId, n.WarmUpTime,n.MonitorTime, n.StartDetectValue, "", "", checkmin, AllowMinEqual, n.MinValue, checkmax, AllowMaxEqual, n.MaxValue, RChartFlag, CheckRChartMin, n.RMinValue, CheckRChartMax, n.RMaxValue, TerminalMsg, Email, McInhibit, McHold, Eocap, LotHold });
                    }
                }
                else
                {
                    View.Rows.Add(new object[] { n.ParameterName, true, secondflag, n.StepName, n.StepValue, n.EventId, n.WarmUpTime,n.MonitorTime, n.StartDetectValue, "", "", checkmin, AllowMinEqual, n.MinValue, checkmax, AllowMaxEqual, n.MaxValue, RChartFlag, CheckRChartMin, n.RMinValue, CheckRChartMax, n.RMaxValue, TerminalMsg, Email, McInhibit, McHold, Eocap, LotHold });
                }

                //View.Rows.Add(new object[] { n.ParameterName, true, secondflag, n.StepName, n.StepValue, n.EventId, n.WarmUpTime, n.StartDetectValue, "", "", checkmin, AllowMinEqual, n.MinValue, checkmax, AllowMaxEqual, n.MaxValue, RChartFlag, CheckRChartMin, n.RMinValue, CheckRChartMax, n.RMaxValue, TerminalMsg, Email, McInhibit, McHold, Eocap, LotHold });
            }

            foreach (var n in lsCollect) // 20170419 add for set new parameter email default value
            {
                if (!MailDic.ContainsKey(n.ParamName))
                {
                    MailDic.Add(n.ParamName, "");
                }
            }

            ViewLimit.MultiSelect = false;
            ViewLimit.DataSource = View;
            if (ViewLimit.Columns["eMail List"] == null)
            {
                 DataGridViewColumn btn = new DataGridViewButtonColumn();
                btn.Name = "eMail List";
                ViewLimit.Columns.Insert(24, btn);

                DataGridViewComboBoxColumn cbo = new DataGridViewComboBoxColumn();
                cbo.Name = "eOCAP Templete";
                cbo.HeaderText = "eOCAP Templete";


                List<string> lseOCAPtmp = new List<string>();
                lseCOAPTemplete.ForEach(n => lseOCAPtmp.Add(n.Templete));
                if (lseOCAPtmp.Count.Equals(0))
                {
                    cbo.Items.Add(" ");
                    lseOCAPtmp.Add(" ");
                }
                else
                {
                    for (int i = 0; i < lseOCAPtmp.Count; i++)
                    {
                        if (!lseOCAPtmp.Contains(lsParameterLG[i].eOCAPTemplete))
                        {
                            if (lsParameterLG[i].eOCAPTemplete.Equals(""))
                            {
                                cbo.Items.Add("");
                                lseOCAPtmp.Add("");
                            }
                            else
                            {
                                cbo.Items.Add(lsParameterLG[i].eOCAPTemplete.ToString());
                                lseOCAPtmp.Add(lsParameterLG[i].eOCAPTemplete.ToString());
                            }
                        }
                    }
                }

                cbo.DataSource = lseOCAPtmp;
                ViewLimit.Columns.Insert(28, cbo);

                for (int i = 0; i < ViewLimit.RowCount; i++)
                {
                    ViewLimit.Rows[i].Cells[28].Value = " ";
                }
            }
          
       
            ViewLimit.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ViewLimit.Columns[0].Frozen = true;
            //this.ViewLimit.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //ViewLimit.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void getEocapSettingList()
        {
            lseCOAPTemplete.Clear();
            DBConnect con = new DBConnect();
            try
            {
                str = "select * from ATT_ECIM.eOCAP where Pkg='" + Division + "' and OperCode='"+ Operation +"'";

                reader = con.queryforResult(str, con.getConnect());

                while (reader.Read())
                {
                    setEocap(reader);
                }
                con.closeData();
            }
            catch (Exception ex)
            {
                con.closeData();
            }
        }

        private void setEocap(OleDbDataReader reader)
        {
            eOCAPTemplete eCOAPTemplete = new eOCAPTemplete
            {
                Pkg = reader["Pkg"].ToString(),
                OperCode = reader["OperCode"].ToString(),
                Templete = reader["Templete"].ToString(),
                CreatedTime = reader["CreatedTime"].ToString(),
                CreateBy = reader["CreatedBy"].ToString()
            };
            lseCOAPTemplete.Add(eCOAPTemplete);
        }

        private void fillGridView(List<Parameter> lsCollect)
        {
            DataTable View = new DataTable("View");
            View.Columns.Add("Parameter name", typeof(string));
            View.Columns.Add("Use", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Second flag", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Step name", typeof(string));
            View.Columns.Add("Step value", typeof(string));
            View.Columns.Add("Event id", typeof(int));
            View.Columns.Add("Warm up time", typeof(int));
            View.Columns.Add("Monitor time", typeof(int));
            View.Columns.Add("Start detect value", typeof(double));
            View.Columns.Add("Unit", typeof(string));
            View.Columns.Add("Data type", typeof(string));
            View.Columns.Add("Use lower limit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Permit equal lower", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Lower limit", typeof(double));
            View.Columns.Add("Use upper limit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Permit equal upper", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Upper limit", typeof(double));
            View.Columns.Add("Rchart flag", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Use Rchart lower limit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Rchart lower limit", typeof(double));
            View.Columns.Add("Use Rchart upper limit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("Rchart upper limit", typeof(double));
            View.Columns.Add("Terminal msg", System.Type.GetType("System.Boolean"));
            View.Columns.Add("eMail", System.Type.GetType("System.Boolean"));
            View.Columns.Add("M/C inhibit", System.Type.GetType("System.Boolean"));
            View.Columns.Add("M/C hold", System.Type.GetType("System.Boolean"));
            View.Columns.Add("eOCAP", System.Type.GetType("System.Boolean"));
            //View.Columns.Add("eOCAP Templete",typeof(string));
            View.Columns.Add("Lot hold", System.Type.GetType("System.Boolean"));

            if (!cboCup.Text.Trim().Equals(""))
            {
                bSelectCup = true;
            }
            else
            {
                bSelectCup = false;
            }

            foreach (var n in lsCollect)
            {
                //20171207 change eMail , M/C Hold, Lot Hold default check.
                if (bSelectCup)
                {
                    if (n.ParamName.StartsWith(cboCup.Text))
                    {
                        View.Rows.Add(new object[] { n.ParamName, true, true, "", "", null, null,null, null, n.ParamUnit, n.ParamType, false, false, null, false, false, null, false, false, null, false, null, false, true, false, true, false, true });
                    }
                }
                else
                {
                    View.Rows.Add(new object[] { n.ParamName, true, true, "", "", null, null,null, null, n.ParamUnit, n.ParamType, false, false, null, false, false, null, false, false, null, false, null, false, true, false, true, false, true });
                }
                //View.Rows.Add(new object[] { n.ParamName, true, true,"","",null,null, null, n.ParamUnit, n.ParamType, false, false, null, false, false,null, false, false,null, false, null, false, false, false, false, false, false });
            }
            ViewLimit.MultiSelect = false;
            ViewLimit.DataSource = View;
            DataGridViewColumn btn = new DataGridViewButtonColumn();
            btn.Name = "eMail List";
            ViewLimit.Columns.Insert(24, btn);

            DataGridViewComboBoxColumn cbo = new DataGridViewComboBoxColumn();
            cbo.Name = "eOCAP Templete";
            cbo.HeaderText = "eOCAP Templete";


            List<string> lseOCAPtmp = new List<string>();
            lseCOAPTemplete.ForEach(n => lseOCAPtmp.Add(n.Templete));
            if (lseOCAPtmp.Count.Equals(0))
            {
                cbo.Items.Add(" ");
                lseOCAPtmp.Add(" ");
            }
            else
            {
                for (int i = 0; i < lseOCAPtmp.Count; i++)
                {
                    if (!lseOCAPtmp.Contains(lsParameterLG[i].eOCAPTemplete))
                    {
                        if (lsParameterLG[i].eOCAPTemplete.Equals(""))
                        {
                            cbo.Items.Add("");
                            lseOCAPtmp.Add("");
                        }
                        else
                        {
                            cbo.Items.Add(lsParameterLG[i].eOCAPTemplete.ToString());
                            lseOCAPtmp.Add(lsParameterLG[i].eOCAPTemplete.ToString());
                        }
                    }
                }
            }
           
            cbo.DataSource = lseOCAPtmp;
            ViewLimit.Columns.Insert(27, cbo);
            for (int i = 0; i < ViewLimit.RowCount; i++)
            {
                ViewLimit.Rows[i].Cells[27].Value = " ";
            }
            ViewLimit.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ViewLimit.Columns[0].Frozen = true;
            //this.ViewLimit.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //ViewLimit.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

        }

        private void fillGridView()
        {
            try
            {
                DataTable View = new DataTable("View");
                View.Columns.Add("Parameter name", typeof(string));
                View.Columns.Add("Use", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Second flag", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Step name", typeof(string));
                View.Columns.Add("Step value", typeof(string));
                View.Columns.Add("Event id", typeof(int));
                View.Columns.Add("Warm up time", typeof(int));
                View.Columns.Add("Monitor time", typeof(int));//20180214 add for EE1 request
                View.Columns.Add("Start detect value", typeof(double));
                View.Columns.Add("Unit", typeof(string));
                View.Columns.Add("Data type", typeof(string));
                View.Columns.Add("Use lower limit", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Permit equal lower", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Lower limit", typeof(double));
                View.Columns.Add("Use upper limit", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Permit equal upper", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Upper limit", typeof(double));
                View.Columns.Add("Rchart flag", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Use Rchart lower limit", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Rchart lower limit", typeof(double));
                View.Columns.Add("Use Rchart upper limit", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Rchart upper limit", typeof(double));
                View.Columns.Add("Terminal msg", System.Type.GetType("System.Boolean"));
                View.Columns.Add("eMail", System.Type.GetType("System.Boolean"));
                View.Columns.Add("M/C inhibit", System.Type.GetType("System.Boolean"));
                View.Columns.Add("M/C hold", System.Type.GetType("System.Boolean"));
                View.Columns.Add("eOCAP", System.Type.GetType("System.Boolean"));
                //View.Columns.Add("eOCAP Templete", typeof(ComboBox));
                View.Columns.Add("Lot hold", System.Type.GetType("System.Boolean"));

                MailDic.Clear();
                foreach (var n in lsParameterLG)
                {
                    bool secondflag = false;
                    bool checkmin = false;
                    bool AllowMinEqual = false;
                    bool checkmax = false;
                    bool AllowMaxEqual = false;
                    bool RChartFlag = false;
                    bool CheckRChartMin = false;
                    bool CheckRChartMax = false;
                    bool TerminalMsg = false;
                    bool Email = false;
                    bool McInhibit = false;
                    bool McHold = false;
                    bool Eocap = false;
                    bool LotHold = false;

                    if (n.SecondFlag == 1) secondflag = true;
                    if (n.CheckMin == 1) checkmin = true;
                    if (n.AllowMinEqual == 1) AllowMinEqual = true;
                    if (n.CheckMax == 1) checkmax = true;
                    if (n.AllowMaxEqual == 1) AllowMaxEqual = true;
                    if (n.RChartFlag == 1) RChartFlag = true;
                    if (n.CheckRMin == 1) CheckRChartMin = true;
                    if (n.CheckRMax == 1) CheckRChartMax = true;
                    if (n.TerminalMsg == 1) TerminalMsg = true;
                    if (n.eMail == 1) Email = true;
                    if (n.MCInhibit == 1) McInhibit = true;
                    if (n.MCHold == 1) McHold = true;
                    if (n.eOCAP == 1) Eocap = true;
                    if (n.LotHold == 1) LotHold = true;
                    MailDic.Add(n.ParameterName, n.eMailAddrs);
                    View.Rows.Add(new object[] { n.ParameterName, true, secondflag, n.StepName, n.StepValue, n.EventId, n.WarmUpTime,n.MonitorTime, n.StartDetectValue, "", "", checkmin, AllowMinEqual, n.MinValue, checkmax, AllowMaxEqual, n.MaxValue, RChartFlag, CheckRChartMin, n.RMinValue, CheckRChartMax, n.RMaxValue, TerminalMsg, Email, McInhibit, McHold, Eocap, LotHold });

                }
                ViewLimit.MultiSelect = false;
                ViewLimit.DataSource = View;
                DataGridViewColumn btn = new DataGridViewButtonColumn();
                btn.Name = "eMail List";
                ViewLimit.Columns.Insert(24, btn);

                DataGridViewComboBoxColumn cbo = new DataGridViewComboBoxColumn();
                cbo.Name = "eOCAP Templete";
                cbo.HeaderText = "eOCAP Templete";
                

                List<string> lseOCAPtmp = new List<string>();
                for (int i = 0; i < lsParameterLG.Count; i++)
                {
                    if (!lseOCAPtmp.Contains(lsParameterLG[i].eOCAPTemplete))
                    {
                        if (lsParameterLG[i].eOCAPTemplete.Equals(""))
                        {
                            cbo.Items.Add("");
                            lseOCAPtmp.Add("");
                        }
                        else
                        {
                            cbo.Items.Add(lsParameterLG[i].eOCAPTemplete.ToString());
                            lseOCAPtmp.Add(lsParameterLG[i].eOCAPTemplete.ToString());
                        }     
                    }
                }
                cbo.DataSource = lseOCAPtmp;
                ViewLimit.Columns.Insert(27, cbo);


                this.ViewLimit.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                ViewLimit.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                ViewLimit.Columns[0].Frozen = true;
                //ViewLimit.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                //for (int i = 0; i < ViewLimit.ColumnCount - 1; i++)
                //{
                //    this.ViewLimit.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //}

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }           
        }

        private void getParameter()
        {
            try
            {
                DBConnect DBConnect = new DBConnect();
                str = "select * from ATT_ECIM.LimitDefinitionParameter where SerialNo = '" + Serialno + "' and revision ='" + Revision + "' order by parametername";
                reader = DBConnect.queryforResult(str, DBConnect.getConnect());
                while (reader.Read())
                {
                    ParameterLimitGroup PLG = new ParameterLimitGroup { };
                    if (reader["SerialNo"].ToString() != "") PLG.SerialNo = Convert.ToInt32(reader["SerialNo"]);
                    if (reader["ParameterName"].ToString() != "") PLG.ParameterName = reader["ParameterName"].ToString();
                    if (reader["Revision"].ToString() != "") PLG.Revision = Convert.ToInt32(reader["Revision"]);
                    if (reader["CheckMin"].ToString() != "") PLG.CheckMin = Convert.ToByte(reader["CheckMin"]);
                    if (reader["AllowMinEqual"].ToString() != "") PLG.AllowMinEqual = Convert.ToByte(reader["AllowMinEqual"]);                 
                    if (reader["MinValue"].ToString() != "") PLG.MinValue = Convert.ToDouble(reader["MinValue"]);
                    if (reader["CheckMax"].ToString() != "") PLG.CheckMax = Convert.ToByte(reader["CheckMax"]);
                    if (reader["AllowMaxEqual"].ToString() != "") PLG.AllowMaxEqual = Convert.ToByte(reader["AllowMaxEqual"]);
                    if (reader["MaxValue"].ToString() != "") PLG.MaxValue = Convert.ToDouble(reader["MaxValue"]);
                    if (reader["TerminalMsg"].ToString() != "") PLG.TerminalMsg = Convert.ToByte(reader["TerminalMsg"]);
                    if (reader["eMail"].ToString() != "") PLG.eMail = Convert.ToByte(reader["eMail"]);
                    if (reader["eMailAddrs"].ToString() != "") PLG.eMailAddrs = reader["eMailAddrs"].ToString();
                    if (reader["MCInhibit"].ToString() != "") PLG.MCInhibit = Convert.ToByte(reader["MCInhibit"]);
                    if (reader["MCHold"].ToString() != "") PLG.MCHold = Convert.ToByte(reader["MCHold"]);
                    if (reader["eOCAP"].ToString() != "") PLG.eOCAP = Convert.ToByte(reader["eOCAP"]);
                    if (reader["eOCAPTemplete"].ToString() != "") PLG.eOCAPTemplete = reader["eOCAPTemplete"].ToString();
                    if (reader["LotHold"].ToString() != "") PLG.LotHold = Convert.ToByte(reader["LotHold"]);
                    if (reader["StepName"].ToString() != "") PLG.StepName = reader["StepName"].ToString();
                    if (reader["StepValue"].ToString() != "") PLG.StepValue = reader["StepValue"].ToString();
                    if (reader["EventId"].ToString() != "") PLG.EventId = Convert.ToInt32(reader["EventId"]);
                    if (reader["WarmUpTime"].ToString() != "") PLG.WarmUpTime = Convert.ToInt32(reader["WarmUpTime"]);
                    if (reader["MonitorTime"].ToString() != "") PLG.MonitorTime = Convert.ToInt32(reader["MonitorTime"]);
                    if (reader["StartDetectValue"].ToString() != "") PLG.StartDetectValue = Convert.ToDouble(reader["StartDetectValue"]);
                    if (reader["SecondFlag"].ToString() != "") PLG.SecondFlag = Convert.ToByte(reader["SecondFlag"]);
                    if (reader["CheckRMin"].ToString() != "") PLG.CheckRMin = Convert.ToByte(reader["CheckRMin"]);
                    if (reader["RMinValue"].ToString() != "") PLG.RMinValue = Convert.ToDouble(reader["RMinValue"]);
                    if (reader["CheckRMax"].ToString() != "") PLG.CheckRMax = Convert.ToByte(reader["CheckRMax"]);
                    if (reader["RMaxValue"].ToString() != "") PLG.RMaxValue = Convert.ToDouble(reader["RMaxValue"]);
                    if (reader["RChartFlag"].ToString() != "") PLG.RChartFlag = Convert.ToByte(reader["RChartFlag"]);
                    
                    lsParameterLG.Add(PLG);
                }
                //return lsParameterLG;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                //return lsParameterLG;
            }
        }

        private List<ParameterLimitGroup> GetParameterHis()
        {
            try
            {
                List<ParameterLimitGroup> lsParameter = new List<ParameterLimitGroup>();
                DBConnect DBConnect = new DBConnect();
                str = "select * from ATT_ECIM.LimitDefinitionParameterHis " +
                    "where SerialNo = '" + Serialno + "' and revision = (select max(revision)" +
                     "from ATT_ECIM.LimitDefinitionParameterHis where SerialNo = '" + Serialno + "')" +
                        " order by parametername";
                reader = DBConnect.queryforResult(str, DBConnect.getConnect());
                while (reader.Read())
                {
                    ParameterLimitGroup PLG = new ParameterLimitGroup { };
                    if (reader["SerialNo"].ToString() != "") PLG.SerialNo = Convert.ToInt32(reader["SerialNo"]);
                    if (reader["ParameterName"].ToString() != "") PLG.ParameterName = reader["ParameterName"].ToString();
                    if (reader["Revision"].ToString() != "") PLG.Revision = Convert.ToInt32(reader["Revision"]);
                    if (reader["CheckMin"].ToString() != "") PLG.CheckMin = Convert.ToByte(reader["CheckMin"]);
                    if (reader["AllowMinEqual"].ToString() != "") PLG.AllowMinEqual = Convert.ToByte(reader["AllowMinEqual"]);
                    if (reader["MinValue"].ToString() != "") PLG.MinValue = Convert.ToDouble(reader["MinValue"]);
                    if (reader["CheckMax"].ToString() != "") PLG.CheckMax = Convert.ToByte(reader["CheckMax"]);
                    if (reader["AllowMaxEqual"].ToString() != "") PLG.AllowMaxEqual = Convert.ToByte(reader["AllowMaxEqual"]);
                    if (reader["MaxValue"].ToString() != "") PLG.MaxValue = Convert.ToDouble(reader["MaxValue"]);
                    if (reader["TerminalMsg"].ToString() != "") PLG.TerminalMsg = Convert.ToByte(reader["TerminalMsg"]);
                    if (reader["eMail"].ToString() != "") PLG.eMail = Convert.ToByte(reader["eMail"]);
                    if (reader["eMailAddrs"].ToString() != "") PLG.eMailAddrs = reader["eMailAddrs"].ToString();
                    if (reader["MCInhibit"].ToString() != "") PLG.MCInhibit = Convert.ToByte(reader["MCInhibit"]);
                    if (reader["MCHold"].ToString() != "") PLG.MCHold = Convert.ToByte(reader["MCHold"]);
                    if (reader["eOCAP"].ToString() != "") PLG.eOCAP = Convert.ToByte(reader["eOCAP"]);
                    if (reader["eOCAPTemplete"].ToString() != "") PLG.eOCAPTemplete = reader["eOCAPTemplete"].ToString();
                    if (reader["LotHold"].ToString() != "") PLG.LotHold = Convert.ToByte(reader["LotHold"]);
                    if (reader["StepName"].ToString() != "") PLG.StepName = reader["StepName"].ToString();
                    if (reader["StepValue"].ToString() != "") PLG.StepValue = reader["StepValue"].ToString();
                    if (reader["EventId"].ToString() != "") PLG.EventId = Convert.ToInt32(reader["EventId"]);
                    if (reader["WarmUpTime"].ToString() != "") PLG.WarmUpTime = Convert.ToInt32(reader["WarmUpTime"]);
                    if (reader["MonitorTime"].ToString() != "") PLG.MonitorTime = Convert.ToInt32(reader["MonitorTime"]);
                    if (reader["StartDetectValue"].ToString() != "") PLG.StartDetectValue = Convert.ToDouble(reader["StartDetectValue"]);
                    if (reader["SecondFlag"].ToString() != "") PLG.SecondFlag = Convert.ToByte(reader["SecondFlag"]);
                    if (reader["CheckRMin"].ToString() != "") PLG.CheckRMin = Convert.ToByte(reader["CheckRMin"]);
                    if (reader["RMinValue"].ToString() != "") PLG.RMinValue = Convert.ToDouble(reader["RMinValue"]);
                    if (reader["CheckRMax"].ToString() != "") PLG.CheckRMax = Convert.ToByte(reader["CheckRMax"]);
                    if (reader["RMaxValue"].ToString() != "") PLG.RMaxValue = Convert.ToDouble(reader["RMaxValue"]);
                    if (reader["RChartFlag"].ToString() != "") PLG.RChartFlag = Convert.ToByte(reader["RChartFlag"]);

                    lsParameter.Add(PLG);
                }
                DBConnect.closeData();
                return lsParameter;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private List<FDCMater> GetMasterHis()
        {
            try
            {
                List<FDCMater> lsMaster = new List<FDCMater>();
                DBConnect DBConnect = new DBConnect();
                str = "select * from ATT_ECIM.LimitDefinitionMasterHistory " +
                    "where SerialNo = '" + Serialno + "' and revision = (select max(revision) " +
                    "from ATT_ECIM.LimitDefinitionMasterHistory where SerialNo = '" + Serialno + "')";

                reader = DBConnect.queryforResult(str, DBConnect.getConnect());
                while (reader.Read())
                {
                    //if (Global.passEFGP)
                    //{
                    //    FDCMater FDCmaster = new FDCMater
                    //    {
                    //        Serialno = Int32.Parse(reader["SERIALNO"].ToString()),
                    //        Revision = Int32.Parse(reader["REVISION"].ToString()),
                    //        Status = reader["STATUS"].ToString(),
                    //        Division = reader["DIVISION"].ToString(),
                    //        OperCode = reader["OPERCODE"].ToString(),
                    //        MachModel = reader["MACHMODEL"].ToString(),
                    //        MachId = reader["MACHID"].ToString(),
                    //        Customer = reader["CUSTOMER"].ToString(),
                    //        Package = reader["PACKAGE"].ToString(),
                    //        Dimension = reader["DIMENSION"].ToString(),

                    //        Lead = reader["LEAD"].ToString(),
                    //        Device = reader["DEVICE"].ToString(),
                    //        BondingNo = reader["BONDINGNO"].ToString(),
                    //        BondingRev = reader["BONDINGREV"].ToString(),
                    //        RecipeName = reader["RECIPENAME"].ToString(),
                    //        CreatedTime = reader["CREATEDTIME"].ToString(),
                    //        CreatedBy = reader["CREATEDBY"].ToString(),
                    //        SubmittedTime = reader["SUBMITTEDTIME"].ToString(),
                    //        SubmittedBy = reader["SUBMITTEDBY"].ToString(),
                    //        ApprovedTime = reader["APPROVEDTIME"].ToString(),

                    //        ApprovedBy = reader["APPROVEDBY"].ToString(),
                    //        EditedTime = reader["EDITEDTIME"].ToString(),
                    //        EditedBy = reader["EDITEDBY"].ToString(),
                    //        Memo1 = reader["MEMO1"].ToString(),
                    //        Memo2 = reader["MEMO2"].ToString(),
                    //        Memo3 = reader["MEMO3"].ToString(),
                    //        //EFGPSN = reader["EFGPSN"].ToString(),
                    //    };
                    //    lsMaster.Add(FDCmaster);
                    //}
                    //else
                    //{
                        FDCMater FDCmaster = new FDCMater
                        {
                            Serialno = Int32.Parse(reader["SERIALNO"].ToString()),
                            Revision = Int32.Parse(reader["REVISION"].ToString()),
                            Status = reader["STATUS"].ToString(),
                            Division = reader["DIVISION"].ToString(),
                            OperCode = reader["OPERCODE"].ToString(),
                            MachModel = reader["MACHMODEL"].ToString(),
                            MachId = reader["MACHID"].ToString(),
                            Customer = reader["CUSTOMER"].ToString(),
                            Package = reader["PACKAGE"].ToString(),
                            Dimension = reader["DIMENSION"].ToString(),

                            Lead = reader["LEAD"].ToString(),
                            Device = reader["DEVICE"].ToString(),
                            BondingNo = reader["BONDINGNO"].ToString(),
                            BondingRev = reader["BONDINGREV"].ToString(),
                            RecipeName = reader["RECIPENAME"].ToString(),
                            CreatedTime = reader["CREATEDTIME"].ToString(),
                            CreatedBy = reader["CREATEDBY"].ToString(),
                            SubmittedTime = reader["SUBMITTEDTIME"].ToString(),
                            SubmittedBy = reader["SUBMITTEDBY"].ToString(),
                            ApprovedTime = reader["APPROVEDTIME"].ToString(),

                            ApprovedBy = reader["APPROVEDBY"].ToString(),
                            EditedTime = reader["EDITEDTIME"].ToString(),
                            EditedBy = reader["EDITEDBY"].ToString(),
                            Memo1 = reader["MEMO1"].ToString(),
                            Memo2 = reader["MEMO2"].ToString(),
                            Memo3 = reader["MEMO3"].ToString(),
                            EFGPSN = reader["EFGPSN"].ToString(),
                        };
                        lsMaster.Add(FDCmaster);
                    }
                    
                //}
                DBConnect.closeData();
                return lsMaster;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void ViewLimit_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            bool ReadOnly;

            switch (status)
            {
                case "Edit":
                    goto case "Register";
                case "Register":
                    ReadOnly = false;
                    break;
                default:
                    ReadOnly = true;
                    break;
            }
            if (e.ColumnIndex == this.ViewLimit.Columns["eMail List"].Index)
            {
                string Parameter = this.ViewLimit.Rows[e.RowIndex].Cells["Parameter name"].Value.ToString();
                if (!Parameter.Trim().Equals(""))
                {
                    string mail = "";

                    foreach (var n in lsParameterLG)
                    {
                        if (n.ParameterName.Equals(Parameter))
                        {
                            mail = n.eMailAddrs;
                        }
                    }

                    if (MailDic.Count > 0 || lsParameterLG.Count == 0)
                    {
                        if (MailDic.ContainsKey(Parameter))
                        {
                            mail = MailDic[Parameter];
                        }
                    }

                    string mailaddress = "";
                    //sMailList = "";
                    ParameterSetting ParameterSetting = new ParameterSetting(Division, Operation, mailaddress, ReadOnly, mail);

                    ParameterSetting.ShowDialog();
                    if (status != submitted)
                    {
                        if (ParameterSetting.sMailList != "")
                        {
                            mailaddress = ParameterSetting.sMailList; // for set all email list
                            if (MailDic.Keys.ToList().Contains(Parameter))
                            {
                                MailDic.Remove(Parameter);
                                MailDic.Add(Parameter, mailaddress); // for set singal email list
                            }
                            else
                            {
                                MailDic.Add(Parameter, ParameterSetting.sMailList); // for set singal email list
                            }
                        }
                        else
                        {
                            mailaddress = "";
                            if (MailDic.Keys.ToList().Contains(Parameter))
                            {
                                MailDic.Remove(Parameter);
                                MailDic.Add(Parameter, mailaddress); // for set singal email list
                            }
                            else
                            {
                                MailDic.Add(Parameter, ParameterSetting.sMailList); // for set singal email list
                            }
                        }
                    }

                }
            }          
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
        }

        private void eMailCheckAll_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < ViewLimit.Rows.Count-1; i++)
            {
                ViewLimit.Rows[i].Cells["Use"].Value = this.ckUse.Checked;
            }

        }

        private void ckSecondFlag_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < ViewLimit.Rows.Count - 1; i++)
            {
                ViewLimit.Rows[i].Cells["Second flag"].Value = this.ckSecondFlag.Checked;
            }

            //foreach (DataGridViewRow row in ViewLimit.Rows)
            //{
            //    row.Cells["Second flag"].Value = this.ckSecondFlag.Checked;
            //}
        }

        private void ckHold_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < ViewLimit.Rows.Count - 1; i++)
            {
                ViewLimit.Rows[i].Cells["Lot hold"].Value = this.ckLotHold.Checked;
            }

            //foreach (DataGridViewRow row in ViewLimit.Rows)
            //{
            //    row.Cells["Lot hold"].Value = this.ckHold.Checked;
            //}
        }

        private int getSerialNo(FDCMater vo)
        {
            int serialNo = -1;
            //ATT_ECIM.LimitDefinitionMaster
            //ATT_ECIM.LimitDefinitionMaster
            DBConnect cn = new DBConnect();
            string sql = "select * from ATT_ECIM.LimitDefinitionMaster where Status<>'Deleted' and Division='"+vo.Division+"' and OperCode='"+vo.OperCode+"' and MachModel='"+vo.MachModel+"' ";
            sql += vo.MachId.Equals("") ? " and MachId is null" : " and MachId='" + vo.MachId + "' ";
            sql += vo.Customer.Equals("") ? " and Customer is null" : " and Customer='" + vo.Customer + "' ";
            sql += vo.Package.Equals("") ? " and Package is null" : " and Package='" + vo.Package + "' ";
            sql += vo.Dimension.Equals("") ? " and Dimension is null" : " and Dimension='" + vo.Dimension + "' ";
            sql += vo.Lead.Equals("") ? " and Lead is null" : " and Lead='" + vo.Lead + "' ";
            sql += vo.Device.Equals("") ? " and Device is null" : " and Device='" + vo.Device + "' ";
            sql += vo.BondingNo.Equals("") ? " and BondingNo is null" : " and BondingNo='" + vo.BondingNo + "' ";
            sql += vo.BondingRev.Equals("") ? " and BondingRev is null" : " and BondingRev='" + vo.BondingRev + "' ";
            sql += vo.RecipeName.Equals("") ? " and RecipeName is null" : " and RecipeName='" + vo.RecipeName + "' ";

            reader = cn.queryforResult(sql, cn.getConnect());
            while (reader.Read())
            {
                serialNo = Int32.Parse(reader["SerialNo"].ToString());
            }

            cn.closeData();
            

            return serialNo;
        }

        private void btnRegisterSubmit_Click(object sender, EventArgs e)
        {            

            DialogResult result = MessageBox.Show("Process submit?", "User confirmation"
                                                , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {               
                if (createLimitItemVos() == true)
                {
                    if (status == edit)
                    {
                        if (!Editsave(false, submitted, lsRegisterFDCMaster, lsRegisterParameter))
                        {
                            return;
                        }
                    }
                    else
                    {
                        if (status ==draft)
                        {// (!Editsave(false, submitted, lsRegisterFDCMaster, lsRegisterParameter))
                            if (!Editsave(false, submitted, lsRegisterFDCMaster, lsRegisterParameter))
                            {
                                return;
                            }
                        }
                        else
                        {
                            if (!save(submitted, lsRegisterFDCMaster, lsRegisterParameter))
                            {
                                return;
                            }
                        }
                       
                    }
                }

                //if (!save("Submitted", lsRegisterFDCMaster, lsRegisterParameter))
                //{
                //    return;
                //}
                   
                
            }                   
        }

        private bool insertRecordToDB(List<FDCMater> lsRegisterFDCMaster, List<ParameterLimitGroup> lsRegisterParameter)
        {
            DBConnect cn = new DBConnect();
            try
            {
                
                if (lsRegisterFDCMaster.Count != 0 && lsRegisterParameter.Count != 0)
                {
                    //master
                    FDCMater voLimitMaster = lsRegisterFDCMaster[0];

                    if (insertRecord(cn, voLimitMaster) != 1)
                    {
                        MessageBox.Show("Submit Fail (insert " + voLimitMaster.RecipeName + ")");
                        cn.closeData();
                        return false;
                    }

                    int intSerial = getSerialNo(voLimitMaster);

                    if (intSerial < 0)
                    {
                        MessageBox.Show("Submit Fail (Exist same record at server.)");
                        cn.closeData();
                        return false;
                    }
                    voLimitMaster.Serialno = intSerial;

                    //parameter
                    for (int i = 0; i < lsRegisterParameter.Count; i++)
                    {
                        ParameterLimitGroup voParam = lsRegisterParameter[i];
                        voParam.SerialNo = intSerial;
                        if (insertParameterRecord(cn, voParam) != 1)
                        {
                            MessageBox.Show("Submit Fail (insert " + voParam.ParameterName + ")");
                            cn.closeData();
                            return false;
                        }
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                cn.closeData();
                return false;
            }
           
        }

        private int insertParameterRecord(OleDbCommand cmd, ParameterLimitGroup voParam)
        {
            try
            {
                int iReturn = -1;
                string sql = "";
                if (ckSetMailList.Checked) // mail list same
                {
                    voParam.eMailAddrs = sMailList;
                }
                else
                {
                    if (voParam.eMail.Equals(1))
                    {
                        voParam.eMailAddrs = MailDic[voParam.ParameterName];
                    }
                    else
                    {
                        voParam.eMailAddrs = "";
                    }
                    
                }

                if (Division.Contains("DPS") || Division.Contains("Flip Chip"))// T3 don't have secod flag 20181129
                {
                    voParam.SecondFlag = 0;
                }
                sql = "insert into ATT_ECIM.LimitDefinitionParameter values ('" + voParam.SerialNo + "','" + voParam.ParameterName + "','" + voParam.Revision + "','" + voParam.CheckMin + "','" + voParam.AllowMinEqual + "', '" + voParam.MinValue + "','" + voParam.CheckMax + "','" + voParam.AllowMaxEqual + "','" + voParam.MaxValue + "','" + voParam.TerminalMsg + "', '" + voParam.eMail + "','" + voParam.eMailAddrs + "','" + voParam.MCInhibit + "','" + voParam.MCHold + "','" + voParam.eOCAP + "', '" + voParam.eOCAPTemplete + "','" + voParam.LotHold + "','" + voParam.StepName + "','" + voParam.StepValue + "','" + voParam.EventId + "', '" + voParam.WarmUpTime + "','" + voParam.StartDetectValue + "','" + voParam.SecondFlag + "','" + voParam.CheckRMin + "','" + voParam.RMinValue + "', '" + voParam.CheckRMax + "','" + voParam.RMaxValue + "','" + voParam.RChartFlag + "','" + voParam.MonitorTime + "')";
                cmd.CommandText = sql;
                iReturn = cmd.ExecuteNonQuery();//cn.insert(sql, cn.getConnect());
                return iReturn;
            }
            catch (Exception ex)
            {
                return -1;
            }
           
        }
        private int insertParameterRecord(DBConnect cn, ParameterLimitGroup voParam)
        {
            try
            {
                int iReturn = -1;
                string sql = "";
                //if (ckSetMailList.Checked) // mail list same
                //{
                //    voParam.eMailAddrs = sMailList;
                //}
                //else
                //{
                //    if (voParam.eMail.Equals(1))
                //    {
                //        voParam.eMailAddrs = MailDic[voParam.ParameterName];
                //    }
                //    else
                //    {
                //        voParam.eMailAddrs = "";
                //    }
                //}

                if (Division.Contains("DPS") || Division.Contains("Flip Chip"))// T3 don't have secod flag 20181129
                {
                    voParam.SecondFlag = 0;
                }
                sql = "insert into ATT_ECIM.LimitDefinitionParameter values ('" + voParam.SerialNo + "','" + voParam.ParameterName + "','" + voParam.Revision + "','" + voParam.CheckMin + "','" + voParam.AllowMinEqual + "', '" + voParam.MinValue + "','" + voParam.CheckMax + "','" + voParam.AllowMaxEqual + "','" + voParam.MaxValue + "','" + voParam.TerminalMsg + "', '" + voParam.eMail + "','" + voParam.eMailAddrs + "','" + voParam.MCInhibit + "','" + voParam.MCHold + "','" + voParam.eOCAP + "', '" + voParam.eOCAPTemplete + "','" + voParam.LotHold + "','" + voParam.StepName + "','" + voParam.StepValue + "','" + voParam.EventId + "', '" + voParam.WarmUpTime + "','" + voParam.StartDetectValue + "','" + voParam.SecondFlag + "','" + voParam.CheckRMin + "','" + voParam.RMinValue + "', '" + voParam.CheckRMax + "','" + voParam.RMaxValue + "','" + voParam.RChartFlag + "','')";
                //cmd.CommandText = sql;
                iReturn = cn.insert(sql, cn.getConnect());
                return iReturn;
            }
            catch (Exception ex)
            {
                return -1;
            }

        }

        private int insertRecord(DBConnect cn, FDCMater vo)
        {
            try
            {
                int iReturn = -1;
                string sql = "";
                if(vo.ApprovedTime != null) vo.ApprovedTime = vo.ApprovedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.ApprovedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.ApprovedTime;
                if (vo.SubmittedTime != null) vo.SubmittedTime = vo.SubmittedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.SubmittedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.SubmittedTime;
                if (vo.EditedTime != null) vo.EditedTime = vo.EditedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.EditedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.EditedTime;
                if (vo.CreatedTime != null) vo.CreatedTime = vo.CreatedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.CreatedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.CreatedTime;
                sql = "insert into ATT_ECIM.LimitDefinitionMaster values (att_ecim.LimitDefinitionMasterid.NEXTVAL , '" + vo.Revision + "','" + vo.Status + "','" + vo.Division + "','" + vo.OperCode + "','" + vo.MachModel + "','" + vo.MachId + "','" + vo.Customer + "','" + vo.Package + "','" + vo.Dimension + "','" + vo.Lead + "','" + vo.Device + "','" + vo.BondingNo + "','" + vo.BondingRev + "','" + vo.RecipeName + "',TO_DATE('" + vo.CreatedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.CreatedBy + "',TO_DATE('" + vo.SubmittedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.SubmittedBy + "',TO_DATE('" + vo.ApprovedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.ApprovedBy + "',TO_DATE('" + vo.EditedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.EditedBy + "','" + vo.Memo1 + "','" + vo.Memo2 + "','" + vo.Memo3 + "','')";
                //if (Global.passEFGP) sql = "insert into ATT_ECIM.LimitDefinitionMaster values (att_ecim.LimitDefinitionMasterid.NEXTVAL , '" + vo.Revision + "','" + vo.Status + "','" + vo.Division + "','" + vo.OperCode + "','" + vo.MachModel + "','" + vo.MachId + "','" + vo.Customer + "','" + vo.Package + "','" + vo.Dimension + "','" + vo.Lead + "','" + vo.Device + "','" + vo.BondingNo + "','" + vo.BondingRev + "','" + vo.RecipeName + "',TO_DATE('" + vo.CreatedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.CreatedBy + "',TO_DATE('" + vo.SubmittedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.SubmittedBy + "',TO_DATE('" + vo.ApprovedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.ApprovedBy + "',TO_DATE('" + vo.EditedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.EditedBy + "','" + vo.Memo1 + "','" + vo.Memo2 + "','" + vo.Memo3 + "')";
                               
                iReturn = cn.insert(sql, cn.getConnect());
                return iReturn;
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        private bool createLimitItemVos()
        {
            Recipe = txtRecipe.Text;
            Package = txtPackage.Text;
            Lead = txtLead.Text;
            //lsRegisterFDCMaster.Clear();
            //lsRegisterParameter.Clear();
            lsRegisterFDCMaster = new List<FDCMater>();
            lsRegisterParameter = new List<ParameterLimitGroup>();
            lsRegisterFDCMaster = setRegisterList();
            if (lsRegisterFDCMaster == null) return false;
            lsRegisterParameter = setParameterList();
            if (lsRegisterParameter == null || lsRegisterParameter.Count==0) return false;
            //if (Recipe.Contains("*"))
            //{
            //    if (Recipe.StartsWith("*") && Recipe.EndsWith("*"))//*123*
            //    {
            //        MessageBox.Show("Recipe Name can only contain 1 *", "User confirmation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return false;
            //    }
            //    else if (!Recipe.StartsWith("*") && !Recipe.EndsWith("*"))//12*34
            //    {
            //        MessageBox.Show("Recipe Name * rule is not correct!", "User confirmation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return false;
            //    }
            //}
            return true;
        }

        private List<ParameterLimitGroup> setParameterList()
        {
            try
            {
                foreach (DataGridViewRow row in ViewLimit.Rows)
                {
                    if (row.Cells["Use"].Value.ToString().Equals("True") && !row.Cells["Parameter name"].Value.ToString().Trim().Equals(""))
                    {
                        ParameterLimitGroup voParam = new ParameterLimitGroup();

                        voParam.ParameterName = row.Cells["Parameter name"].Value.ToString();
                       
                        // Validate Parameter Name
                        bool validate = false;
                        if (bCopy)
                        {
                            validate = true;
                        }
                        else
                        {
                            var tmp = from n in lsCollect
                                      where n.ParamName == voParam.ParameterName
                                      select n;
                            foreach (var n in tmp)
                            {
                                validate = true;
                            }
                        }

                        ///<summary> for not compare parameter setting
                        /// 
                        /// </summary>
                        validate = true;

                        if (!validate)//if parameter name is not in original parameter name return null
                        {
                            MessageBox.Show("Invalid input. Please check again! " + "\n\n " + voParam.ParameterName + "  Parameter Name is not correct", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lsRegisterParameter.Clear();
                            return null;
                        }

                        voParam.Revision = 1;
                        if (!row.Cells["Use lower limit"].Value.ToString().Equals("True") && !row.Cells["Use upper limit"].Value.ToString().Equals("True"))
                        {
                            MessageBox.Show("Invalid input. Please check again! " + "\n\n Parameter Name: " + voParam.ParameterName + ",\n Use upper limit or lower limit","Information for user",MessageBoxButtons.OK,MessageBoxIcon.Information);
                            lsRegisterParameter.Clear();
                            return null;
                        }

                       //Min
                        voParam.CheckMin = row.Cells["Use lower limit"].Value.ToString().Equals("True") ? (byte)1: (byte)0;
                        voParam.AllowMinEqual = row.Cells["Permit equal lower"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        if (!row.Cells["Lower limit"].Value.ToString().Trim().Equals("") && voParam.CheckMin.Equals(1))
                            voParam.MinValue = double.Parse(row.Cells["Lower limit"].Value.ToString());
                        else                        
                            voParam.MinValue = 0;
                        
                        if (row.Cells["Use lower limit"].Value.ToString().Equals("True") && row.Cells["Lower limit"].Value.ToString().Trim().Equals(""))
                        {
                            MessageBox.Show("Invalid input. Please check again! " + "\n\n Parameter Name: " + voParam.ParameterName + ", lower limit", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lsRegisterParameter.Clear();
                            return null;
                        }

                        //Max
                        voParam.CheckMax = row.Cells["Use upper limit"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;
                        voParam.AllowMaxEqual = row.Cells["Permit equal upper"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        if (!row.Cells["Upper limit"].Value.ToString().Trim().Equals("") && voParam.CheckMax.Equals(1))
                            voParam.MaxValue = double.Parse(row.Cells["Upper limit"].Value.ToString());
                        else
                            voParam.MaxValue = 0;

                        if (row.Cells["Use upper limit"].Value.ToString().Equals("True") && row.Cells["Upper limit"].Value.ToString().Trim().Equals(""))
                        {
                            MessageBox.Show("Invalid input. Please check again! " + "\n\n Parameter Name: " + voParam.ParameterName + ", upper limit", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lsRegisterParameter.Clear();
                            return null;
                        }

                        //Validate Min and Max rule
                        if (voParam.CheckMax.Equals(1) && voParam.CheckMin.Equals(1))
                        {
                            if (voParam.MaxValue < voParam.MinValue)
                            {
                                MessageBox.Show("Please check MinValue and MaxValue again!! " + "\n\n Parameter Name: " + voParam.ParameterName + " ", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                lsRegisterParameter.Clear();
                                return null;
                            }
                        }

                        // Termial message
                        voParam.TerminalMsg = row.Cells["Terminal msg"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        // eMail
                        voParam.eMail = row.Cells["eMail"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        //set all mail list same
                        if (ckSetMailList.Checked == true )//&& voParam.eMail.Equals((byte)1))  not check eMail checked or not
                        {
                            if (sMailList.Length<=0)
                            {
                                MessageBox.Show("Invalid input. Please check again! " + "\n\n" + voParam.ParameterName + " eMail List ", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                lsRegisterParameter.Clear();
                                return null;
                            }
                            else
                            {
                                voParam.eMailAddrs = sMailList;
                                if (MailDic.Keys.ToList().Contains(voParam.ParameterName))
                                {
                                    MailDic.Remove(voParam.ParameterName);
                                    MailDic.Add(voParam.ParameterName, sMailList); // for set singal email list
                                }
                                else
                                {
                                    MailDic.Add(voParam.ParameterName, sMailList); // for set singal email list
                                }

                            }
                            
                        }
                        else
                        {
                            try
                            {
                                if (voParam.eMail.Equals(1))
                                {
                                    // eMail Addrs
                                    voParam.eMailAddrs = MailDic[voParam.ParameterName];
                                }                              
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Invalid input. Please check again! " + "\n\n" + voParam.ParameterName + " eMail List ", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return null;
                            }
                               
                        }

                        // MC Inhbit
                        voParam.MCInhibit = row.Cells["M/C inhibit"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        // MC hold
                        voParam.MCHold = row.Cells["M/C hold"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        // eOCAP
                        voParam.eOCAP = row.Cells["eOCAP"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        // eOCAP templete
                        try
                        {
                            voParam.eOCAPTemplete = row.Cells["eOCAP Templete"].Value.ToString();
                        }
                        catch (Exception)
                        {
                            voParam.eOCAPTemplete = "";
                        }
                        

                        // Lot Hold
                        voParam.LotHold = row.Cells["Lot hold"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        //Second flag
                        voParam.SecondFlag = row.Cells["Second flag"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        //step name is not null, check step value is null or not
                        if (!row.Cells["Step name"].Value.ToString().Trim().Equals("") && 
                            (row.Cells["Step value"].Value.ToString().Trim().Equals("") || row.Cells["Step value"].Value.ToString().Trim().Equals("0")))
                        {
                            MessageBox.Show("Invalid input. Please check again !" + "\n\n" + voParam.ParameterName + " \"Step value\" Can not be Null", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lsRegisterParameter.Clear();
                            return null;
                        }

                        //step value is not null, check step name is null or not
                        if (row.Cells["Step name"].Value.ToString().Trim().Equals("") && !row.Cells["Step value"].Value.ToString().Trim().Equals(""))
                        {
                            MessageBox.Show("Invalid input. Please check again !" + "\n\n" + voParam.ParameterName + " \"Step Name\" Can not be Null", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lsRegisterParameter.Clear();
                            return null;
                        }

                        if (!(IsNumeric(row.Cells["Event id"].Value.ToString().Trim()) && IsNumeric(row.Cells["Warm up time"].Value.ToString().Trim()) && IsNumeric(row.Cells["Monitor time"].Value.ToString().Trim()) && IsNumeric(row.Cells["Start detect value"].Value.ToString().Replace(".","").Trim()) ))
                        {                            
                            MessageBox.Show("Invalid input. Please check again !" + "\n\n" + voParam.ParameterName + " Event id, Warm up time ,Monitor time ,Start detect value", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lsRegisterParameter.Clear();
                            return null;
                        }

                        //step name
                        voParam.StepName = !row.Cells["Step name"].Value.ToString().Trim().Equals("") ? row.Cells["Step name"].Value.ToString().Trim() : "";

                        //step value
                        voParam.StepValue = !row.Cells["Step value"].Value.ToString().Trim().Equals("") ? row.Cells["Step value"].Value.ToString().Trim() : "";
                        
                        //event id
                        voParam.EventId = !row.Cells["Event id"].Value.ToString().Trim().Equals("") && !row.Cells["Event id"].Value.ToString().Trim().Equals("") ? Int32.Parse(row.Cells["Event id"].Value.ToString().Trim()) : 0;

                        //warm up time
                        voParam.WarmUpTime = !row.Cells["Warm up time"].Value.ToString().Trim().Equals(0) && !row.Cells["Warm up time"].Value.ToString().Trim().Equals("") ? Int32.Parse(row.Cells["Warm up time"].Value.ToString().Trim()) : 0;

                        //Monitor time
                        voParam.MonitorTime = !row.Cells["Monitor time"].Value.ToString().Trim().Equals(0) && !row.Cells["Monitor time"].Value.ToString().Trim().Equals("") ? Int32.Parse(row.Cells["Monitor time"].Value.ToString().Trim()) : 0;


                        //Start DetectValue
                        voParam.StartDetectValue = !(row.Cells["Start detect value"].Value.ToString().Trim().Equals(0.0) && row.Cells["Start detect value"].Value.ToString().Trim().Equals(0)) && !row.Cells["Start detect value"].Value.ToString().Trim().Equals("") ? Double.Parse(row.Cells["Start detect value"].Value.ToString().Trim()) : 0.0;

                        //RChart Flag 
                        voParam.RChartFlag = row.Cells["Rchart flag"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        //RChart Min
                        voParam.CheckRMin = row.Cells["Use Rchart lower limit"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        //RChart Min Value
                        voParam.RMinValue = !row.Cells["Rchart lower limit"].Value.ToString().Trim().Equals("") ? Double.Parse(row.Cells["Rchart lower limit"].Value.ToString().Trim()) : 0.0;

                        if (row.Cells["Use Rchart lower limit"].Value.ToString().Equals("True") && row.Cells["Rchart lower limit"].Value.ToString().Trim().Equals(""))
                        {
                            MessageBox.Show("Invaild RChart input. Please check again! " + "\n\n" + voParam.ParameterName + " Rchart lower limit", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lsRegisterParameter.Clear();
                            return null;
                        }

                        //RChart Max
                        voParam.CheckRMax = row.Cells["Use Rchart upper limit"].Value.ToString().Equals("True") ? (byte)1 : (byte)0;

                        //RChart Man Value
                        voParam.RMaxValue = !row.Cells["Rchart upper limit"].Value.ToString().Trim().Equals("") ? Double.Parse(row.Cells["Rchart upper limit"].Value.ToString().Trim()) : 0.0;

                        if (row.Cells["Use Rchart upper limit"].Value.ToString().Equals("True") && row.Cells["Rchart upper limit"].Value.ToString().Trim().Equals(""))
                        {
                            MessageBox.Show("Invaild RChart input. Please check again " + "\n\n" + voParam.ParameterName + " Rchart upper limit", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lsRegisterParameter.Clear();
                            return null;
                        }

                        if (row.Cells["Rchart flag"].Value.ToString().Equals("True"))
                        {
                            if (!voParam.CheckRMin.Equals((byte)1) && !voParam.CheckRMax.Equals((byte)1))
                            {
                                MessageBox.Show("Invaild RChart input. Please check again " + "\n\n" + voParam.ParameterName + " Lower limit", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                lsRegisterParameter.Clear();
                                return null;
                            }
                            else
                            {
                                voParam.CheckRMin = (byte) 0;
                                voParam.RMinValue = 0.0;
                                voParam.CheckRMax = (byte)0;
                                voParam.RMaxValue = 0.0;
                            }
                        }

                        lsRegisterParameter.Add(voParam);
                    }

                    if (row.Cells["Use"].Value.ToString().Equals("True") && row.Cells["Parameter name"].Value.ToString().Trim().Equals(""))
                    {
                        MessageBox.Show("Invalid input. Please check again " + "\n\n  Parameter Name is not correct", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        lsRegisterParameter.Clear();
                        return null;
                    }
                }
                
           
                return lsRegisterParameter;
            }
            catch (Exception ex)
            {
                return lsRegisterParameter;
            }
        }

        private List<FDCMater> setRegisterList()
        {
            FDCMater MasterVo = new FDCMater();
            var tmp = from n in lisFDCMaster
                      where n.RecipeName == Recipe && n.Package == Package
                      select n;
            List<FDCMater> ls = new List<FDCMater>();
            if (lisFDCMaster.Count>1)
            {
                foreach (var n in tmp)
                {
                    ls.Add(n);
                }
            }
            else
            {
                lisFDCMaster.ForEach(n => ls.Add(n));
            }
           

            if (ls.Count>0)
            {
                MasterVo = ls[0];
            }

            if (MasterVo.BackupTime == null)
            {
                MasterVo.BackupTime = "";
            }
            if (status == edit || status == draft)
            {
                //MasterVo.Revision = (Int32.Parse(MasterVo.Revision.ToString().Trim()) + 1).ToString();
                if (status == edit)
                {// set new data to MasterVo 20180329
                    MasterVo.RecipeName = Recipe;
                    MasterVo.Package = txtPackage.Text;
                    MasterVo.Lead = txtLead.Text;
                }
            }
            else
            {//submitted
                MasterVo.Revision = 1;
                MasterVo.Status = "Submitted";
                MasterVo.Customer = "";
                MasterVo.RecipeName = Recipe;

                MasterVo.Package = txtPackage.Text;
                MasterVo.Dimension = "";
                MasterVo.Lead = txtLead.Text;
                MasterVo.Device = "";
                MasterVo.BondingNo = "";
                MasterVo.BondingRev = "";

                string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                MasterVo.CreatedTime = now;
                MasterVo.CreatedBy = Global.UserID; //ConResource.ID;
            }
          
            MasterVo.Division = Division;

            var oper = Operation.Split((char)32);
            if (oper != null && oper.Length > 1)
            {
                MasterVo.OperCode = oper[oper.Length - 1];
                string str = "";
                for (int i = 0; i < oper.Length-1; i++)
                {
                    str += oper[i];
                }
                MasterVo.OperName = str;

                if (Model.Trim().Equals("CL300(BSPT002") || Model.Trim().Equals("CL300(BSPT01,2,3)"))
                {
                    MasterVo.MachModel = Model.Substring(0, Model.IndexOf("("));
                }
                else
                {
                    //if (status == edit)
                    //{
                    //    MasterVo.MachModel = Model.Trim();
                    //    MasterVo.MachId = cboEqpNo.Text.Trim();
                    //}
                    //else
                    //{
                    MasterVo.MachModel = Model.Trim();
                    MasterVo.MachId = cboEqpNo.Text.Trim();
                    if (MasterVo.MachId.Trim().Equals("") && !MasterVo.Division.Contains("DPS"))
                    {
                        MessageBox.Show("Equip No Can not be Null", "Inform User");
                        return null;
                    }
                   // }
                   
                }

                if (cboEqpNo.SelectedIndex >= 0)
                {
                    if (EqpNoSelect)
                    {
                        if (cboEqpNo.SelectedIndex == 0)
                        {
                            MasterVo.MachId = "";
                        }
                        else
                        {
                            MasterVo.MachId = cboEqpNo.Text;
                        }
                    }                  
                }                

                if (Recipe.Trim().IndexOf("*") > 0 && Recipe.Trim().IndexOf("*") < Recipe.Trim().Length - 1)
                {
                    MessageBox.Show(Recipe + " --  * mark just can set before Recipe Name or After Recipe Name", "Invalid Parameter Format");
                    lsRegisterFDCMaster.Clear();
                    return null;
                }
                MasterVo.Memo1 = "";
            }
            lsRegisterFDCMaster.Add(MasterVo);
            return lsRegisterFDCMaster;
        }

        private void cboEqpNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            EqpNoSelect = true;
        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }


        static bool IsNumeric(string str)
        {
            for (int i = str.Length; --i >= 0;)
            {
                if (!char.IsDigit(str[i]))
                {
                    return false;
                }
            }
            return true;
        }

        private void ckeMail_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < ViewLimit.Rows.Count - 1; i++)
            {
                ViewLimit.Rows[i].Cells["eMail"].Value = this.ckeMail.Checked;
            }
        }

        private void ckMCHold_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < ViewLimit.Rows.Count - 1; i++)
            {
                ViewLimit.Rows[i].Cells["M/C hold"].Value = this.ckMCHold.Checked;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            iEMS.tcTab2.TabPages.RemoveAt(Parent.TabIndex);
        }

        private void btnSaveAsDraft_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Process save as draft?", "User confirmation"
                                                , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (createLimitItemVos() == true)
                {
                    if (status == edit)
                    {
                        if (!Editsave(false,draft, lsRegisterFDCMaster, lsRegisterParameter))
                        {
                            return;
                        }
                    }
                    else
                    {
                        if (!save(draft, lsRegisterFDCMaster, lsRegisterParameter))
                        {
                            return;
                        }
                    }
                    
                }
            }
        }

        private bool save(string status, List<FDCMater> lsRegisterFDCMaster, List<ParameterLimitGroup> lsRegisterParameter)
        {
            FDCMater voMaster = lsRegisterFDCMaster[0];
            voMaster.Status = status;

            if (!checkParametersExist(lsRegisterParameter,voMaster.MachModel,voMaster.MachId))
            {                
                return false;
            }

            if (voMaster.RecipeName.Trim().Equals(""))
            {
                MessageBox.Show("Recipe Name can not be null" ,"Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            if (status.Equals(submitted))
            {
                string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                voMaster.SubmittedTime = now;
                voMaster.SubmittedBy = Global.UserID; //ConResource.ID;
            }

            //is there same record
          
            if (isThereSameRecordRegister(voMaster))
            {
                MessageBox.Show("Exist same record at server. Can't register it ", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            foreach (ParameterLimitGroup voParam in lsRegisterParameter)
            {
                if (ckSetMailList.Checked)
                {
                    if (voParam.eMail.Equals(1))
                    {
                        voParam.eMailAddrs = MailDic[voParam.ParameterName];                     
                    }
                }
                else
                {
                    if (voParam.eMail.Equals(1))
                    {
                        voParam.eMailAddrs = MailDic[voParam.ParameterName];
                    }
                    else
                    {
                        voParam.eMailAddrs = "";
                    }
                }
            }

            // insert to DB
            bool bResult = insertRecordToDB(lsRegisterFDCMaster, lsRegisterParameter);

            if (!bResult)
            {
                MessageBox.Show("Failed to register.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            else
            {
                MessageBox.Show("Success registration.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                switch (status)
                {
                    case "Submitted":
                        iEMS.tcTab.SelectedIndex = 1;
                        break;
                    case "Draft":
                        iEMS.tcTab.SelectedIndex = 2;
                        break;
                    case "Deleted":
                        iEMS.tcTab.SelectedIndex = 3;
                        break;
                    default:
                        break;
                }
                
                ((iEMS)this.Tag).screenFdcList();
                iEMS.tcTab2.TabPages.RemoveAt(Parent.TabIndex);
                return true;
                //this.Close();
            }
        }

        private bool checkParametersExist(List<ParameterLimitGroup> lsRegisterParameter, string machModel, string machId)
        {
            List<Parameter> lsAllParameter = new List<Parameter>();
            lsAllParameter = GetParametersRecordByModelAndMachine(machModel, machId);

            foreach (ParameterLimitGroup p in lsRegisterParameter)
            {
                bool bfind = false;
                //string ParameterName = "";

                for (int i = 0; i < lsAllParameter.Count; i++)
                {
                    if (p.ParameterName.Equals(lsAllParameter[i].ParamName))
                    {
                        bfind = true;
                        break;
                    }
                    //ParameterName = p.ParameterName;                    
                }

                if (!bfind)
                {
                    MessageBox.Show("資料庫中找不到 \"" + p.ParameterName + "\" \n 請填寫Change Request申請單", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
            }
            return true;
        }
        

        public List<Parameter> GetParametersRecordByModelAndMachine(string model, string machID)
        {
            List<Parameter> lsParameters = new List<Parameter>();
            DBConnect cn = new DBConnect();
            string SQL = "select a.* from Att_ECIM.Parameters a where machmodel ='" + model + "'";
            if (!machID.Trim().Equals(""))
            {
                SQL += " and  machID ='" + machID + "'";
            }

             OleDbDataReader reader = cn.queryforResult(SQL, cn.getConnect());

            while (reader.Read())
            {
                ReadAllParameters(reader, lsParameters);
            }
            reader.Close();
            return lsParameters;
        }

        private void ReadAllParameters(OleDbDataReader reader, List<Parameter> lsParameter)
        {

            Parameter P = new Parameter
            {
                MachModel = reader["MachModel"].ToString(),
                MachId = reader["MachID"].ToString(),
                ParamName = reader["ParamName"].ToString(),
                ParamType = reader["ParamType"].ToString(),
                DataType = Convert.ToByte(reader["DataType"]),
                ParamUnit = reader["ParamUnit"].ToString(),
                CreatedTime = reader["CreatedTime"].ToString(),
                CreatedBy = reader["CreatedBy"].ToString()
            };
            lsParameter.Add(P);
        }

        private bool isThereSameParameter(FDCMater vo)
        {
            bool Result = false;
            //FDCMater vo = lsRegisterFDCMaster[0]; 
            List<int> serial = new List<int>(getLimitDefineSerial(vo));

            if (serial.Count > 0)
            {
                if (CompareParameterBySerial(serial, vo))
                {// is duplicate
                    MessageBox.Show("ParameterName is duplicate.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Result = true;
                }
                else
                {
                    Result = false;
                }                
            }
            else
            {
                Result = false;
            }

            return Result;
        }

        private List<int> getLimitDefineSerial(FDCMater vo)
        {
            
            DBConnect cn = new DBConnect();
            string sql = "select SERIALNO from ATT_ECIM.LimitDefinitionMaster a where a.machid = '" + vo.MachId + "' and a.status = 'Approved' and(a.recipename = '*' or a.recipename is null)";
            List<int> serialno = new List<int>();

            reader = cn.queryforResult(sql, cn.getConnect());

            while (reader.Read())
            {
                if (reader.HasRows)
                {
                    serialno.Add(Int32.Parse(reader["SerialNo"].ToString()));
                }
            }
            cn.closeData();

            return serialno;
        }

        private bool CompareParameterBySerial(List<int> serial, FDCMater vo)
        { // if duplicate return true   if not duplicate return false

            bool result = false;

            List<string> DBParameter = new List<string>(getParameterFromDB(serial));
            List<string> NewParameter = new List<string>(getParameterFromDB(vo.Serialno));

            if (DBParameter.Count>0)
            {
                for (int i = 0; i < NewParameter.Count; i++)
                {
                    string ParameterName = NewParameter[i];

                    if (DBParameter.Contains(ParameterName))
                    {
                        MessageBox.Show("Exist the same record("+ParameterName+") at server. \n Can not Approved it .", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return true;
                    }
                }
                result = false;
            }
            else
            {
                result = false;
            }

            return result;
        }

        private List<string> getParameterFromDB(int serialno)
        {
            List<string> dbParameter = new List<string>();
            DBConnect cn = new DBConnect();
            try
            {
                string sql = "select * from ATT_ECIM.LimitDefinitionParameter a where a.serialno = " + serialno + " ";

                reader = cn.queryforResult(sql, cn.getConnect());

                while (reader.Read())
                {
                    dbParameter.Add(reader["Parametername"].ToString());
                }

                cn.closeData();
            }
            catch (Exception ex)
            { }


            return dbParameter;
        }

        private List<string> getParameterFromDB(List<int> serial)
        {
            List<string> dbParameter = new List<string>();
            DBConnect cn = new DBConnect();
            try
            {
                string sql = "select * from ATT_ECIM.LimitDefinitionParameter a where a.serialno = " + serial[0] + " ";

                if (serial.Count > 1)
                {
                    for (int i = 1; i < serial.Count; i++)
                    {
                        sql += "or a.serialno =" + serial[i] + " ";
                    }
                }

                reader = cn.queryforResult(sql, cn.getConnect());

                while (reader.Read())
                {
                    dbParameter.Add(reader["Parametername"].ToString());
                }
                
                cn.closeData();
            }
            catch (Exception ex)
            {}
            
         
            return dbParameter;
        }

        private bool Editsave(bool Copy,string status, List<FDCMater> lsRegisterFDCMaster, List<ParameterLimitGroup> lsRegisterParameter)
        {
            FDCMater voMaster = lsRegisterFDCMaster[0];

            if (voMaster.RecipeName.Trim().Equals("")) {
                MessageBox.Show("Recipe Name can not be null");
                return false;
            }
            
            string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            if (status.Equals(submitted))
            {
                voMaster.SubmittedTime = now;
                voMaster.SubmittedBy = Global.UserID; //ConResource.ID;
            }

            //is there same record

            if (isThereSameRecordEdit(voMaster))
            {
                MessageBox.Show("Exist same record at server. Can't register it ", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            if (!isThisLatestRecord(voMaster))
            {
                MessageBox.Show("Selected item is not the latest. Please refresh the information.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            else
            {
                if (!Copy)
                {
                    voMaster.Revision = voMaster.Revision + 1;
                }                
            }

            voMaster.EditedBy = Global.UserID;
            voMaster.EditedTime = now;

            voMaster.Status = status;

            if (Copy)
            {
                bool enabledEquipmentNo = false;
                if (!Division.Trim().StartsWith("DPS"))
                {
                    enabledEquipmentNo = true;
                }

                if (enabledEquipmentNo)
                {
                    if (cboEqpNo.Text.Equals("") || !cboEqpNo.Text.Equals(Eqpno))
                    {
                        MessageBox.Show("Please select equipment No. have to Copy & Check controllable parameters", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }
                }

                createCopyDataToRegisterTab(voMaster);
                voMaster.Status = submitted;

            }
            else
            {
                bool bResult = updateRecordToDB(lsRegisterFDCMaster, lsRegisterParameter);

                if (!bResult)
                {
                    MessageBox.Show("Failed to register.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
                else
                {
                    if (status.Equals(submitted))
                    {
                        MessageBox.Show("Submit complete.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
               
                        iEMS.tcTab.SelectedIndex = 1;
                        ((iEMS)this.Tag).screenFdcList();
                        iEMS.tcTab2.TabPages.RemoveAt(iEMS.tcTab2.SelectedIndex);
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Save as draft complete.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        switch (status)
                        {
                            case "Submitted":
                                iEMS.tcTab.SelectedIndex = 1;
                                break;
                            case "Draft":
                                iEMS.tcTab.SelectedIndex = 2;
                                break;
                            case "Deleted":
                                iEMS.tcTab.SelectedIndex = 3;
                                break;
                            default:
                                break;
                        }

                    ((iEMS)this.Tag).screenFdcList();
                        iEMS.tcTab2.TabPages.RemoveAt(iEMS.tcTab2.SelectedIndex);
                        return true;
                    }
                }
            }
            return true;
        }

        private bool updateRecordToDB(List<FDCMater> lsRegisterFDCMaster, List<ParameterLimitGroup> lsRegisterParameter)
        {
            DBConnect cn = new DBConnect();
            OleDbConnection conn = cn.getConnect();
            OleDbCommand cmd = conn.CreateCommand();
            cmd.Connection = conn;
            OleDbTransaction trans = conn.BeginTransaction();
            cmd.Transaction = trans;

            //FDCMater voLimitGroupMasterForHistory = lsRegisterFDCMaster[0];
            FDCMater voLimitGroupMasterForHistory = lsFDCMasterHistory[0];
            string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            if (voLimitGroupMasterForHistory.Status.Equals(draft))
            {
                voLimitGroupMasterForHistory.CreatedTime = now;
                voLimitGroupMasterForHistory.CreatedBy = Global.UserID;

                voLimitGroupMasterForHistory.SubmittedTime = null;
                voLimitGroupMasterForHistory.SubmittedBy = "";
                voLimitGroupMasterForHistory.ApprovedTime = null;
                voLimitGroupMasterForHistory.ApprovedBy = "";
            }
            else if (voLimitGroupMasterForHistory.Status.Equals(submitted))
            {
                voLimitGroupMasterForHistory.CreatedTime = now;
                voLimitGroupMasterForHistory.CreatedBy = Global.UserID;
                voLimitGroupMasterForHistory.SubmittedTime = now;
                voLimitGroupMasterForHistory.SubmittedBy = Global.UserID;

                voLimitGroupMasterForHistory.ApprovedTime = null;
                voLimitGroupMasterForHistory.ApprovedBy = "";
                voLimitGroupMasterForHistory.EditedTime = null;
                voLimitGroupMasterForHistory.EditedBy = "";
            }
            //history master 
            if (insertHistoryRecord(cmd, voLimitGroupMasterForHistory) != 1)
            {
                MessageBox.Show("Submit Fail (insert history " + voLimitGroupMasterForHistory.RecipeName + ")");
                trans.Rollback();
                cn.closeData();
                return false;
            }

            int iSerial = voLimitGroupMasterForHistory.Serialno;
            String sMachId = voLimitGroupMasterForHistory.MachId.Trim();
            
            int iRevision = voLimitGroupMasterForHistory.Revision;
            //history parameter
            for (int i = 0; i < lsParameterHistory.Count; i++)
            {
                //ParameterLimitGroup voParam = lsRegisterParameter[i];
                ParameterLimitGroup voParam = lsParameterHistory[i];
                voParam.SerialNo = iSerial;
                voParam.Revision = iRevision;
                if (insertHistoryParameterRecord(cmd, voParam) != 1)
                {
                    MessageBox.Show("Submit Fail (insert history" + voParam.ParameterName + ")");
                    trans.Rollback();
                    cn.closeData();
                    return false;
                }
            }


            //delete previous record
            if (lsRegisterFDCMaster[0].OperCode.Trim().IndexOf("PLAT") > -1 && lsRegisterFDCMaster[0].MachModel.ToUpper().StartsWith("RAIDER") && cboCup.Text != "")
            {
                delete(cmd, lsRegisterFDCMaster[0].Serialno, cboCup.Text);
            }
            else if (delete(cmd, lsRegisterFDCMaster[0].Serialno) != lsParameterHistory.Count())
            { }

            //insert new record
            
            //master
            FDCMater voLimitMaster = lisFDCMaster[0];
            if (!voLimitMaster.MachId.Equals(sMachId)) voLimitMaster.MachId = sMachId;
            if (update(cmd, voLimitMaster) != 1)
            {
                MessageBox.Show("Submit Fail (insert " + voLimitMaster.RecipeName + ")");
                trans.Rollback();
                cn.closeData();
                return false;
            }

            int intSerial = voLimitMaster.Serialno;
            int intRevision = voLimitMaster.Revision;
            //parameter
            for (int i = 0; i < lsRegisterParameter.Count; i++)
            {
                ParameterLimitGroup voParam = lsRegisterParameter[i];
                voParam.SerialNo = intSerial;
                voParam.Revision = intRevision;
                if (insertParameterRecord(cmd, voParam) != 1)
                {
                    MessageBox.Show("Submit Fail (insert " + voParam.ParameterName + ")");
                    trans.Rollback();
                    cn.closeData();
                    return false;
                }
            }
            trans.Commit();

            cn.closeData();

            return true;
        }

        private int update(OleDbCommand cmd, FDCMater voLimitMaster)
        {
            int iReturn = -1;
            if (voLimitMaster.ApprovedTime != null) voLimitMaster.ApprovedTime = voLimitMaster.ApprovedTime.IndexOf("午") > -1 ? Convert.ToDateTime(voLimitMaster.ApprovedTime).ToString("yyyy/MM/dd HH:mm:ss") : voLimitMaster.ApprovedTime;
            if (voLimitMaster.SubmittedTime != null) voLimitMaster.SubmittedTime = voLimitMaster.SubmittedTime.IndexOf("午") > -1 ? Convert.ToDateTime(voLimitMaster.SubmittedTime).ToString("yyyy/MM/dd HH:mm:ss") : voLimitMaster.SubmittedTime;
            if (voLimitMaster.EditedTime != null) voLimitMaster.EditedTime = voLimitMaster.EditedTime.IndexOf("午") > -1 ? Convert.ToDateTime(voLimitMaster.EditedTime).ToString("yyyy/MM/dd HH:mm:ss") : voLimitMaster.EditedTime;
            if (voLimitMaster.CreatedTime != null) voLimitMaster.CreatedTime = voLimitMaster.CreatedTime.IndexOf("午") > -1 ? Convert.ToDateTime(voLimitMaster.CreatedTime).ToString("yyyy/MM/dd HH:mm:ss") : voLimitMaster.CreatedTime;
            string sql = "update ATT_ECIM.LimitDefinitionMaster set Revision='"+voLimitMaster.Revision+"',Status='"+voLimitMaster.Status+"',Division='"+voLimitMaster.Division+"',OperCode='"+voLimitMaster.OperCode+"',MachModel='"+voLimitMaster.MachModel+"',MachId='"+voLimitMaster.MachId+"', Customer='"+voLimitMaster.Customer+"',Package='"+voLimitMaster.Package+"',Dimension='"+voLimitMaster.Dimension+"',Lead='"+voLimitMaster.Lead+"',Device='"+voLimitMaster.Device+"',BondingNo='"+voLimitMaster.BondingNo+"',BondingRev='"+voLimitMaster.BondingRev+"',RecipeName='"+voLimitMaster.RecipeName+ "',SubmittedTime=TO_DATE('" + voLimitMaster.SubmittedTime + "', 'YYYY/MM/DD HH24:MI:SS'),SubmittedBy='" + voLimitMaster.SubmittedBy+ "',ApprovedTime=TO_DATE('" + voLimitMaster.ApprovedTime + "', 'YYYY/MM/DD HH24:MI:SS'),ApprovedBy='" + voLimitMaster.ApprovedBy+ "',EditedTime=TO_DATE('" + voLimitMaster.EditedTime + "'),EditedBy='" + voLimitMaster.EditedBy+"',Memo1='"+voLimitMaster.Memo1+"',Memo2='"+voLimitMaster.Memo2+"',Memo3='"+voLimitMaster.Memo3+"',EFGPSN ='"+voLimitMaster.EFGPSN+"' where SerialNo='"+voLimitMaster.Serialno+"'";
            //if (Global.passEFGP) sql = "update ATT_ECIM.LimitDefinitionMaster set Revision='" + voLimitMaster.Revision + "',Status='" + voLimitMaster.Status + "',Division='" + voLimitMaster.Division + "',OperCode='" + voLimitMaster.OperCode + "',MachModel='" + voLimitMaster.MachModel + "',MachId='" + voLimitMaster.MachId + "', Customer='" + voLimitMaster.Customer + "',Package='" + voLimitMaster.Package + "',Dimension='" + voLimitMaster.Dimension + "',Lead='" + voLimitMaster.Lead + "',Device='" + voLimitMaster.Device + "',BondingNo='" + voLimitMaster.BondingNo + "',BondingRev='" + voLimitMaster.BondingRev + "',RecipeName='" + voLimitMaster.RecipeName + "',SubmittedTime=TO_DATE('" + voLimitMaster.SubmittedTime + "', 'YYYY/MM/DD HH24:MI:SS'),SubmittedBy='" + voLimitMaster.SubmittedBy + "',ApprovedTime=TO_DATE('" + voLimitMaster.ApprovedTime + "', 'YYYY/MM/DD HH24:MI:SS'),ApprovedBy='" + voLimitMaster.ApprovedBy + "',EditedTime=TO_DATE('" + voLimitMaster.EditedTime + "'),EditedBy='" + voLimitMaster.EditedBy + "',Memo1='" + voLimitMaster.Memo1 + "',Memo2='" + voLimitMaster.Memo2 + "',Memo3='" + voLimitMaster.Memo3 + "' where SerialNo='" + voLimitMaster.Serialno + "'"; 
            cmd.CommandText = sql;
            iReturn = cmd.ExecuteNonQuery();//cn.insert(sql, cn.getConnect());
            return iReturn;
        }

        private int delete(OleDbCommand cmd, int serialno)
        {
            try
            {
                int iReturn = -1;
                string sql = "delete from ATT_ECIM.LimitDefinitionParameter where SerialNo='"+serialno+"'";
                cmd.CommandText = sql;
                iReturn = cmd.ExecuteNonQuery();//cn.insert(sql, cn.getConnect());
                return iReturn;
            }
            catch (Exception ex)
            {
                //cn.closeData();
                return -1;
            }
        }

        private int delete(OleDbCommand cmd, int serialno, string text)
        {
            int iReturn = -1;
            text = null == text ? "%" : text + " %";

            string sql = "delete from ATT_ECIM.LimitDefinitionParameter where SerialNo='" + serialno + "' and PARAMETERNAME like '"+text+"'";
            cmd.CommandText = sql;
            iReturn = cmd.ExecuteNonQuery();//cn.insert(sql, cn.getConnect());
            //cn.closeData();
            return iReturn;
        }

        private int insertHistoryParameterRecord(OleDbCommand cmd, ParameterLimitGroup voParam)
        {
            try
            {
                int iReturn = -1;
                string sql = "";
                //if (ckSetMailList.Checked) // mail list same
                //{
                //    voParam.eMailAddrs = sMailList;
                //}
                //else
                //{
                //    voParam.eMailAddrs = MailDic[voParam.ParameterName];

                //}

                //insert into ATT_ECIM.LimitDefinitionParameterHis values (?,?,?,?,?, ?,?,?,?,?, ?,?,?,?,?, ?,?,?,?,?, ?,?,?,?,?, ?,?,?)
                sql = "insert into ATT_ECIM.LimitDefinitionParameterHis values ('" + voParam.SerialNo + "','" + voParam.ParameterName + "','" + voParam.Revision + "','" + voParam.CheckMin + "','" + voParam.AllowMinEqual + "', '" + voParam.MinValue + "','" + voParam.CheckMax + "','" + voParam.AllowMaxEqual + "','" + voParam.MaxValue + "','" + voParam.TerminalMsg + "', '" + voParam.eMail + "','" + voParam.eMailAddrs + "','" + voParam.MCInhibit + "','" + voParam.MCHold + "','" + voParam.eOCAP + "', '" + voParam.eOCAPTemplete + "','" + voParam.LotHold + "','" + voParam.StepName + "','" + voParam.StepValue + "','" + voParam.EventId + "', '" + voParam.WarmUpTime + "','" + voParam.StartDetectValue + "','" + voParam.SecondFlag + "','" + voParam.CheckRMin + "','" + voParam.RMinValue + "', '" + voParam.CheckRMax + "','" + voParam.RMaxValue + "','" + voParam.RChartFlag + "','" + voParam.MonitorTime + "')";
                cmd.CommandText = sql;
                iReturn = cmd.ExecuteNonQuery();//cn.insert(sql, cn.getConnect());
                return iReturn;
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        private int insertHistoryRecord(OleDbCommand cmd, FDCMater vo)
        {
            try
            {
                int iReturn = -1;
                string sql = "";

                if (vo.ApprovedTime != null) vo.ApprovedTime = vo.ApprovedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.ApprovedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.ApprovedTime;
                if (vo.SubmittedTime != null) vo.SubmittedTime = vo.SubmittedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.SubmittedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.SubmittedTime;
                if (vo.EditedTime != null) vo.EditedTime = vo.EditedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.EditedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.EditedTime;
                if (vo.CreatedTime != null) vo.CreatedTime = vo.CreatedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.CreatedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.CreatedTime;

                sql = "insert into ATT_ECIM.LimitDefinitionMasterHistory values ('" + vo.Serialno + "' , '" + vo.Revision + "','" + vo.Status + "','" + vo.Division + "','" + vo.OperCode + "','" + vo.MachModel + "','" + vo.MachId + "','" + vo.Customer + "','" + vo.Package + "','" + vo.Dimension + "','" + vo.Lead + "','" + vo.Device + "','" + vo.BondingNo + "','" + vo.BondingRev + "','" + vo.RecipeName + "',TO_DATE('" + Convert.ToDateTime(vo.CreatedTime).ToString("u").Substring(0, 19) + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.CreatedBy + "',TO_DATE('" + vo.SubmittedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.SubmittedBy + "',TO_DATE('" + vo.ApprovedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.ApprovedBy + "',TO_DATE('" + vo.EditedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.EditedBy + "',TO_DATE('" + vo.BackupTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.BackupBy + "','" + vo.Memo1 + "','" + vo.Memo2 + "','" + vo.Memo3 + "','" + vo.EFGPSN + "')";
                //if(Global.passEFGP)sql = "insert into ATT_ECIM.LimitDefinitionMasterHistory values ('" + vo.Serialno + "' , '" + vo.Revision + "','" + vo.Status + "','" + vo.Division + "','" + vo.OperCode + "','" + vo.MachModel + "','" + vo.MachId + "','" + vo.Customer + "','" + vo.Package + "','" + vo.Dimension + "','" + vo.Lead + "','" + vo.Device + "','" + vo.BondingNo + "','" + vo.BondingRev + "','" + vo.RecipeName + "',TO_DATE('" + Convert.ToDateTime(vo.CreatedTime).ToString("u").Substring(0, 19) + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.CreatedBy + "',TO_DATE('" + vo.SubmittedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.SubmittedBy + "',TO_DATE('" + vo.ApprovedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.ApprovedBy + "',TO_DATE('" + vo.EditedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.EditedBy + "',TO_DATE('" + vo.BackupTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.BackupBy + "','" + vo.Memo1 + "','" + vo.Memo2 + "','" + vo.Memo3 + "')";
                cmd.CommandText = sql;
                iReturn = cmd.ExecuteNonQuery();//cn.insert(sql, cn.getConnect());
                return iReturn;
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        private void createCopyDataToRegisterTab(FDCMater voMaster)
        {
            ((iEMS)this.Tag).OpenForCopy(voMaster);
        }

        private bool isThereSameRecordEdit(FDCMater voMaster)
        {
            int serialNo = getSerialNo(voMaster);
            if (serialNo < 0)
                return false;

            if (serialNo == voMaster.Serialno) return false;
            else return true;
        }

        private bool isThereSameRecordRegister(FDCMater voMaster)
        {
            int serialNo = getSerialNo(voMaster);
                if (serialNo >= 0) return true;
                else
                {
                    return false;
                }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Process Delete?", "User confirmation"
                                               , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (delete())
                {
                    MessageBox.Show("Delete completed", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    switch (status)
                    {
                        case "Submitted":
                            iEMS.tcTab.SelectedIndex = 1;
                            break;
                        case "Draft":
                            iEMS.tcTab.SelectedIndex = 2;
                            break;
                        case "Deleted":
                            iEMS.tcTab.SelectedIndex = 3;
                            break;
                        default:
                            break;
                    }
                    
                ((iEMS)this.Tag).screenFdcList();
                iEMS.tcTab2.TabPages.RemoveAt(iEMS.tcTab2.SelectedIndex);
                }
            }
        }

        private bool delete()
        {
            FDCMater vo = lisFDCMaster[0];
            if (!isThisLatestRecord(vo))
            {
                MessageBox.Show("Selected item is not the latest. Please refresh the information.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            vo.Status = "Deleted";
            string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            vo.EditedTime = now;
            vo.EditedBy = Global.UserID; //ConResource.ID;

            if (delete(vo) > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        private int delete(FDCMater vo)
        {
            DBConnect cn = new DBConnect();
            try
            {
               
                int iReturn = -1;
                string sql = "update ATT_ECIM.LimitDefinitionMaster set Status='" + vo.Status + "', EditedTime=TO_DATE('" + vo.EditedTime + "', 'YYYY/MM/DD HH24:MI:SS'), EditedBy='" + vo.EditedBy + "', Memo3='" + vo.Memo3 + "' where SerialNo='" + vo.Serialno + "'";
                iReturn = cn.insert(sql, cn.getConnect());
                cn.closeData();
                return iReturn;
            }
            catch (Exception ex)
            {
                cn.closeData();
                return -1;
            }
          
        }

        private bool isThisLatestRecord(FDCMater vo)
        {
            DBConnect cn = new DBConnect();
            try
            {
                            //select Revision, Status, EditedTime from ATT_ECIM.LimitDefinitionMaster where SerialNo=?
                string sql = "select Revision, Status, EditedTime from ATT_ECIM.LimitDefinitionMaster where SerialNo='" + vo.Serialno + "'";// and machid='"+ vo.MachId +"'";
               
                reader = cn.queryforResult(sql, cn.getConnect());
                if (reader.HasRows)
                {
                    reader.Read();

                    int reversion =Int32.Parse(reader["Revision"].ToString());
                    string status = reader["STATUS"].ToString().Trim();
                    // string editedTime = reader["EditedTime"].ToString().Trim();
                    DateTime editedTime = new DateTime();
                    DateTime editedTime1 = new DateTime();
                    if (reader["EditedTime"].ToString().Trim() != "")
                    {
                        editedTime = reader.GetDateTime(reader.GetOrdinal("EditedTime"));
                        
                    }
                    TimeSpan ts = new TimeSpan();
                    ts = editedTime - editedTime1;

                    double lvoTime = 0, ledtiedTime = 0;

                    if (editedTime != null) ledtiedTime = ts.TotalSeconds;
                    if (vo.EditedTime != null) lvoTime = ts.TotalSeconds;

                    if (reversion == vo.Revision)
                    {
                        if (status.Equals(vo.Status))
                        {
                            if ((lvoTime == ledtiedTime) || (lvoTime + 1 == ledtiedTime)) return true;
                            else return false;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            finally { cn.closeData(); }            
        }

        private void btnApprove_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Process approve?", "User confirmation"
                                               , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (Approve())
                {
                    MessageBox.Show("Approve completed", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    try
                    {
                        if (!(Division.IndexOf("T1") > -1)) //for T1 don't needs change tab
                        {
                            iEMS.tcTab.SelectedIndex = 0;
                            ((iEMS)this.Tag).screenFdcList();
                            //iEMS.tcTab2.TabPages.RemoveAt(Parent.TabIndex);
                            //MessageBox.Show("Before \n SelectedIndex = " + iEMS.tcTab2.SelectedIndex + "\n Count = " + iEMS.tcTab2.TabCount);
                            iEMS.tcTab2.TabPages.RemoveAt(iEMS.tcTab2.SelectedIndex);
                            //iEMS.tcTab2.TabPages.Remove(iEMS.tcTab2.SelectedTab);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    
                    //return true;

                }
               
            }
        }

        private bool Approve()
        {
            if (Global.isFDC)//for change to FDC
            {
                Global.Level = "FDC";
            }

            if (Global.Level != "FDC")
            {
                MessageBox.Show("Please ask your Manager to Approve!!", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            FDCMater vo = lisFDCMaster[0];
            if (!isThisLatestRecord(vo))
            {
                MessageBox.Show("Selected item is not the latest. Please refresh the information.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            //compare with DB , find parameter is duplicate will return false
            if (txtRecipe.Text.Equals("*") && isThereSameParameter(vo))
            {               
                return false;
            }
            
            vo.Status = Approved;
            string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            vo.ApprovedTime = now;
            vo.ApprovedBy = Global.UserID; //ConResource.ID;

            //marked by Jim 20180608
            //vo.EditedTime = now;
            //vo.EditedBy = Global.UserID;
            

            string msg = Global.UserID + " 於 " + now + " 執行 Approved,\nRecipe Name = " + vo.RecipeName +" ,\nEqpNo = " + vo.MachId;

            if (approve_reject(vo) == 1)
            {
                
                send_email(msg, "iEMS-Approve Notice", "Chuck.Chen@amkor.com,jim.yao@amkor.com");
                return true;
            }
            else
            {
                return false;
            }
        }

        public void send_email(string msg, string mysubject, string address)
        {
            MailMessage mailMessage = new MailMessage("ATT_eCIM_Admin@amkor.com", address);
            mailMessage.Subject = mysubject;//E-mail主旨
            mailMessage.CC.Add("ATT_eCIM_Admin@amkor.com");
            mailMessage.Body = msg;//E-mail內容   
            SmtpClient smtpClient = new SmtpClient("t1ns01.amkor.com", 25);
            //SmtpClient smtpClient = new SmtpClient("exim.amkor.com", 25); 
            smtpClient.Send(mailMessage);         
        }

        private int approve_reject(FDCMater vo)
        {
            int iReturn = -1;
            bool CanRequest = true;
            //vo.EditedTime = vo.EditedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.EditedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.EditedTime;
            //string sql = "update ATT_ECIM.LimitDefinitionMaster set Status='" + vo.Status + "', ApprovedTime=TO_DATE('" + vo.ApprovedTime + "', 'YYYY/MM/DD HH24:MI:SS'), ApprovedBy='" + vo.ApprovedBy + "', EditedTime=TO_DATE('" + vo.EditedTime + "', 'YYYY/MM/DD HH24:MI:SS'), EditedBy='" + vo.EditedBy + "', Memo2='" + vo.Memo2 + "' where SerialNo='" + vo.Serialno + "'";
            string sql = "update ATT_ECIM.LimitDefinitionMaster set Status='" + vo.Status + "', ApprovedTime=TO_DATE('" + vo.ApprovedTime + "', 'YYYY/MM/DD HH24:MI:SS'), ApprovedBy='" + vo.ApprovedBy + "', Memo2='" + vo.Memo2 + "' where SerialNo='" + vo.Serialno + "'";

            /*
             * 1. check Division
             * 2. if EFGPSN is null, can Rquest EFGP 
             * 3. if EFGPSN is not null, check SN status is close or not
             * 4. if is close, can Request EFGP or show message to User for wait approval.
             */
            if (!Global.passEFGP)//add Flag for pass EFGP by Jim 20180511
            {
                if (vo.Division.IndexOf("T1") > -1)//only for T1 Bump needs send EFGP
                {
                    EasyFlow EF = new EasyFlow();
                    //20180214 for send EF     
                    if (Global.isUAT)//change EFGP environment 20180213
                    {
                        EF.SetEFGPUrl(Global.EFGP_UAT);
                        //EF.SetEFGPUrl(Global.EFGP_PROD);//for test
                    }
                    else
                    {
                        EF.SetEFGPUrl(Global.EFGP_PROD);
                    }

                    if (vo.EFGPSN != null)//(EF.getProcessStatus(vo.EFGPSN).IndexOf("close") > -1)//(vo.EFGPSN.Trim().Equals(""))//needs request EFGP
                    {//already request
                        if (!vo.EFGPSN.Trim().Equals(""))
                        {
                            CanRequest = false;

                            if ((EF.getProcessStatus(vo.EFGPSN).IndexOf("close") > -1))
                            {
                                CanRequest = true;
                            }
                        }
                    }

                    if (CanRequest && vo.Status.Equals(Reject))// easyflow form approve complete and want to reject
                    {
                        DBConnect cn = new DBConnect();
                        iReturn = cn.insert(sql, cn.getConnect());
                        cn.closeData();
                    }
                    else if (CanRequest && vo.Status.Equals(Approved)) //easyflow form approve complete and want to approve
                    {
                        string sProcessID = "eCIMFDC01";
                        string sUnitID = EF.getUnitId(Global.UserID);
                        string sFormID = EF.getFormOID(sProcessID);
                        string sService = EF.getFormFieldTemplate(sProcessID);

                        sService = EF.CombineFDCApproval(sService, sql, vo, sUnitID);
                        sService = sService.Replace("defaultValue", "");
                        sService = sService.Replace("perDataProId=\"\"> <", "perDataProId=\"\"><");
                        string sSerialNo = EF.invokeProcess(sProcessID, Global.UserID, sUnitID, sFormID, sService, "");
                        string sPath = "C://FDC_Log//";
                        //save 
                        WriteLog(sPath,sSerialNo,sService);

                        //save SerialNo to DB
                        sql = "update ATT_ECIM.LimitDefinitionMaster set EFGPSN='" + sSerialNo + "'where SerialNo='" + vo.Serialno + "'";
                        DBConnect cn = new DBConnect();
                        iReturn = cn.insert(sql, cn.getConnect());
                        cn.closeData();
                    }
                    else
                    {
                        //show message for wait approval
                        MessageBox.Show("This Setting already send to Approval,\n Please wait for Approval Finish.\n Easy Flow SerialNo : " + vo.EFGPSN + "", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else//for don't needs request EFGP
                {
                    DBConnect cn = new DBConnect();
                    iReturn = cn.insert(sql, cn.getConnect());
                    cn.closeData();
                }
            }
            else //pass EFGP add by Jim 20180511
            {
                DBConnect cn = new DBConnect();
                iReturn = cn.insert(sql, cn.getConnect());
                cn.closeData();
            }
           
            return iReturn;
        }

        private void WriteLog(string sPath, string sSerialNo, string sService)
        {
            try
            {
                string sFileName = DateTime.Now.ToString(DateTime.Now.ToString("yyyy_MM_dd")) + ".txt";
                string sFullPath = sPath + sFileName;
                string msg = DateTime.Now.ToString("HH:mm:ss") + " / " + sSerialNo + " / " + sService;
                if (!Directory.Exists(sPath))
                {
                    Directory.CreateDirectory(sPath);
                }
                if (!System.IO.File.Exists(sFullPath))
                {
                    //File.Create(sFullPath);
                    File.WriteAllText(sFullPath, msg + "\r\n");
                }
                else
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(sFullPath, true))
                    {
                        file.WriteLine(msg);
                    }
                }

                //FileStream fs = new FileStream(sFullPath, FileMode.Create);
                //StreamWriter sw = new StreamWriter(fs);
               
                //sw.WriteLine(msg);
                //sw.Flush();
                //sw.Close();
                //fs.Dispose();

            }
            catch (Exception ex)
            {
                
            }
            
           
        }

        private void cboCup_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sCup = cboCup.Text;
            //getParameter();
            fillGridView(lsCollect, lsParameterLG);

        }

        private void lbEqpNo_Click(object sender, EventArgs e)
        {

        }

        private void btnReject_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Process reject?", "User confirmation"
                                               , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (reject())
                {
                    MessageBox.Show("Reject complete.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //switch (status)
                    //{
                    //    case "Submitted":
                    //        iEMS.tcTab.SelectedIndex = 1;
                    //        break;
                    //    case "Draft":
                    //        iEMS.tcTab.SelectedIndex = 2;
                    //        break;
                    //    case "Deleted":
                    //        iEMS.tcTab.SelectedIndex = 3;
                    //        break;
                    //    default:
                    //        break;
                    //}

                    ((iEMS)this.Tag).screenFdcList();
                    iEMS.tcTab2.TabPages.RemoveAt(Parent.TabIndex);
                }
            }
        }

        private bool reject()
        {
            FDCMater voMaster = lisFDCMaster[0];
            if (!isThisLatestRecord(voMaster))
            {
                MessageBox.Show("Selected item is not the latest. Please refresh the information.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            voMaster.Status = Reject;
            string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            voMaster.EditedTime = now;
            voMaster.EditedBy = Global.UserID;

            if (approve_reject(voMaster) >=1)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        private void ckSetMailList_CheckedChanged(object sender, EventArgs e)
        {
            if (ckSetMailList.Checked == true)
            {
                bool ReadOnly;

                switch (status)
                {
                    case "Edit":
                        goto case "Register";
                    case "Register":
                        ReadOnly = false;
                        break;
                    default:
                        ReadOnly = true;
                        break;
                }
              
                string mail = "";

                if (MailDic.Count > 0 || lsParameterLG.Count == 0)
                {
                    if (MailDic.ContainsKey("ALL"))
                    {
                        mail = MailDic["ALL"];
                    }
                }

                sMailList = "";

                ParameterSetting ParameterSetting = new ParameterSetting(Division, Operation, sMailList, ReadOnly, mail);

                ParameterSetting.ShowDialog();
                if (ParameterSetting.sMailList != "")
                {
                    sMailList = ParameterSetting.sMailList; // for set all email list
                    if (MailDic.Keys.ToList().Contains("ALL"))
                    {
                        MailDic.Remove("ALL");
                        MailDic.Add("ALL", sMailList); // for set singal email list
                    }
                    else
                    {
                        MailDic.Add("ALL", ParameterSetting.sMailList); // for set singal email list
                    }
                }
            }
            else
            {
                sMailList = "";
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Process copy?", "User confirmation"
                                                , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                bCopy = true;
                if (createLimitItemVos() == true)
                {              
                    if (status == edit)
                    {
                        if (bCopy)
                        {
                            if (!Editsave(true, "Copy", lsRegisterFDCMaster, lsRegisterParameter))
                            {
                                return;
                            }
                        }
                        else
                        {
                            if (!Editsave(true, draft, lsRegisterFDCMaster, lsRegisterParameter))
                            {
                                return;
                            }
                        }
                        
                    }
                    else
                    {
                        if (!save(draft, lsRegisterFDCMaster, lsRegisterParameter))
                        {
                            return;
                        }
                    }
                }
            }
        }

        private void ckUse_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < ViewLimit.Rows.Count - 1; i++)
            {
                ViewLimit.Rows[i].Cells["Use"].Value = this.ckUse.Checked;
            }
        }

        private void ViewLimit_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex == 2) // for not edit parameter name
            {
                e.Cancel = true;
            }
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnCompare_Click(object sender, EventArgs e)
        {
            try
            {
                if (lsFDCMasterHistory.Count>0 && lsParameterHistory.Count>0)
                {
                    Compare_Parameter CP = new Compare_Parameter(lisFDCMaster, lsFDCMasterHistory, lsParameterLG, lsParameterHistory);
                    CP.ShowDialog();
                }
                else
                {//can not find edit record
                    MessageBox.Show("This is the Last Version, can not compare!!!","Info User");
                }

            }
            catch (Exception)
            {

                throw;
            }

        }
    }
}
