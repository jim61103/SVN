using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
using System.Data.OracleClient;
using System.Diagnostics;
using Encoder.Security;
using System.Runtime.InteropServices;

namespace iEMS_Setting
{

   

    public partial class iEMS : Form
    {
       

        //OleDbConnection cn;
        //OleDbCommand cmd;
        OleDbDataReader reader;
        OleDbDataAdapter adt;
        public string str = "";

        //public Dictionary<int,string> dicView = new Dictionary<int,string>();

        public static DataGridView dgApproved = new DataGridView();
        public static DataGridView dgSubmitted = new DataGridView();
        public static DataGridView dgDraft = new DataGridView();
        public static DataGridView dgDeleted = new DataGridView();

        public static TabControl tcTab = new TabControl();
        public static TabControl tcTab2 = new TabControl();

        public int tcTabIndex = 0;
        public bool tcTabSelect = false;

        public Dictionary<string, string> dicDivision = new Dictionary<string, string>();
        public Dictionary<string, string> dicOperation = new Dictionary<string, string>();
        public Dictionary<string, string> dicEquipModel = new Dictionary<string, string>();

        public List<string> lsDivision = new List<string>();
        public List<string> lsOperation = new List<string>();
        public List<string> lsEqpModel = new List<string>();
        public List<string> lsEqpNo = new List<string>();
        public List<string> lsStatus = new List<string>();
        public List<string> lseCOAPTemplete = new List<string>();

        public static List<Parameter> lsParameter = new List<Parameter>();
        public static List<Parameter> lsCollect = new List<Parameter>();
        public static List<cimrsc> lsCimRsc = new List<cimrsc>();
        public static List<FDCMater> lsFDCmater = new List<FDCMater>();
        public string sDivision = "", sOperation = "", sEqpModel = "", sEqpNo = "",sRecipe="", Serialno="",sPackage="";
        //將視窗移動到最上層
        [DllImport("USER32.DLL")]//引用User32.dll
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        //尋找視窗
        [DllImport("USER32.DLL")]//引用User32.dll
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public iEMS()
        {
                InitializeComponent();
                //this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
                dgApproved = ApprovedView;
                dgSubmitted = SubmittedView;
                dgDraft = DraftView;
                dgDeleted = SubmittedView;
                tcTab = tabControl1;
                tcTab2 = tabControl2;                  
        }

        /// <summary> 
        /// 判定是否運行於64bit作業系統
        /// </summary> 
        /// <returns>是否為64bit的作業系統</returns> 
        public static bool Is64bit()
        {
            if (Environment.Is64BitOperatingSystem)
                return true;
            else
                return false;
            //return IntPtr.Size == 8;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                Process[] MyProcess = Process.GetProcessesByName("iEMS-Setting");
                if (MyProcess.Count() > 0)
                {
                    for (int i = 0; i < MyProcess.Count() - 1; i++)
                    {
                        ////提到最上層
                        //IntPtr main = FindWindow("WindowsForms10.Window.8.app.0.1ca0192_r12_ad1", null);
                        //if (main == IntPtr.Zero)
                        //{
                        //    return;
                        //}
                        //SetForegroundWindow(main);
                        if (MyProcess[i].StartTime > MyProcess[i + 1].StartTime)
                        {
                            MyProcess[i].Kill();
                        }
                        else
                        {                          
                            MyProcess[i+1].Kill();
                        }
                    }
                }

                try
                {
                    LogInGUI LogIn = new LogInGUI();
                    this.Hide();
                    LogIn.ShowDialog();

                    bool LogInResule = LogIn.Success;

                    if (!LogInResule)
                    {
                        this.Close();
                        return;
                    }
                    else
                    {
                        Global.UserID = LogIn.sUserID;
                        this.Show();
                    }
                }
                catch (Exception ex)
                {

                    throw;
                }
                //DateTime last = new DateTime();
                //DateTime now = new DateTime();
                //string filename = "";
                //bool bis64 = false;
                //string spath = "";
                //bis64 = Is64bit();

                //MessageBox.Show(bis64.ToString());

                //if (bis64)
                //{
                //    spath = @"C:\Program Files (x86)\Amkor";
                //}
                //else
                //{
                //    spath = @"C:\Program Files\Amkor";
                //}
                //MessageBox.Show(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
                //spath = "";
                //spath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                //MessageBox.Show(spath);
                //try
                //{
                //    foreach (string fname in System.IO.Directory.GetFiles(spath, "Equip_Screen_Config-*.obj", SearchOption.AllDirectories))
                //    {
                //        now = File.GetLastWriteTime(fname); //File.GetLastAccessTime(fname);
                //        TimeSpan ts = last - now;
                //        if (ts.TotalSeconds < 0) //now is the last access file
                //        {
                //            last = now;
                //            filename = fname;
                //        }
                //    }
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("Error for get Equip_Screen_Config \n" + ex.ToString());
                //    this.Close();
                //}
                

                //MessageBox.Show(filename);
                //try
                //{
                //    string ID = filename.Substring(filename.IndexOf("Equip_Screen_Config-") + 20, filename.IndexOf("-ATT") - (filename.IndexOf("Equip_Screen_Config-") + 20));
                //    Global.UserID = ID;
                //}
                //catch (Exception ex)
                //{
                //    throw new System.Exception("File name is " + filename);
                //}
                
                
                //get Config
                //Config config = new Config();
                //char s = '\\';
                //string[] temp = filename.Split(s);
                //string Path = temp[0] + "\\" + temp[1] + "\\" + temp[2] + "\\" + temp[3]+ "\\Att_Config.properties";
                //if (System.IO.File.Exists(Path))
                //{
                //    getConfigProperties(Path, config);
                //}
                //Global.Env ="@ "+ config.Environment;
                //// isUAT = True only for 110.1 used, else always =false
                //if (config.Environment.Equals("UAT") || Global.isUAT)
                //{
                //    Global.ConnectDB = Global.BUMP_UAT_ConnectDB;
                //    Global.Env = "@ UAT";
                //    //don't show UAT
                //}
                //else
                //{
                //    if (config.ProductionLine.Equals("DLP"))
                //    {
                //        Global.ConnectDB = Global.DLPConnectDB;
                //    }
                //    else //if (config.ProductionLine.Equals("Bump"))
                //    {
                //        Global.ConnectDB = Global.BUMPConnectDB;
                //    }
                //}

                if (Global.isUAT)
                {
                    Global.ConnectDB = Global.BUMP_UAT_ConnectDB;
                    //Global.ConnectDB = Global.sConnString;
                    Global.Env = "@ UAT";
                    
                }
                else
                {
                    Global.Env = "@ Production";
                    Global.ConnectDB = Global.BUMPConnectDB;
                }

                this.Text += Global.Env;
                //Global.UserID = "90379";
                //get ecim level
                try
                {
                    DBConnect cn = new DBConnect();
                    string sql = "select * from getUserInfoLevelNew where employeename = '" + Global.UserID + "'";
                    reader = cn.queryforResult(sql, cn.getConnect());
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Global.Level = reader["ECIMAUTHLEVELNAME"].ToString().Trim();
                        }
                    }
                    else
                    {
                        cn.closeData();
                        Global.ConnectDB = Global.DLPConnectDB;
                        cn = new DBConnect();
                        reader = cn.queryforResult(sql, cn.getConnect());
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                Global.Level = reader["ECIMAUTHLEVELNAME"].ToString().Trim();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please check already register eCIM Account!");
                            this.Close();
                        }
                    }


                    if (Global.Level.Equals("Operator"))
                    {
                        MessageBox.Show(Global.UserID + " is " + Global.Level + " , Can not use.");
                        this.Close();
                    }
                    else
                    {
                        LoadCimRsc.RunWorkerAsync();
                    }
                    cn.closeData();
                }
                catch (Exception ex)
                {                   
                    throw new System.Exception("Please check DB Connection File \n In C:\\oracle\\ora92\\network\\ADMIN\\tnsnames.ora");
                }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not find UserID \n"+ex.Message.ToString());
                this.Close();
            }
         
        }

        private void getConfigProperties(string path,Config config)
        {
            try
            {
                FileStream ff;
                StreamReader reader;
                ff = File.Open(path, FileMode.Open);
                reader = new StreamReader(ff, Encoding.UTF8);
           
                while (!reader.EndOfStream)
                {
                    Encoder.Security.Encoder encode = new Encoder.Security.Encoder();
                    EncoderType type = EncoderType.TripleDES;
                    string value = encode.Decrypt(type, reader.ReadLine());

                    char e = '=';
                    string[] str = value.Split(e);

                    switch (str[0])
                    {
                        case "SiteID":
                            config.SiteID = str[1];
                            break;
                        case "FactoryID":
                            config.FactoryID = str[1];
                            break;
                        case "LineID":
                            config.LineID = str[1];
                            break;
                        case "ProductionLine":
                            config.ProductionLine = str[1];
                            break;
                        case "Environment":
                            config.Environment = str[1];
                            break;
                        case "ConfigServer":
                            config.ConfigServer = str[1];
                            break;
                        case "ConfigServerID":
                            config.ConfigServerID = str[1];
                            break;
                        case "ConfigServerPWD":
                            config.ConfigServerPWD = str[1];
                            break;
                        case "ConfigFolder":
                            config.ConfigFolder = str[1];
                            break;
                        case "ConfigFile":
                            config.ConfigFile = str[1];
                            break;
                        case "LocalConfigPath":
                            config.LocalConfigPath = str[1];
                            break;
                    }
                }
                reader.Dispose();
                reader.Close();               
            }
            catch (Exception ex)
            {
                reader.Dispose();
                reader.Close();               
            }           
        }

        public void getLineCode()
        {
            try
            {
                lsDivision.Clear();
                cboDivision.Items.Clear();

                DBConnect cn = new DBConnect();
                string sql = "select * from getLine order by productionlinename";
                reader = cn.queryforResult(sql, cn.getConnect());

                while (reader.Read())
                {
                    lsDivision.Add(reader["PRODUCTIONLINENAME"].ToString());
                    cboDivision.Items.Add(reader["PRODUCTIONLINENAME"].ToString());
                }

                cn.closeData();

                //var tmp = from n in lsCimRsc
                //          orderby n.LineCode ascending
                //          select n;

                //foreach (var rsc in tmp)
                //{
                //    if (!lsDivision.Contains(rsc.LineCode))
                //    {
                //        lsDivision.Add(rsc.LineCode);
                //        cboDivision.Items.Add(rsc.LineCode);
                //    }
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("Get Division Error");
            }
           
        }

        private void Add_TabPage(string v, ViewLimitGroup view)
        {
            tabControl2.TabPages.Add(v);
            tabControl2.SelectTab(tabControl2.TabPages.Count - 1);
            //dicView.Add(tcTab2.SelectedIndex,view.Division + "_" + view.Operation + "_" + view.Model + "_" + view.Eqpno + "_" + view.Recipe);
            view.Name = v;
            view.FormBorderStyle = FormBorderStyle.None;
            view.Dock = DockStyle.Fill;
            view.TopLevel = false;           
            view.Parent = tabControl2.SelectedTab;
            view.Show();
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            screenFdcList();
        }

        public void OpenForCopy(FDCMater vo)
        {
            getCollectPararms();

            ValidateInput();

            int PageCount = tabControl2.TabPages.Count;
            for (int i = 0; i < PageCount; i++)
            {
                tabControl2.TabPages.RemoveAt(0);

            }

            //dicView.Clear();

            //if (lsCollect.Count > 0)
            //{
                List<FDCMater> ls = new List<FDCMater>();

                vo.ApprovedBy = "";
                vo.ApprovedTime = "";
                vo.CreatedBy = "";
                vo.CreatedTime = "";
                vo.SubmittedBy = "";
                vo.SubmittedTime = "";
                vo.EditedBy = "";
                vo.EditedTime = "";
                ls.Add(vo);
             ViewLimitGroup view = new ViewLimitGroup(ls,true, vo.Division, vo.OperCode + (char)32 + vo.OperCode, vo.MachModel, vo.MachId, sRecipe,vo.Package, "Register",lsEqpNo, false,lsCollect, this);
            //ViewLimitGroup view = new ViewLimitGroup(ls,true,vo.Division, vo.OperName + (char)32 + vo.OperName, vo.MachModel, vo.MachId, "Register", lsEqpNo, lsCollect, this);
                Add_TabPage("Register limit group", view);
            //}
            //else
            //{
            //    MessageBox.Show("Please set Parameter First.");
            //}
        }

        public void screenFdcList()
        {
            getCollectPararms();
            getAvailableParms();
            try
            {
                lsStatus.Clear();
                lsFDCmater.Clear();
                if(ApprovedView.Rows.Count>0) ApprovedView.DataSource=null;
                if (SubmittedView.Rows.Count > 0) SubmittedView.DataSource = null;
                if (DraftView.Rows.Count > 0) DraftView.DataSource = null;
                if (DeletedView.Rows.Count > 0) DeletedView.DataSource = null;
                getLimitDefineData();
                
                foreach (FDCMater n in lsFDCmater)
                {
                    if (!lsStatus.Contains(n.Status))
                    {
                        lsStatus.Add(n.Status);
                    }
                }

                for (int i = 0; i < lsStatus.Count; i++)
                {
                    setDataGridView(lsStatus[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error");

            }
        }

        private void setDataGridView(string v)
        {
            try
            {
                var tmp = from n in lsFDCmater
                          where n.Status == v
                          select n;

                if (v.IndexOf("Submitted") >-1 || v.IndexOf("Rejected") >-1)
                {
                    tmp = from n in lsFDCmater
                          where n.Status == "Rejected" || n.Status == "Submitted"
                          select n;
                }

                lbsearchtime.Text = "Searching result：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                switch (v)
                {
                    case "Approved":

                        DataTable Approved = new DataTable(v);
                        
                        //Approved.Columns.Add("select", typeof(System.Windows.Forms.Button));
                        Approved.Columns.Add("STATUS", typeof(string));

                        Approved.Columns.Add("Operation", typeof(string));
                        Approved.Columns.Add("Equip model", typeof(string));
                        Approved.Columns.Add("Equip No", typeof(string));
                        Approved.Columns.Add("Recipe Name", typeof(string));

                        Approved.Columns.Add("Package", typeof(string));

                        Approved.Columns.Add("Created Time", typeof(string));
                        Approved.Columns.Add("Created by", typeof(string));
                        Approved.Columns.Add("Edited Time", typeof(string));
                        Approved.Columns.Add("Edited by", typeof(string));
                        Approved.Columns.Add("Submitted Time", typeof(string));
                        Approved.Columns.Add("Submitted by", typeof(string));
                        Approved.Columns.Add("Approved Time", typeof(string));
                        Approved.Columns.Add("Approved by", typeof(string));
                        Approved.Columns.Add("Customer", typeof(string));
                       
                        Approved.Columns.Add("Dimension", typeof(string));
                        Approved.Columns.Add("Lead", typeof(string));
                        Approved.Columns.Add("Device", typeof(string));
                        Approved.Columns.Add("Bonding number(Revision)", typeof(string));
                        Approved.Columns.Add("Division", typeof(string));
                        foreach (var n in tmp)
                        {
                            Approved.Rows.Add(new string[] {n.Status, n.OperCode, n.MachModel, n.MachId, n.RecipeName, n.Package, n.CreatedTime, n.CreatedBy, n.EditedTime, n.EditedBy, n.SubmittedTime, n.SubmittedBy, n.ApprovedTime, n.ApprovedBy, n.Customer, n.Dimension, n.Lead, n.Device, n.BondingRev  ,n.Division });
                        }
                        ApprovedView.MultiSelect = false;
                        ApprovedView.DataSource = Approved;
                        ApprovedView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        this.ApprovedView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                        break;
                    case "Deleted":

                        DataTable Deleted = new DataTable(v);

                        //Approved.Columns.Add("select", typeof(System.Windows.Forms.Button));
                        Deleted.Columns.Add("STATUS", typeof(string));
                     
                        Deleted.Columns.Add("Operation", typeof(string));
                        Deleted.Columns.Add("Equip model", typeof(string));
                        Deleted.Columns.Add("Equip No", typeof(string));
                        Deleted.Columns.Add("Recipe Name", typeof(string));

                        Deleted.Columns.Add("Package", typeof(string));

                        Deleted.Columns.Add("Created Time", typeof(string));
                        Deleted.Columns.Add("Created by", typeof(string));
                        Deleted.Columns.Add("Edited Time", typeof(string));
                        Deleted.Columns.Add("Edited by", typeof(string));
                        Deleted.Columns.Add("Submitted Time", typeof(string));
                        Deleted.Columns.Add("Submitted by", typeof(string));
                        Deleted.Columns.Add("Approved Time", typeof(string));
                        Deleted.Columns.Add("Approved by", typeof(string));
                        Deleted.Columns.Add("Customer", typeof(string));
                        
                        Deleted.Columns.Add("Dimension", typeof(string));
                        Deleted.Columns.Add("Lead", typeof(string));
                        Deleted.Columns.Add("Device", typeof(string));
                        Deleted.Columns.Add("Bonding number(Revision)", typeof(string));
                        Deleted.Columns.Add("Division", typeof(string));
                        foreach (var n in tmp)
                        {
                            Deleted.Rows.Add(new string[] { n.Status, n.OperCode, n.MachModel, n.MachId, n.RecipeName, n.Package, n.CreatedTime, n.CreatedBy, n.EditedTime, n.EditedBy, n.SubmittedTime, n.SubmittedBy, n.ApprovedTime, n.ApprovedBy, n.Customer, n.Dimension, n.Lead, n.Device, n.BondingRev, n.Division });
                        }
                        DeletedView.MultiSelect = false;
                        DeletedView.DataSource = Deleted;
                        DeletedView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        this.DeletedView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                        break;
                    case "Draft":

                        DataTable Draft = new DataTable(v);

                        Draft.Columns.Add("STATUS", typeof(string));
                       
                        Draft.Columns.Add("Operation", typeof(string));
                        Draft.Columns.Add("Equip model", typeof(string));
                        Draft.Columns.Add("Equip No", typeof(string));
                        Draft.Columns.Add("Recipe Name", typeof(string));

                        Draft.Columns.Add("Package", typeof(string));

                        Draft.Columns.Add("Created Time", typeof(string));
                        Draft.Columns.Add("Created by", typeof(string));
                        Draft.Columns.Add("Edited Time", typeof(string));
                        Draft.Columns.Add("Edited by", typeof(string));
                        Draft.Columns.Add("Submitted Time", typeof(string));
                        Draft.Columns.Add("Submitted by", typeof(string));
                        Draft.Columns.Add("Approved Time", typeof(string));
                        Draft.Columns.Add("Approved by", typeof(string));
                        Draft.Columns.Add("Customer", typeof(string));
                        
                        Draft.Columns.Add("Dimension", typeof(string));
                        Draft.Columns.Add("Lead", typeof(string));
                        Draft.Columns.Add("Device", typeof(string));
                        Draft.Columns.Add("Bonding number(Revision)", typeof(string));
                        Draft.Columns.Add("Division", typeof(string));
                        foreach (var n in tmp)
                        {
                            Draft.Rows.Add(new string[] { n.Status, n.OperCode, n.MachModel, n.MachId, n.RecipeName, n.Package, n.CreatedTime, n.CreatedBy, n.EditedTime, n.EditedBy, n.SubmittedTime, n.SubmittedBy, n.ApprovedTime, n.ApprovedBy, n.Customer, n.Dimension, n.Lead, n.Device, n.BondingRev, n.Division });
                        }

                        DraftView.MultiSelect = false;
                        DraftView.DataSource = Draft;
                        DraftView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        this.DraftView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                        //for (int i = 0; i < DraftView.ColumnCount - 1; i++)
                        //{
                        //    this.DraftView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                        //}
                        break;
                    case "Rejected":
                        goto case "Submitted";
                    case "Submitted":

                        DataTable Submitted = new DataTable(v);

                        Submitted.Columns.Add("STATUS", typeof(string));
                       
                        Submitted.Columns.Add("Operation", typeof(string));
                        Submitted.Columns.Add("Equip model", typeof(string));
                        Submitted.Columns.Add("Equip No", typeof(string));
                        Submitted.Columns.Add("Recipe Name", typeof(string));

                        Submitted.Columns.Add("Package", typeof(string));

                        Submitted.Columns.Add("Created Time", typeof(string));
                        Submitted.Columns.Add("Created by", typeof(string));
                        Submitted.Columns.Add("Edited Time", typeof(string));
                        Submitted.Columns.Add("Edited by", typeof(string));
                        Submitted.Columns.Add("Submitted Time", typeof(string));
                        Submitted.Columns.Add("Submitted by", typeof(string));
                        Submitted.Columns.Add("Approved Time", typeof(string));
                        Submitted.Columns.Add("Approved by", typeof(string));
                        Submitted.Columns.Add("Customer", typeof(string));
                        
                        Submitted.Columns.Add("Dimension", typeof(string));
                        Submitted.Columns.Add("Lead", typeof(string));
                        Submitted.Columns.Add("Device", typeof(string));
                        Submitted.Columns.Add("Bonding number(Revision)", typeof(string));
                        Submitted.Columns.Add("Division", typeof(string));

                        foreach (var n in tmp)
                        {                           
                            Submitted.Rows.Add(new string[] { n.Status, n.OperCode, n.MachModel, n.MachId, n.RecipeName, n.Package, n.CreatedTime, n.CreatedBy, n.EditedTime, n.EditedBy, n.SubmittedTime, n.SubmittedBy, n.ApprovedTime, n.ApprovedBy, n.Customer, n.Dimension, n.Lead, n.Device, n.BondingRev, n.Division });
                        }
                        SubmittedView.MultiSelect = false;
                        //adt.Fill(Submitted);
                        SubmittedView.DataSource = Submitted;
                        SubmittedView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        this.SubmittedView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                        //for (int i = 0; i < SubmittedView.ColumnCount - 1; i++)
                        //{
                        //    this.SubmittedView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                        //}
                        break;
                }             
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        private void getLimitDefineData()
        {
            try
            {
                DBConnect DBConnect = new DBConnect();
                str = ValidateSQL();
                //str += "Order by createdtime";
                reader = DBConnect.queryforResult(str, DBConnect.getConnect());
                lsFDCmater.Clear();
                while (reader.Read())
                {
                    if (Global.passEFGP)
                    {
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
                            //EFGPSN = reader["EFGPSN"].ToString(),
                        };
                        lsFDCmater.Add(FDCmaster);
                    }
                    else
                    {
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
                        lsFDCmater.Add(FDCmaster);
                    }
                    
                    
                }
                DBConnect.closeData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Get LimitDefine Error");

            }
          
        }

        public string ValidateSQL()
        {
            string s = "select * from ATT_ECIM.LimitDefinitionMaster where Division='"+sDivision + "'";
            if (sOperation!="")
            {
                s += " and OperCode='" + sOperation + "'";
            }
             if (sEqpModel!="")
            {
                s+= " and MachModel='" + sEqpModel + "'";
            }
             if (sEqpNo!="")
            {
                s += " and MachId='" + sEqpNo + "'";
            }

            return s;
        }

        private void cboDivision_SelectedIndexChanged(object sender, EventArgs e)
        {           
            cboOperation.Items.Clear();
            cboOperation.Text = "";

            cboEquipModel.Items.Clear();
            cboEquipModel.Text = "";

            cboEquipNO.Items.Clear();


            sDivision = cboDivision.Text;

            setString();
            getOperation();

            //if (!sDivision.StartsWith("T3"))
            if (!sDivision.Contains("DPS"))
            {
                cboEquipNO.Enabled = true;
            }
            else
            {
                cboEquipNO.Enabled = false;
            }
        }

        private void setString()
        {
            sDivision = cboDivision.Text;
            sOperation = cboOperation.Text;
            sEqpModel = cboEquipModel.Text;
            sEqpNo = cboEquipNO.Text;
        }

        private void LoadCimRsc_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            DBConnect DBConnect = new DBConnect();
            str = "select * from Getcimrsc";
            reader = DBConnect.queryforResult(str, DBConnect.getConnect());
            while (reader.Read())
            {
                cimrsc rsc = new cimrsc { Factory = reader["RSFACD"].ToString(), LineCode = reader["RLLINE"].ToString(), Operation = reader["RSTYPE"].ToString(), Model = reader["RSSTYP"].ToString(), EqpNo = reader["RSMCH1"].ToString() };
                lsCimRsc.Add(rsc);
                worker.ReportProgress(1);
            }
            worker.ReportProgress(2);
            DBConnect.closeData();
        }

        private void LoadCimRsc_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            getLineCode();
            cboDivision.Text = lsDivision[0];
        }

        private void LoadCimRsc_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage.Equals("1"))
            {
                toolStripStatusLabel1.ForeColor = Color.Red;
                toolStripStatusLabel1.Text = "Loading...";
            }
            else
            {
                toolStripStatusLabel1.ForeColor = Color.Green;

                //get version
                try
                {
                    toolStripStatusLabel1.Text = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(); //Environment.Version.ToString() ; //"Loading Complete";
                }
                catch (Exception)
                {

                }

                toolStripStatusLabel2.ForeColor = Color.Black;
                toolStripStatusLabel2.Text = Global.UserID + " | " + Global.UserID + " | " + Global.Level;
            }
            
        }

        private void cboEquipModel_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboEquipNO.Items.Clear();
            cboEquipNO.Text = "";

            sEqpModel = cboEquipModel.Text;

            setString();
            getEqpNo();
        }

        public void getEqpNo()
        {
            cboEquipNO.Items.Clear();
            lsEqpNo.Clear();
            //cboEquipNO.Items.Add("");

            ////change T1_BP_FLUX_CLEANER to T1_BP_FLUX CLEANER for query EqpNo
            //if (sOperation == "T1_BP_FLUX_CLEANER")
            //{
            //    sOperation = "T1_BP_FLUX CLEANER";
            //}

            string tmp = "";
            try
            {
                if (dicEquipModel[sDivision + " " + sEqpModel] != null)
                {
                    tmp = dicEquipModel[sDivision + " " + sEqpModel];
                    string[] array = tmp.Split(':');
                    foreach (var item in array)
                    {
                        //add for not add space 20161228
                        if (!item.ToString().Equals(""))
                        {
                            lsEqpNo.Add(item);
                            cboEquipNO.Items.Add(item);
                        }                        
                    }
                }
            }
            catch (Exception)
            {

                DBConnect cn = new DBConnect();
                string sql = "select * from getcimrsc_ecim where RSSTYP='" + sEqpModel + "' and RLLINE='" + sDivision + "'";
                reader = cn.queryforResult(sql, cn.getConnect());

                while (reader.Read())
                {
                    string str = reader["RSMCH1"].ToString();
                    tmp += ":" + str;
                    lsEqpNo.Add(str);
                    cboEquipNO.Items.Add(str);
                }
                cn.closeData();
                dicEquipModel.Add(sDivision + " " + sEqpModel, tmp);
            }
            //finally
            //{
            //    if (sOperation == "T1_BP_FLUX CLEANER")
            //    {
            //        sOperation = "T1_BP_FLUX_CLEANER";
            //    }
            //}
            //var tmp = from n in lsCimRsc
            //          where n.LineCode.Equals(sDivision) && n.Operation.Equals(sOperation) && n.Model.Equals(sEqpModel)
            //          select n;
            //foreach (var item in tmp)
            //{
            //    if (!lsEqpNo.Contains(item.EqpNo))
            //    {
            //        cboEquipNO.Items.Add(item.EqpNo);
            //        lsEqpNo.Add(item.EqpNo);
            //    }
            //}
        }

        private void cboEquipNO_SelectedIndexChanged(object sender, EventArgs e)
        {
            sEqpNo = cboEquipNO.Text;
            setString();
        }

        private void tabControl2_DrawItem(object sender, DrawItemEventArgs e)
        {
            /*如果将 DrawMode 属性设置为 OwnerDrawFixed， 
          则每当 TabControl 需要绘制它的一个选项卡时，它就会引发 DrawItem 事件*/
            try
            {
                this.tabControl2.TabPages[e.Index].BackColor = Color.LightBlue;
                Rectangle tabRect = this.tabControl2.GetTabRect(e.Index);
                
                //if (this.tabControl2.TabCount > 1)
                //{
                //    tabRect = this.tabControl2.GetTabRect(tcTab2.TabCount);
                //    e.Graphics.DrawString(this.tabControl2.TabPages[tcTab2.TabCount].Text, this.Font, SystemBrushes.ControlText, (float)(tabRect.X + 2), (float)(tabRect.Y + 2));
                //}
                //else
                //{
                    e.Graphics.DrawString(this.tabControl2.TabPages[e.Index].Text, this.Font, SystemBrushes.ControlText, (float)(tabRect.X + 2), (float)(tabRect.Y + 2));
                //}

                

                using (Pen pen = new Pen(Color.White))
                {
                    tabRect.Offset(tabRect.Width - 15, 2);
                    tabRect.Width = 15;
                    tabRect.Height = 15;
                    e.Graphics.DrawRectangle(pen, tabRect);
                }
                Color color = (e.State == DrawItemState.Selected) ? Color.LightBlue : Color.White;
                using (Brush brush = new SolidBrush(color))
                {
                    e.Graphics.FillRectangle(brush, tabRect);
                }
                using (Pen pen2 = new Pen(Color.Red))
                {
                    Point point = new Point(tabRect.X + 3, tabRect.Y + 3);
                    Point point2 = new Point((tabRect.X + tabRect.Width) - 3, (tabRect.Y + tabRect.Height) - 3);
                    e.Graphics.DrawLine(pen2, point, point2);
                    Point point3 = new Point(tabRect.X + 3, (tabRect.Y + tabRect.Height) - 3);
                    Point point4 = new Point((tabRect.X + tabRect.Width) - 3, tabRect.Y + 3);
                    e.Graphics.DrawLine(pen2, point3, point4);
                }
                e.Graphics.Dispose();
                //tabControl2.SelectedIndex = tabControl2.TabCount;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void parameterSettingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            getCollectPararms();
            getAvailableParms();
            if (sOperation =="")
            {
                MessageBox.Show("Operation must be selected");
            }
            else if (sEqpModel == "")
            {
                MessageBox.Show("Equip Model must be selected");
            }else
            {
                ParameterSetting PS = new ParameterSetting(sDivision, sOperation, sEqpModel, sEqpNo,lsParameter,lsCollect);
                PS.ShowDialog();
                lsCollect = PS.lsCollectParameter;
            }

        }

        public void getAvailableParms()
        {
            lsParameter.Clear();
            try
            {
                DBConnect con = new DBConnect();               
                str = "select * from ATT_ECIM.Parameters where MachModel='" + sEqpModel + "'";

                if (!sEqpNo.Equals(""))
                {
                    str += " and MachId ='" + sEqpNo + "'";
                }
                else
                {
                    str += " and MachId is null ";
                }

                str += " ORDER BY PARAMNAME ";

                reader = con.queryforResult(str, con.getConnect());

                while (reader.Read())
                {
                    setParameter(reader);
                }
                con.closeData();
            }
            catch (Exception ex)
            {

            }

        }

        private void setParameter(OleDbDataReader reader)
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

        public void getCollectPararms()
        {
            lsCollect.Clear();
            DBConnect con = new DBConnect();
            try
            {                          
                str = "select * from ATT_ECIM.PARAMETERSPERDIV where Division='" + sDivision + "' and OperCode='" + sOperation + "' and MachModel='" + sEqpModel + "'";

                if (!sEqpNo.Equals(""))
                {
                    str += " and MachId ='" + sEqpNo + "'";
                }
                else
                {
                    str += " and MachId is null ";
                }

                str += " ORDER BY PARAMNAME ";

                reader = con.queryforResult(str, con.getConnect());

                while (reader.Read())
                {
                    setCollect(reader);
                }
            }
            catch (Exception ex)
            {
                
            }
            finally
            {
                con.closeData();
            }
        }

        private void setCollect(OleDbDataReader reader)
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
            lsCollect.Add(P);
        }

        private void eMailSettingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            eMailSetting eMS = new eMailSetting(lsDivision);
            eMS.Show();
        }     

        public void RemoveTab2(int index)
        {
            tcTab2.TabPages.RemoveAt(index);
        }

        private void registerRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            getCollectPararms();

            if (ValidateInput())
            {
                int PageCount = tabControl2.TabPages.Count;
                for (int i = 0; i < PageCount; i++)
                {
                    tabControl2.TabPages.RemoveAt(0);

                }

                //dicView.Clear();

                if (lsCollect.Count > 0)
                {
                    ViewLimitGroup view = new ViewLimitGroup(sDivision, sOperation + (char)32 + sOperation, sEqpModel, sEqpNo, "Register", lsEqpNo, lsCollect, this);
                    Add_TabPage("Register limit group", view);
                }
                else
                {
                    MessageBox.Show("Please set Parameter First.");
                }
            }
           
        }

        private Boolean ValidateInput()
        {
            string msg = "";
            if (sDivision == "")
            {
                msg = "Please select Division";
            }
            else if (sOperation == "")
            {
                msg = "Please select Operation";
            }
            else if (sEqpModel == "")
            {
                msg = "Please select Equip Model";
            }
            else if (sEqpNo == "")
            {
                //if (!sDivision.Contains("T3"))//for T3 can pass Eqp Select 20180516
                if (!sDivision.Contains("DPS"))//for T6 use T3 rule
                {
                    msg = "Please select Equip No";
                }                
            }
            if (msg != "")
            {
                MessageBox.Show(msg);
                return false;
            }
            return true;
        }

        private void editEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                getCollectPararms();
                getAvailableParms();
                //ValidateInput();

                if (tcTabSelect)
                {
                    if (tcTab.SelectedTab.Text != "Deleted")
                    {
                        string Div = "", Oper = "", EqpModel = "", EqpNo = "";
                        switch (tcTab.SelectedTab.Text)
                        {
                            case "Approved":
                                Div = this.ApprovedView.Rows[tcTabIndex].Cells["Division"].Value.ToString();
                                Oper = this.ApprovedView.Rows[tcTabIndex].Cells["Operation"].Value.ToString();
                                EqpModel = this.ApprovedView.Rows[tcTabIndex].Cells["Equip model"].Value.ToString();
                                EqpNo = this.ApprovedView.Rows[tcTabIndex].Cells["Equip No"].Value.ToString();
                                sRecipe = this.ApprovedView.Rows[tcTabIndex].Cells["Recipe Name"].Value.ToString();
                                sPackage = this.ApprovedView.Rows[tcTabIndex].Cells["Package"].Value.ToString();
                                break;

                            case "Submitted":
                                Div = this.SubmittedView.Rows[tcTabIndex].Cells["Division"].Value.ToString();
                                Oper = this.SubmittedView.Rows[tcTabIndex].Cells["Operation"].Value.ToString();
                                EqpModel = this.SubmittedView.Rows[tcTabIndex].Cells["Equip model"].Value.ToString();
                                EqpNo = this.SubmittedView.Rows[tcTabIndex].Cells["Equip No"].Value.ToString();
                                sRecipe = this.SubmittedView.Rows[tcTabIndex].Cells["Recipe Name"].Value.ToString();
                                sPackage = this.SubmittedView.Rows[tcTabIndex].Cells["Package"].Value.ToString();
                                break;

                            case "Draft":
                                Div = this.DraftView.Rows[tcTabIndex].Cells["Division"].Value.ToString();
                                Oper = this.DraftView.Rows[tcTabIndex].Cells["Operation"].Value.ToString();
                                EqpModel = this.DraftView.Rows[tcTabIndex].Cells["Equip model"].Value.ToString();
                                EqpNo = this.DraftView.Rows[tcTabIndex].Cells["Equip No"].Value.ToString();
                                sRecipe = this.DraftView.Rows[tcTabIndex].Cells["Recipe Name"].Value.ToString();
                                sPackage = this.DraftView.Rows[tcTabIndex].Cells["Package"].Value.ToString();
                                break;
                        }
                        lsFDCmater.Clear();
                        getLimitDefineData();//lsFDCmater reset

                        var tmp = from n in lsFDCmater
                                  where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Status.Equals(tcTab.SelectedTab.Text) && n.Package == sPackage
                                  select n;
                        List<FDCMater> ls = new List<FDCMater>();
                        foreach (var n in tmp)
                        {
                            ls.Add(n);
                        }

                        if (ls.Count < 1)
                        {
                            tmp = from n in lsFDCmater
                                  where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Package == sPackage
                                  select n;
                            foreach (var n in tmp)
                            {
                                ls.Add(n);
                            }
                        }

                        int PageCount = tcTab2.TabCount;
                        int flag = 0;
                        for (int i = 0; i < PageCount; i++)
                        {
                            if (tcTab2.TabPages[flag].Text.IndexOf("View") > -1)
                            {
                                tcTab2.TabPages.RemoveAt(flag);
                            }
                            else
                            {
                                flag++;
                            }
                        }
                        bool CanRequest = true;
                        if (Div.IndexOf("T1")>-1)
                        {
                            if (!Global.passEFGP)//for pass check EFGP
                            {
                                if (tcTab.SelectedTab.Text.Equals("Submitted"))
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

                                    if (ls[0].EFGPSN != null)
                                    {//already request                                
                                        if (!ls[0].EFGPSN.Equals(""))
                                        {
                                            CanRequest = false;
                                            if ((EF.getProcessStatus(ls[0].EFGPSN).IndexOf("close") > -1))
                                            {
                                                CanRequest = true;
                                            }
                                        }
                                    }
                                }
                            }                            
                        }

                        if (CanRequest)
                        {
                            //例如，注意項目的取法
                            //ViewLimitGroup view = new ViewLimitGroup(sDivision, sOperation + (char)32 + sOperation, sEqpModel, sEqpNo, "Register", lsEqpNo, lsCollect, this);
                            ViewLimitGroup view = new ViewLimitGroup(ls, lsCollect, Div, Oper + (char)32 + Oper, EqpModel, EqpNo, sRecipe, sPackage, "Edit", lsEqpNo, false, this);

                            Add_TabPage("Edit limit group", view);
                            //MessageBox.Show(sRecipe);
                        }
                        else
                        {
                            //show message for wait approval
                            MessageBox.Show("This Setting already send to Approval,\n Please wait for Approval Finish.\n Easy Flow SerialNo : "+ls[0].EFGPSN+"", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Select parameter that will be Edit. ", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not connect to EasyFlow , \nPlease make sure EasyFlow Server is Available.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
        }

        private void deleteDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tcTabSelect)
            {
                if (tcTab.SelectedTab.Text != "Deleted")
                {
                    //if ()
                    //{

                    //}

                    string Div = "", Oper = "", EqpModel = "", EqpNo = "", STATUS = "";
                    switch (tcTab.SelectedTab.Text)
                    {
                        case "Approved":
                            Div = this.ApprovedView.Rows[tcTabIndex].Cells["Division"].Value.ToString();
                            Oper = this.ApprovedView.Rows[tcTabIndex].Cells["Operation"].Value.ToString();
                            EqpModel = this.ApprovedView.Rows[tcTabIndex].Cells["Equip model"].Value.ToString();
                            EqpNo = this.ApprovedView.Rows[tcTabIndex].Cells["Equip No"].Value.ToString();
                            sRecipe = this.ApprovedView.Rows[tcTabIndex].Cells["Recipe Name"].Value.ToString();
                            sPackage = this.ApprovedView.Rows[tcTabIndex].Cells["Package"].Value.ToString();
                            STATUS = this.ApprovedView.Rows[tcTabIndex].Cells["STATUS"].Value.ToString();
                            break;

                        case "Submitted":
                            Div = this.SubmittedView.Rows[tcTabIndex].Cells["Division"].Value.ToString();
                            Oper = this.SubmittedView.Rows[tcTabIndex].Cells["Operation"].Value.ToString();
                            EqpModel = this.SubmittedView.Rows[tcTabIndex].Cells["Equip model"].Value.ToString();
                            EqpNo = this.SubmittedView.Rows[tcTabIndex].Cells["Equip No"].Value.ToString();
                            sRecipe = this.SubmittedView.Rows[tcTabIndex].Cells["Recipe Name"].Value.ToString();
                            sPackage = this.SubmittedView.Rows[tcTabIndex].Cells["Package"].Value.ToString();
                            STATUS = this.SubmittedView.Rows[tcTabIndex].Cells["STATUS"].Value.ToString();
                            break;

                        case "Draft":
                            Div = this.DraftView.Rows[tcTabIndex].Cells["Division"].Value.ToString();
                            Oper = this.DraftView.Rows[tcTabIndex].Cells["Operation"].Value.ToString();
                            EqpModel = this.DraftView.Rows[tcTabIndex].Cells["Equip model"].Value.ToString();
                            EqpNo = this.DraftView.Rows[tcTabIndex].Cells["Equip No"].Value.ToString();
                            sRecipe = this.DraftView.Rows[tcTabIndex].Cells["Recipe Name"].Value.ToString();
                            sPackage = this.DraftView.Rows[tcTabIndex].Cells["Package"].Value.ToString();
                            STATUS = this.DraftView.Rows[tcTabIndex].Cells["STATUS"].Value.ToString();
                            break;
                    }
                    lsFDCmater.Clear();
                    getLimitDefineData();

                    var tmp = from n in lsFDCmater
                              where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Status.Equals(STATUS) && n.Package == sPackage
                              select n;
                    List<FDCMater> ls = new List<FDCMater>();
                    foreach (var n in tmp)
                    {
                        ls.Add(n);
                    }

                    bool CanRequest = true;
                    if (Div.IndexOf("T1") > -1)
                    {
                        if (!Global.passEFGP)
                        {
                            if (tcTab.SelectedTab.Text.Equals("Submitted"))
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

                                if (ls[0].EFGPSN != null)
                                {//already request                                
                                    if (!ls[0].EFGPSN.Equals(""))
                                    {
                                        CanRequest = false;
                                        if ((EF.getProcessStatus(ls[0].EFGPSN).IndexOf("close") > -1))
                                        {
                                            CanRequest = true;
                                        }
                                    }
                                }
                            }
                        }                        
                    }

                    if (CanRequest)
                    {   //例如，注意項目的取法
                        ViewLimitGroup view = new ViewLimitGroup(ls, Div, Oper, EqpModel, EqpNo, sRecipe, sPackage, "Deleted", false, this);
                        //if (tcTab2.TabCount >0 && tcTab2.SelectedTab.Text!= "Delete  limit group")
                        //{
                        //    tcTab2.TabPages.RemoveAt(tcTab2.SelectedIndex);
                        //}                    

                        Add_TabPage("Delete limit group", view);
                        //MessageBox.Show(sRecipe);
                    }
                    else
                    {
                        //show message for wait approval
                        MessageBox.Show("This Setting already send to Approval,\n Please wait for Approval Finish.\n Easy Flow SerialNo : " + ls[0].EFGPSN + "", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Select parameter that will be deleted. ", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
          
        }
        

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tabControl2_ControlRemoved(object sender, ControlEventArgs e)
        {
            if (tcTab2.SelectedIndex == 0 && tcTab2.TabCount.Equals(1))
            {
                tcTab2.SelectedIndex = -1; 
            }
            else if (tcTab2.SelectedIndex == tcTab2.TabCount - 1)            
            {
                tabControl2.SelectTab(tcTab2.TabCount - 2);
            }
            else
            {
                tabControl2.SelectTab(tcTab2.TabCount - 1);
            }
            
            //MessageBox.Show("After \n SelectedIndex = "+tcTab2.SelectedIndex +"\n Count = "+ tcTab2.TabCount);
        }

        private void SubmittedView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                tcTabIndex = e.RowIndex;
                tcTabSelect = true;
                if (e.ColumnIndex == this.SubmittedView.Columns["Select1"].Index)
                {
                    string Div = "", Oper = "", EqpModel = "", EqpNo = "", CreateTime = "", sStatus = "";
                    //例如，注意項目的取法
                    Div = this.SubmittedView.Rows[e.RowIndex].Cells["Division"].Value.ToString();
                    Oper = this.SubmittedView.Rows[e.RowIndex].Cells["Operation"].Value.ToString();
                    EqpModel = this.SubmittedView.Rows[e.RowIndex].Cells["Equip model"].Value.ToString();
                    EqpNo = this.SubmittedView.Rows[e.RowIndex].Cells["Equip No"].Value.ToString();
                    sRecipe = this.SubmittedView.Rows[e.RowIndex].Cells["Recipe Name"].Value.ToString();
                    sPackage = this.SubmittedView.Rows[e.RowIndex].Cells["Package"].Value.ToString();
                    CreateTime = this.SubmittedView.Rows[e.RowIndex].Cells["Created Time"].Value.ToString();
                    sStatus = this.SubmittedView.Rows[e.RowIndex].Cells["STATUS"].Value.ToString();
                    getLimitDefineData();
                    var tmp = from n in lsFDCmater
                              where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Status.Equals(sStatus) && n.Package == sPackage && n.CreatedTime == CreateTime
                              select n;
                    List<FDCMater> ls = new List<FDCMater>();
                    foreach (var n in tmp)
                    {
                        ls.Add(n);
                    }

                    ViewLimitGroup view = new ViewLimitGroup(ls, Div, Oper, EqpModel, EqpNo, sRecipe, sPackage ,"Submitted", false, this);

                    Add_TabPage("Approve limit group", view);
                    //MessageBox.Show(sRecipe);
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void ApprovedView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                tcTabIndex = e.RowIndex;
                tcTabSelect = true;
                if (e.ColumnIndex == this.ApprovedView.Columns["Select"].Index)
                {
                    string Div = "", Oper = "", EqpModel = "", EqpNo = "", CreateTime = "", sStatus = "";
                    //例如，注意項目的取法
                    Div = this.ApprovedView.Rows[e.RowIndex].Cells["Division"].Value.ToString();
                    Oper = this.ApprovedView.Rows[e.RowIndex].Cells["Operation"].Value.ToString();
                    EqpModel = this.ApprovedView.Rows[e.RowIndex].Cells["Equip model"].Value.ToString();
                    EqpNo = this.ApprovedView.Rows[e.RowIndex].Cells["Equip No"].Value.ToString();
                    sRecipe = this.ApprovedView.Rows[e.RowIndex].Cells["Recipe Name"].Value.ToString();
                    sPackage = this.ApprovedView.Rows[e.RowIndex].Cells["Package"].Value.ToString();
                    CreateTime = this.ApprovedView.Rows[e.RowIndex].Cells["Created Time"].Value.ToString();
                    sStatus = this.ApprovedView.Rows[e.RowIndex].Cells["STATUS"].Value.ToString();
                    getLimitDefineData();
                    var tmp = from n in lsFDCmater
                              where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Status.Equals(sStatus) && n.Package == sPackage && n.CreatedTime == CreateTime
                              select n;
                    List<FDCMater> ls = new List<FDCMater>();
                    foreach (var n in tmp)
                    {
                        ls.Add(n);
                    }

                    ViewLimitGroup view = new ViewLimitGroup(ls, Div, Oper, EqpModel, EqpNo, sRecipe, sPackage , "Approved", true, this);
                    Add_TabPage("View limit group", view);
                    //MessageBox.Show(sRecipe);
                }
            }
            catch (Exception ex)
            {
                
            }
           
        }

        private void DraftView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                tcTabIndex = e.RowIndex;
                tcTabSelect = true;
                if (e.ColumnIndex == this.DraftView.Columns["Select2"].Index)
                {
                    string Div = "", Oper = "", EqpModel = "", EqpNo = "", CreateTime = "", sStatus = "";
                    //例如，注意項目的取法
                    Div = this.DraftView.Rows[e.RowIndex].Cells["Division"].Value.ToString();
                    Oper = this.DraftView.Rows[e.RowIndex].Cells["Operation"].Value.ToString();
                    EqpModel = this.DraftView.Rows[e.RowIndex].Cells["Equip model"].Value.ToString();
                    EqpNo = this.DraftView.Rows[e.RowIndex].Cells["Equip No"].Value.ToString();
                    sRecipe = this.DraftView.Rows[e.RowIndex].Cells["Recipe Name"].Value.ToString();
                    sPackage = this.DraftView.Rows[e.RowIndex].Cells["Package"].Value.ToString();
                    CreateTime = this.DraftView.Rows[e.RowIndex].Cells["Created Time"].Value.ToString();
                    sStatus = this.DraftView.Rows[e.RowIndex].Cells["STATUS"].Value.ToString();
                    getLimitDefineData();
                    var tmp = from n in lsFDCmater
                              where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Status.Equals(sStatus) && n.Package == sPackage && n.CreatedTime == CreateTime
                              select n;
                    List<FDCMater> ls = new List<FDCMater>();
                    foreach (var n in tmp)
                    {
                        ls.Add(n);
                    }

                    //ViewLimitGroup view = new ViewLimitGroup(ls, Div, Oper, EqpModel, EqpNo, sRecipe, "Draft", false, this);
                    ViewLimitGroup view = new ViewLimitGroup(ls, lsCollect, Div, Oper + (char)32 + Oper, EqpModel, EqpNo, sRecipe, sPackage, "Draft",lsEqpNo, false, this);            
                    Add_TabPage("Approve limit group", view);
                }
            }
            catch (Exception ex)
            {

            }
          
        }

        private void DeletedView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                tcTabIndex = e.RowIndex;
                tcTabSelect = true;
                if (e.ColumnIndex == this.DeletedView.Columns["Select3"].Index)
                {
                    string Div = "", Oper = "", EqpModel = "", EqpNo = "", CreateTime = "", sStatus = "";
                    //例如，注意項目的取法
                    Div = this.DeletedView.Rows[e.RowIndex].Cells["Division"].Value.ToString();
                    Oper = this.DeletedView.Rows[e.RowIndex].Cells["Operation"].Value.ToString();
                    EqpModel = this.DeletedView.Rows[e.RowIndex].Cells["Equip model"].Value.ToString();
                    EqpNo = this.DeletedView.Rows[e.RowIndex].Cells["Equip No"].Value.ToString();
                    sRecipe = this.DeletedView.Rows[e.RowIndex].Cells["Recipe Name"].Value.ToString();
                    sPackage = this.DeletedView.Rows[e.RowIndex].Cells["Package"].Value.ToString();
                    CreateTime = this.DeletedView.Rows[e.RowIndex].Cells["Created Time"].Value.ToString();
                    sStatus = this.DeletedView.Rows[e.RowIndex].Cells["STATUS"].Value.ToString();
                    getLimitDefineData();
                    var tmp = from n in lsFDCmater
                              where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Status.Equals(sStatus) && n.Package == sPackage && n.CreatedTime == CreateTime
                              select n;
                    List<FDCMater> ls = new List<FDCMater>();
                    foreach (var n in tmp)
                    {
                        ls.Add(n);
                    }

                    ViewLimitGroup view = new ViewLimitGroup(ls, Div, Oper, EqpModel, EqpNo, sRecipe, sPackage , "Deleted", true, this);
                    Add_TabPage("View limit group", view);
                }
            }
            catch (Exception ex)
            {
                
            }
          
        }

        private void DeletedView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string Div = "", Oper = "", EqpModel = "", EqpNo = "", CreateTime ="", sStatus = "";
                //例如，注意項目的取法
                Div = this.DeletedView.Rows[e.RowIndex].Cells["Division"].Value.ToString();
                Oper = this.DeletedView.Rows[e.RowIndex].Cells["Operation"].Value.ToString();
                EqpModel = this.DeletedView.Rows[e.RowIndex].Cells["Equip model"].Value.ToString();
                EqpNo = this.DeletedView.Rows[e.RowIndex].Cells["Equip No"].Value.ToString();
                sRecipe = this.DeletedView.Rows[e.RowIndex].Cells["Recipe Name"].Value.ToString();
                sPackage = this.DeletedView.Rows[e.RowIndex].Cells["Package"].Value.ToString();
                CreateTime = this.DeletedView.Rows[e.RowIndex].Cells["Created Time"].Value.ToString();
                sStatus = this.DeletedView.Rows[e.RowIndex].Cells["STATUS"].Value.ToString();
                getLimitDefineData();
                var tmp = from n in lsFDCmater
                          where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Status.Equals(sStatus) && n.Package == sPackage && n.CreatedTime == CreateTime
                          select n;
                List<FDCMater> ls = new List<FDCMater>();
                foreach (var n in tmp)
                {
                    ls.Add(n);
                }

                ViewLimitGroup view = new ViewLimitGroup(ls, Div, Oper, EqpModel, EqpNo, sRecipe,sPackage, "Deleted", true, this);
                Add_TabPage("View limit group", view);
            }
            catch (Exception ex)
            {
                
            }
          
        }

        private void DraftView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string Div = "", Oper = "", EqpModel = "", EqpNo = "", CreateTime = "";
                //例如，注意項目的取法
                Div = this.DraftView.Rows[e.RowIndex].Cells["Division"].Value.ToString();
                Oper = this.DraftView.Rows[e.RowIndex].Cells["Operation"].Value.ToString();
                EqpModel = this.DraftView.Rows[e.RowIndex].Cells["Equip model"].Value.ToString();
                EqpNo = this.DraftView.Rows[e.RowIndex].Cells["Equip No"].Value.ToString();
                sRecipe = this.DraftView.Rows[e.RowIndex].Cells["Recipe Name"].Value.ToString();
                sPackage = this.DraftView.Rows[e.RowIndex].Cells["Package"].Value.ToString();
                CreateTime = this.DraftView.Rows[e.RowIndex].Cells["Created Time"].Value.ToString();
                getLimitDefineData();
                var tmp = from n in lsFDCmater
                          where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Status.Equals(tcTab.SelectedTab.Text) && n.Package == sPackage && n.CreatedTime == CreateTime
                          select n;
                List<FDCMater> ls = new List<FDCMater>();
                foreach (var n in tmp)
                {
                    ls.Add(n);
                }

                ViewLimitGroup view = new ViewLimitGroup(ls, lsCollect, Div, Oper + (char)32 + Oper, EqpModel, EqpNo, sRecipe, sPackage, "Draft", lsEqpNo, false, this);
                //ViewLimitGroup view = new ViewLimitGroup(ls, Div, Oper, EqpModel, EqpNo, sRecipe, sPackage , "Draft", false, this);

                Add_TabPage("Approve limit group", view);
            }
            catch (Exception ex)
            {
                
            }
           
        }

        private void exportToCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                toolStripStatusLabel3.Text = "";
                string internalSQL = "select m.serialno from ATT_ECIM.LimitDefinitionMaster m where m.division = '"+cboDivision.Text+ "' and (m.status = 'Approved' or m.status = 'Submitted')";
                if (cboOperation.Text != "") internalSQL += " and m.opercode = '" + cboOperation.Text + "'";
                if (cboEquipModel.Text != "") internalSQL += " and m.machmodel = '" + cboEquipModel.Text + "'";
                if (cboEquipNO.Text != "") internalSQL += " and m.machid = '" + cboEquipNO.Text + "'";

                string sql = "select MT.Status,MT.Division,MT.Opercode,MT.Machmodel,MT.Machid,MT.Recipename,p.parametername AS Parameter_name,p.secondflag AS Second_flag,p.stepname as Step_name,p.stepvalue as Step_value,p.eventid as Event_id,p.warmuptime as Warm_up_time,p.monitortime as Monitor_time,p.startdetectvalue as start_detect_value,p.checkmin as use_lower_limit,p.allowminequal as permit_equal_lower,p.minvalue as lower_limit,p.checkmax as use_upper_limit,p.allowmaxequal as permit_equal_upper,p.maxvalue as upper_limit,p.rchartflag as Rchart_flag,p.checkrmin as use_rchart_lower_limit,p.rminvalue as rchart_lower_limit,p.checkrmax as use_rchart_upper_limit,p.rmaxvalue as rchart_upper_limit,p.terminalmsg as terminal_msg,p.email as email,p.emailaddrs as email_list,p.mcinhibit as MC_inhibit,p.mchold as MC_hold,p.eocaptemplete as eocap_templete,p.eocap as eocap,p.lothold AS Lot_hold from ATT_ECIM.LimitDefinitionParameter p,ATT_ECIM.LimitDefinitionMaster MT where p.serialno in (" + internalSQL + ") and MT.Serialno = p.serialno";
                DBConnect cn = new DBConnect();
                reader = cn.queryforResult(sql, cn.getConnect());
                DataTable dt = new DataTable();
                dt.Load(reader);

                saveFileDialog1.InitialDirectory = @"D:\";
                saveFileDialog1.Filter = "Comma-Delimited Files (*.csv)|*.csv";
                saveFileDialog1.CheckPathExists = false;
                DialogResult result = saveFileDialog1.ShowDialog();

                if (result == DialogResult.OK)
                {
                    SaveToCSV(dt, saveFileDialog1.FileName);
                }
            }
            catch (Exception ex)
            {
                
            }
        }

        private void SaveToCSV(DataTable dt, string fileName)
        {
            try
            {
                string data = "";
                StreamWriter wr = new StreamWriter(fileName, false, System.Text.Encoding.Default);
                foreach (DataColumn column in dt.Columns)
                {
                    data += column.ColumnName + ",";
                }
                data += "\n";
                wr.Write(data);
                data = "";

                foreach (DataRow row in dt.Rows)
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (column.ColumnName.Equals("EMAIL_LIST"))
                        {
                            string mail = row[column].ToString().Trim().Replace(",", " | ");
                            data += mail + ",";
                        }
                        else
                        {
                            data += row[column].ToString().Trim() + ",";
                        }
                    }
                    data += "\n";
                    wr.Write(data);
                    data = "";
                }
                data += "\n";

                wr.Dispose();
                wr.Close();
                
                toolStripStatusLabel3.Text = "File save complete";
                MessageBox.Show("File save complet in " + fileName);
            }
            catch (Exception ex)
            {
                toolStripStatusLabel3.Text = ex.Message;
                MessageBox.Show(ex.Message);
            }            
        }

        private void SubmittedView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string Div = "", Oper = "", EqpModel = "", EqpNo = "", CreateTime = "", sStatus = "";
                //例如，注意項目的取法
                Div = this.SubmittedView.Rows[e.RowIndex].Cells["Division"].Value.ToString();
                Oper = this.SubmittedView.Rows[e.RowIndex].Cells["Operation"].Value.ToString();
                EqpModel = this.SubmittedView.Rows[e.RowIndex].Cells["Equip model"].Value.ToString();
                EqpNo = this.SubmittedView.Rows[e.RowIndex].Cells["Equip No"].Value.ToString();
                sRecipe = this.SubmittedView.Rows[e.RowIndex].Cells["Recipe Name"].Value.ToString();
                sPackage = this.SubmittedView.Rows[e.RowIndex].Cells["Package"].Value.ToString();
                CreateTime = this.SubmittedView.Rows[e.RowIndex].Cells["Created Time"].Value.ToString();
                sStatus = this.SubmittedView.Rows[e.RowIndex].Cells["STATUS"].Value.ToString();
                getLimitDefineData();
                var tmp = from n in lsFDCmater
                          where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && (n.Status.Equals(sStatus)||n.Status.Equals("Rejected")) && n.Package == sPackage && n.CreatedTime == CreateTime
                          select n;
                List<FDCMater> ls = new List<FDCMater>();
                foreach (var n in tmp)
                {
                    ls.Add(n);
                }

                ViewLimitGroup view = new ViewLimitGroup(ls, Div, Oper, EqpModel, EqpNo, sRecipe,sPackage, "Submitted", false, this);

                Add_TabPage("Approve limit group", view);
            }
            catch (Exception ex)
            {
                
            }
           
        }

        private void ApprovedView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string Div = "", Oper = "", EqpModel = "", EqpNo = "", CreateTime = "", sStatus = "";
                //例如，注意項目的取法
                Div = this.ApprovedView.Rows[e.RowIndex].Cells["Division"].Value.ToString();
                Oper = this.ApprovedView.Rows[e.RowIndex].Cells["Operation"].Value.ToString();
                EqpModel = this.ApprovedView.Rows[e.RowIndex].Cells["Equip model"].Value.ToString();
                EqpNo = this.ApprovedView.Rows[e.RowIndex].Cells["Equip No"].Value.ToString();
                sRecipe = this.ApprovedView.Rows[e.RowIndex].Cells["Recipe Name"].Value.ToString();
                sPackage = this.ApprovedView.Rows[e.RowIndex].Cells["Package"].Value.ToString();
                CreateTime = this.ApprovedView.Rows[e.RowIndex].Cells["Created Time"].Value.ToString();
                sStatus = this.ApprovedView.Rows[e.RowIndex].Cells["STATUS"].Value.ToString();
                getLimitDefineData();
                var tmp = from n in lsFDCmater
                          where n.Division.Equals(Div) && n.OperCode.Equals(Oper) && n.MachModel.Equals(EqpModel) && n.MachId.Equals(EqpNo) && n.RecipeName.Equals(sRecipe) && n.Status.Equals(sStatus) && n.Package == sPackage && n.CreatedTime == CreateTime
                          select n;
                List<FDCMater> ls = new List<FDCMater>();
                foreach (var n in tmp)
                {
                    ls.Add(n);
                }

                ViewLimitGroup view = new ViewLimitGroup(ls, Div, Oper, EqpModel, EqpNo, sRecipe,sPackage , "Approved", true, this);
                Add_TabPage("View limit group", view);
            }
            catch (Exception ex)
            {

            }
            
        }

        private void tabControl2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
               int x = e.X;
                int y = e.Y;

                Rectangle tabRect = this.tabControl2.GetTabRect(this.tabControl2.SelectedIndex);
                tabRect.Offset(tabRect.Width - 0x12, 2);
                tabRect.Width = 15;
                tabRect.Height = 15;
                if ((((x > tabRect.X) && (x < tabRect.Right)) && (y > tabRect.Y)) && (y < tabRect.Bottom))
                {
                    this.tabControl2.TabPages.Remove(this.tabControl2.SelectedTab);
                }
            }
        }

        public void getOperation()
        {
            try
            {
                cboOperation.Items.Clear();
                lsOperation.Clear();
                string tmp = "";
                try
                {
                    if (dicDivision[sDivision] != null)
                    {
                        tmp = dicDivision[sDivision];
                        string[] array = tmp.Split(':');
                        foreach (var item in array)
                        {
                            lsOperation.Add(item);
                            cboOperation.Items.Add(item);
                        }
                    }
                }
                catch (Exception)
                {
                    DBConnect cn = new DBConnect();
                    string sql = "select distinct rstype OPDESC from getcimrsc_ecim where RLLINE like '%" + sDivision + "%'";
                    reader = cn.queryforResult(sql, cn.getConnect());

                    while (reader.Read())
                    {
                        string str = reader["OPDESC"].ToString().Replace(' ', '_');
                        tmp += ":"+ str;
                        lsOperation.Add(str);
                        cboOperation.Items.Add(str);
                    }
                    cn.closeData();

                    dicDivision.Add(sDivision, tmp);
                }
                    

                //var operation = from n in lsCimRsc
                //                where n.LineCode.Equals(sDivision)
                //                select n;
              
                //foreach (var item in operation)
                //{
                //    if (!lsOperation.Contains(item.Operation))
                //    {
                //        lsOperation.Add(item.Operation);
                //        cboOperation.Items.Add(item.Operation);
                //    }
                //}              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Get Operation Error");
            }
        }

        private void cboOperation_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboEquipModel.Items.Clear();
            cboEquipModel.Text = "";

            cboEquipNO.Items.Clear();
            cboEquipNO.Text = "";

            sEqpModel = "";
            sEqpNo = "";

            sOperation = cboOperation.Text;

            setString();
            getEqpModel();
            if (!sDivision.Contains("DPS"))
            {//!=T3
                cboEquipNO.Enabled = true;
            }
            else
            {//T3
                if (!sOperation.Trim().Equals("Oven"))
                {
                    cboEquipNO.Enabled = false;
                }
            }
        }

        public void getEqpModel()
        {
            cboEquipModel.Items.Clear();
            lsEqpModel.Clear();
            string tmp = "";
            try
            {
                if (dicOperation[sOperation] != null)
                {
                    tmp = dicOperation[sOperation];
                    string[] array = tmp.Split(':');
                    foreach (var item in array)
                    {
                        lsEqpModel.Add(item);
                        cboEquipModel.Items.Add(item);
                    }
                }
            }
            catch (Exception)
            {
                DBConnect cn = new DBConnect();
                string sql = "select distinct rsstyp OPDESC from getcimrsc_ecim where rstype like '%" + sOperation + "%'";
                reader = cn.queryforResult(sql, cn.getConnect());

                while (reader.Read())
                {
                    string str = reader["OPDESC"].ToString();
                    tmp += ":" + str;
                    lsEqpModel.Add(str);
                    cboEquipModel.Items.Add(str);
                }
                cn.closeData();
                dicOperation.Add(sOperation, tmp);
            }

            //var rsc = from n in lsCimRsc
            //          where n.LineCode.Equals(sDivision) && n.Operation.Equals(sOperation)
            //          select n;

            //foreach (var item in rsc)
            //{
            //    if (!lsEqpModel.Contains(item.Model))
            //    {
            //        lsEqpModel.Add(item.Model);
            //        cboEquipModel.Items.Add(item.Model);
            //    }
            //}
        }

    }
}
