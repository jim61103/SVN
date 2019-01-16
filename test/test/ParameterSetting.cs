using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace iEMS_Setting
{
    public partial class ParameterSetting : Form
    {
        public string existMail = "";
        public int ParameterIndex;
        public int CollectIndex;
        public bool Parameter = false;
        public string model;
        public string eqpno;
        private string division;
        private string operation;
        public string sMailList;
        private bool ReadOnly;
        string str = "";
        OleDbDataReader reader;
        public List<Mail> lsMail = new List<Mail>();
        public List<Mail> lsSentMail = new List<Mail>();
        public List<Parameter> lsParameter = new List<Parameter>();
        public List<Parameter> lsCollectParameter = new List<Parameter>();

        public ParameterSetting()
        {
            InitializeComponent();
        }

        public ParameterSetting(string division, string operation, string sMailList, bool ReadOnly,string existMail)
        {
            this.division = division;
            this.operation = operation;
            this.sMailList = sMailList;
            this.ReadOnly = ReadOnly;
            this.existMail = existMail;
            if (existMail == null)
            {
                this.existMail = "";
            }
            //this.Parameter = !ReadOnly;

            InitializeComponent();
            btnAdd.Visible = !ReadOnly;
            btnRemove.Visible = !ReadOnly;
            btnParameterAll.Visible = !ReadOnly;
            btnCollectRemoveAll.Visible = !ReadOnly;

            MailListRun();

        }

        private void MailListRun()
        {
            label1.Text = "Select eMail address";
            label2.Text = "Available eMail address";
            label3.Text = "eMail address to be sent";
            getMailAddress();
            setParameterView();
            setCollectView();
        }

        private void setCollectView()
        {
            if (!Parameter)// for Mail Setting
            {
                DataTable Collect = new DataTable();
                Collect.Columns.Add("eMail Address", typeof(string));
                Collect.Columns.Add("Division", typeof(string));
                Collect.Columns.Add("Name", typeof(string));
                Collect.Columns.Add("Title", typeof(string));
                Collect.Columns.Add("Process", typeof(string));

                var sent = from n in lsSentMail
                           orderby n.Name
                           select n;

                foreach (var n in sent)
                {
                    Collect.Rows.Add(new object[] { n.MailAddr, n.Division, n.Name, n.Title, n.Process });
                }

                //var exist = from n in lsMail
                //            where existMail.IndexOf(n.MailAddr) > 0
                //            select n;

                //foreach (var n in exist)
                //{
                //    Collect.Rows.Add(new object[] { n.MailAddr, n.Division, n.Name, n.Title, n.Process });
                //}

                CollectView.DataSource = Collect;
                //DataGridViewColumn btn = new DataGridViewButtonColumn();
                //btn.Name = "Remove";
                //CollectView.Columns.Insert(0, btn);

                this.CollectView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                CollectView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                CollectView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //for (int i = 0; i < CollectView.ColumnCount - 1; i++)
                //{
                //    this.CollectView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //}
            }
            else //for Parameter Setting
            {
                DataTable Collect = new DataTable();
                Collect.Columns.Add("Parameter name", typeof(string));
                Collect.Columns.Add("Unit", typeof(string));
                Collect.Columns.Add("Parameter Type", typeof(string));
                Collect.Columns.Add("Data type", typeof(string));

                var collect = from n in lsCollectParameter
                              orderby n.ParamName
                              select n;

                foreach (var n in collect)
                {
                    string Type = "";
                    if (n.DataType.Equals("0"))
                    {
                        Type = "NUMBER";
                    }
                    else
                    {
                        Type = "STRING";
                    }
                    Collect.Rows.Add(new object[] { n.ParamName, n.ParamUnit, n.ParamType, Type });
                }

                CollectView.DataSource = Collect;
                //DataGridViewColumn btn = new DataGridViewButtonColumn();
                //btn.Name = "Remove";
                //CollectView.Columns.Insert(0, btn);

                this.CollectView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                CollectView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                CollectView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //for (int i = 0; i < CollectView.ColumnCount - 1; i++)
                //{
                //    this.CollectView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //}
            }
          

        }

        private void setParameterView()
        {
            if (!Parameter)//for Mail setting
            {
                DataTable parameter = new DataTable();
                parameter.Columns.Add("eMail Address", typeof(string));
                parameter.Columns.Add("Division", typeof(string));
                parameter.Columns.Add("Name", typeof(string));
                parameter.Columns.Add("Title", typeof(string));
                parameter.Columns.Add("Process", typeof(string));

                //existMail
                foreach (var n in lsMail)
                {
                    if (existMail.IndexOf(n.MailAddr) >=0)
                    {
                        lsSentMail.Add(n);
                    }
                }

                int count = lsMail.Count;
                for (int i = 0; i < lsSentMail.Count; i++)
                {
                    for (int j = 0; j < count; j++)
                    {
                        if (lsSentMail[i].MailAddr.Equals(lsMail[j].MailAddr))
                        {
                            lsMail.Remove(lsMail[j]);
                            j--;
                            count--;
                        }
                    }
                }

                //add to datatable
                var mail = from n in lsMail
                           where existMail.IndexOf(n.MailAddr)<0
                           orderby n.Name
                           select n;

                foreach (var n in mail)
                {                    
                    parameter.Rows.Add(new object[] { n.MailAddr, n.Division, n.Name, n.Title, n.Process });
                }
                ParameterView.DataSource = parameter;

                this.ParameterView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                ParameterView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                ParameterView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //for (int i = 0; i < ParameterView.ColumnCount - 1; i++)
                //{
                //    this.ParameterView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //}
            }
            else  //for Parameter setting
            {
                DataTable parameter = new DataTable();
                parameter.Columns.Add("Parameter name", typeof(string));
                parameter.Columns.Add("Unit", typeof(string));
                parameter.Columns.Add("Parameter Type", typeof(string));
                parameter.Columns.Add("Data type", typeof(string));

                var PName = from n in lsParameter
                            orderby n.ParamName
                            select n;

                foreach (var n in PName)
                {
                    string Type = "";
                    if (n.DataType.Equals("0"))
                    {
                        Type = "NUMBER";
                    }
                    else
                    {
                        Type = "STRING";
                    }
                    parameter.Rows.Add(new object[] { n.ParamName, n.ParamUnit, n.ParamType, Type });
                }

                ParameterView.DataSource = parameter;

                this.ParameterView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                ParameterView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                ParameterView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //for (int i = 0; i < ParameterView.ColumnCount - 1; i++)
                //{
                //    this.ParameterView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //}
            }
           
        }

        private void getMailAddress()
        {
            DBConnect con = new DBConnect();
            str = "select * from ATT_ECIM.eMails where Division='" + division + "' order by NAME";
            reader = con.queryforResult(str, con.getConnect());
            while (reader.Read())
            {
                Mail eMail = new Mail
                {
                    MailAddr = reader["MailAddr"].ToString(),
                    Division = reader["Division"].ToString(),
                    Name = reader["Name"].ToString(),
                    Title = reader["Title"].ToString(),
                    Process = reader["Process"].ToString(),
                    CreatedTime = reader["CreatedTime"].ToString(),
                    CreatedBy = reader["CreatedBy"].ToString(),
                };
                lsMail.Add(eMail);
            }
            con.closeData();
        }

        private void ParameterView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (Parameter)// for parameter setting use
                {
                    if (e.ColumnIndex == this.ParameterView.Columns["Add"].Index && !ReadOnly)
                    {
                        //MessageBox.Show(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                        //例如，注意項目的取法
                        string ParameterName = this.ParameterView.Rows[e.RowIndex].Cells["Parameter name"].Value.ToString();

                        if (setInsert(ParameterName))
                        {
                            ParameterView.Rows.RemoveAt(e.RowIndex);
                            List<Parameter> temp = lsParameter;
                            var tmp = from n in temp
                                      where n.ParamName == ParameterName
                                      select n;
                            foreach (var n in tmp)
                            {
                                lsParameter.Remove(n);
                                lsCollectParameter.Add(n);
                                break;
                            }

                            //object Collect = CollectView.CurrentCell;

                            setParameterView();
                            setCollectView();
                            try
                            {
                                ParameterView.CurrentCell = ParameterView.Rows[ParameterIndex].Cells[0];
                                CollectView.CurrentCell = CollectView.Rows[CollectIndex].Cells[0];
                            }
                            catch (Exception ex)
                            { }

                        }
                    }
                }
                else
                {
                    if (e.ColumnIndex == this.ParameterView.Columns["Add"].Index && !ReadOnly)
                    {
                        //例如，注意項目的取法
                        string eMailAddress = this.ParameterView.Rows[e.RowIndex].Cells["eMail Address"].Value.ToString();

                        ParameterView.Rows.RemoveAt(e.RowIndex);
                        List<Mail> temp = lsMail;
                        var tmp = from n in temp
                                    where n.MailAddr == eMailAddress
                                    select n;
                        foreach (var n in tmp)
                        {
                            lsMail.Remove(n);
                            lsSentMail.Add(n);
                            break;
                        }
                        
                        setParameterView();
                        setCollectView();
                        try
                        {
                            ParameterView.CurrentCell = ParameterView.Rows[ParameterIndex].Cells[0];
                            CollectView.CurrentCell = CollectView.Rows[CollectIndex].Cells[0];
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {}
            
        }
        private void setInsertTmp(string p,OleDbCommand cmd)
        {

            try
            {
                if (lsParameter.Count > 0)
                {
                    string sql = "";
                   
                    List<Parameter> lsP = lsParameter;
                    var tmp = from n in lsP
                              where n.ParamName == p
                              select n;
                    string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                    //Convert.ToDateTime(now)
                    //TO_DATE('" & a & "', 'YYYY/MM/DD HH24:MI:SS')
                    foreach (var n in tmp)
                    {
                        sql = "insert into ATT_ECIM.PARAMETERSPERDIV values ('" + division + "','" + operation + "','" + model + "','" + eqpno + "','" + n.ParamName + "','" + n.ParamType + "','" + n.DataType + "','" + n.ParamUnit + "',TO_DATE('" + now + "', 'YYYY/MM/DD HH24:MI:SS'),'" + n.CreatedBy + "')";
                        break;
                    }
                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }               
            }
            catch (Exception ex)
            {
               
            }
        }

        private bool setInsert(string p)
        {
            try
            {
                if (lsParameter.Count>0)
                {
                    string sql = "";
                    DBConnect con = new DBConnect();
                    List<Parameter> lsP = lsParameter;
                    var tmp = from n in lsP
                              where n.ParamName == p
                              select n;
                    string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                    //Convert.ToDateTime(now)
                    //TO_DATE('" & a & "', 'YYYY/MM/DD HH24:MI:SS')
                    foreach (var n in tmp)
                    {
                        sql = "insert into ATT_ECIM.PARAMETERSPERDIV values ('" + division + "','" + operation + "','" + model + "','" + eqpno + "','" + n.ParamName + "','" + n.ParamType + "','" + n.DataType + "','" + n.ParamUnit + "',TO_DATE('" + now + "', 'YYYY/MM/DD HH24:MI:SS'),'" + n.CreatedBy + "')";
                        break;
                    }

                    int i = con.insert(sql, con.getConnect());
                    con.closeData();
                    if (!i.Equals(0))
                    {
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Error in :" + sql);
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
        }

        private void setDeleteTemp(string p, OleDbCommand cmd)
        {
            try
            {
                if (lsCollectParameter.Count > 0)
                {
                    string sql = "";
                    DBConnect con = new DBConnect();
                    List<Parameter> lsP = lsCollectParameter;
                    var tmp = from n in lsP
                              where n.ParamName == p
                              select n;

                    foreach (var n in tmp)
                    {
                        sql = "delete from ATT_ECIM.PARAMETERSPERDIV where Division='" + division + "' and OperCode='" + operation + "' and MachModel='" + model + "' and ParamName='" + n.ParamName + "'";
                        break;
                    }

                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }               
            }
            catch (Exception ex)
            {
               
            }
        }

        private bool setDelete(string p)
        {
            try
            {
                if (lsCollectParameter.Count>0)
                {
                    string sql = "";
                    DBConnect con = new DBConnect();
                    List<Parameter> lsP = lsCollectParameter;
                    var tmp = from n in lsP
                              where n.ParamName == p
                              select n;

                    foreach (var n in tmp)
                    {
                        sql = "delete from ATT_ECIM.PARAMETERSPERDIV where Division='" + division + "' and OperCode='" + operation + "' and MachModel='" + model + "' and ParamName='" + n.ParamName + "'";
                        break;
                    }

                    int i = con.insert(sql, con.getConnect()); //Excute cmd
                    con.closeData();
                    if (!i.Equals(0))
                    {
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Error in :" + sql);
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
        }

        private void CollectView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (Parameter) // for parameter setting
                {
                    if (e.ColumnIndex == this.CollectView.Columns["Remove"].Index && !ReadOnly)
                    {
                        string ParameterName = this.CollectView.Rows[e.RowIndex].Cells["Parameter name"].Value.ToString();

                        if (setDelete(ParameterName))
                        {
                            CollectView.Rows.RemoveAt(e.RowIndex);
                            List<Parameter> send = lsCollectParameter;
                            var sent = from n in send
                                       where n.ParamName == ParameterName
                                       select n;
                            foreach (var s in sent)
                            {
                                // ParameterView.Rows.Add(new object[] { n.MailAddr, n.Division, n.Name, n.Title, n.Process });
                                lsCollectParameter.Remove(s);
                                lsParameter.Add(s);
                                break;
                            }
                            setParameterView();
                            setCollectView();
                            try
                            {
                                ParameterView.CurrentCell = ParameterView.Rows[ParameterIndex].Cells[0];
                                CollectView.CurrentCell = CollectView.Rows[CollectIndex].Cells[0];
                            }
                            catch (Exception ex)
                            { }
                        }
                    }
                }
                else//for email setting
                {
                    if (e.ColumnIndex == this.CollectView.Columns["Remove"].Index && !ReadOnly)
                    {
                        string eMailAddress = this.CollectView.Rows[e.RowIndex].Cells["eMail Address"].Value.ToString();

                        
                        List<Mail> send = lsSentMail;
                        var sent = from n in send
                                    where n.MailAddr == eMailAddress
                                    select n;
                        foreach (var s in sent)
                        {
                            // ParameterView.Rows.Add(new object[] { n.MailAddr, n.Division, n.Name, n.Title, n.Process });
                            lsSentMail.Remove(s);
                            lsMail.Add(s);
                            break;
                        }
                        CollectView.Rows.RemoveAt(e.RowIndex);
                        existMail = existMail.Replace(eMailAddress, "");
                        setParameterView();
                        setCollectView();
                        try
                        {
                            ParameterView.CurrentCell = ParameterView.Rows[ParameterIndex].Cells[0];
                            CollectView.CurrentCell = CollectView.Rows[CollectIndex].Cells[0];
                        }
                        catch (Exception ex)
                        {

                        }
                        
                    }
                }

            }
            catch (Exception ex)
            {}           
        }

        private void btnParameterAll_Click(object sender, EventArgs e)
        {
            try
            {
                if (Parameter)
                {
                    if (!ReadOnly)
                    {
                        int count = lsParameter.Count();
                        OleDbConnection con = new OleDbConnection(Global.ConnectDB);
                        con.Open();
                        OleDbCommand cmd = con.CreateCommand();
                        cmd.Connection = con;
                        OleDbTransaction trans = con.BeginTransaction();
                        cmd.Transaction = trans;
                        
                        for (int i = 0; i < count; i++)
                        {
                            setInsertTmp(lsParameter[0].ParamName,cmd);
                           
                            lsCollectParameter.Add(lsParameter[0]);
                            lsParameter.RemoveAt(0);
                          
                        }

                        trans.Commit();
                        con.Close();
                        con.Dispose();
                        setParameterView();
                        setCollectView();
                    }
                }
                else
                {
                    if (!ReadOnly)
                    {
                        int count = lsMail.Count();
                        for (int i = 0; i < count; i++)
                        {
                            lsSentMail.Add(lsMail[0]);
                            lsMail.RemoveAt(0);
                        }
                        setParameterView();
                        setCollectView();
                    }
                }
              
            }
            catch (Exception ex)
            {}
        
        }

        private void btnCollectRemoveAll_Click(object sender, EventArgs e)
        {
            try
            {
                if (Parameter)
                {
                    if (!ReadOnly)
                    {
                        int count = lsCollectParameter.Count();

                        OleDbConnection con = new OleDbConnection(Global.ConnectDB);
                        con.Open();
                        OleDbCommand cmd = con.CreateCommand();
                        cmd.Connection = con;
                        OleDbTransaction trans = con.BeginTransaction();
                        cmd.Transaction = trans;

                        for (int i = 0; i < count; i++)
                        {
                            setDeleteTemp(lsCollectParameter[0].ParamName, cmd);
                            
                                lsParameter.Add(lsCollectParameter[0]);
                                lsCollectParameter.RemoveAt(0);
                                                        
                        }

                        trans.Commit();
                        con.Close();
                        con.Dispose();
                        setParameterView();
                        setCollectView();
                    }
                }
                else
                {
                    if (!ReadOnly)
                    {
                        int count = lsSentMail.Count();
                        for (int i = 0; i < count; i++)
                        {
                            lsMail.Add(lsSentMail[0]);
                            lsSentMail.RemoveAt(0);
                        }
                        existMail = "";
                        setParameterView();
                        setCollectView();
                    }
                }
               
               
            }
            catch (Exception ex)
            {}
         
        }

        private void ParameterSetting_Load(object sender, EventArgs e)
        {

        }
        public ParameterSetting(string division,string operation,string model,string eqpno,List<Parameter>lsParameter,List<Parameter>lsCollect)
        {
            ReadOnly = false;
            Parameter = true;
            this.division = division;
            this.operation = operation;
            this.eqpno = eqpno;
            this.model = model;
            this.lsParameter = lsParameter;
            this.lsCollectParameter = lsCollect;
            InitializeComponent();

            getParameterRun();

        }

        private void getParameterRun()
        {
            //getAvailableParms();
            //getControlPararms();
            setLabel();
            setGridView();
        }

        private void getControlPararms()
        {
            try
            {
                DBConnect con = new DBConnect();
                //select * from ATT_ECIM.PARAMETERSPERDIV where Division=? and OperCode=? and MachModel=?
                str = "select * from ATT_ECIM.PARAMETERSPERDIV where Division='" + division + "' and OperCode='" + operation + "' and MachModel='" + model + "'";

                if (!eqpno.Equals(""))
                {
                    str += " and MachId ='" + eqpno + "'";
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
                con.closeData();                
            }
            catch (Exception ex)
            {

            }
        }

        private void setGridView()
        {
            setList();
            setParameterView();
            setCollectView();
        }

        private void setList()
        {
            try
            {
                List<Parameter> collect = lsCollectParameter;
                List<Parameter> parameter = lsParameter;
                List<Parameter> temp = new List<Parameter>();
                int count = collect.Count();
                
                if (parameter.Count > 0 && collect.Count > 0)
                {
                    foreach (var p in parameter)
                    {
                        bool find = false;
                        foreach (var c in collect)
                        {
                            if (p.ParamName == c.ParamName)
                            {
                                find = true;
                            }
                        }

                        if (!find)
                        {
                            temp.Add(p);
                        }
                    }
                    lsParameter = temp;
                }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
        }

        private void setLabel()
        {
            //Division : T1 Bump / Operation code : T1_BP_PLATER / Equip model : Raider 312 / Equip No : T1_BP_BPLT001
            label1.Text = "Division : "+division+" / Operation code : "+operation+" / Equip model : "+model+" / Equip No : "+eqpno;
            label2.Text = "Available parameter";
            label3.Text = "Controllable parameter";
        }

        private void getAvailableParms()
        {
            try
            {
                DBConnect con = new DBConnect();
                //select * from ATT_ECIM.Parameters where MachModel=? and MachId =? ORDER BY PARAMNAME 
                str = "select * from ATT_ECIM.Parameters where MachModel='" + model + "'";

                if (!eqpno.Equals(""))
                {
                    str += " and MachId ='" + eqpno + "'";
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
                DataType =Convert.ToByte(reader["DataType"]),
                ParamUnit = reader["ParamUnit"].ToString(),
                CreatedTime = reader["CreatedTime"].ToString(),
                CreatedBy =reader["CreatedBy"].ToString()
            };
            lsParameter.Add(P);
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
            lsCollectParameter.Add(P);
        }

        private void ParameterSetting_FormClosed(object sender, FormClosedEventArgs e)
        {

            if (!ReadOnly)
            {
                for (int i = 0; i < lsSentMail.Count; i++)
                {
                    sMailList += lsSentMail[i].MailAddr;
                    if (i < lsSentMail.Count - 1)
                    {
                        sMailList += ",";
                    }
                }
                lsSentMail.Clear();
            }
        }

        private void ParameterView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ParameterIndex = e.RowIndex;
        }

        private void CollectView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            CollectIndex = e.RowIndex;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                if (Parameter) //for parameter setting
                {
                    if (!ReadOnly)
                    {
                        int count = ParameterView.Rows.Count;
                        int iParam = ParameterView.Rows.Count;
                        int iCollect = CollectView.Rows.Count;
                        for (int i = count; i > 0; i--)
                        {
                            DataGridViewRow row = ParameterView.Rows[i - 1];
                            if (row.Cells[1].Selected)
                            {
                                string ParameterName = this.ParameterView.Rows[i - 1].Cells["Parameter name"].Value.ToString();
                                iParam--;
                                if (setInsert(ParameterName))
                                {
                                    ParameterView.Rows.RemoveAt(row.Index);
                                    List<Parameter> temp = lsParameter;
                                    var tmp = from n in temp
                                              where n.ParamName == ParameterName
                                              select n;
                                    foreach (var n in tmp)
                                    {
                                        lsParameter.Remove(n);
                                        lsCollectParameter.Add(n);
                                        break;
                                    }
                                }
                            }
                        }

                        setParameterView();
                        setCollectView();
                        try
                        {
                            //ParameterView.CurrentCell = ParameterView.Rows[iParam-1].Cells[1];
                            //CollectView.CurrentCell = CollectView.Rows[iCollect].Cells[1];
                            ParameterView.CurrentCell = ParameterView.Rows[ParameterIndex].Cells[1];
                            CollectView.CurrentCell = CollectView.Rows[CollectIndex].Cells[1];
                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show(ex.ToString());
                        }
                    }                   
                }
                else //for email setting
                {
                    if (!ReadOnly)
                    {
                        int count = ParameterView.Rows.Count;
                        for (int i = count; i > 0; i--)
                        {
                            DataGridViewRow row = ParameterView.Rows[i - 1];
                            if (row.Cells[1].Selected)
                            {
                                string eMailAddr = this.ParameterView.Rows[i - 1].Cells["eMail Address"].Value.ToString();
    
                                ParameterView.Rows.RemoveAt(row.Index);
                                List<Mail> temp = lsMail;
                                var tmp = from n in temp
                                            where n.MailAddr == eMailAddr
                                            select n;
                                foreach (var n in tmp)
                                {
                                    lsMail.Remove(n);
                                    lsSentMail.Add(n);
                                    break;
                                }                                
                            }
                        }

                        setParameterView();
                        setCollectView();
                        try
                        {
                            ParameterView.CurrentCell = ParameterView.Rows[ParameterIndex].Cells[0];
                            CollectView.CurrentCell = CollectView.Rows[CollectIndex].Cells[0];
                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show(ex.ToString());
                        }
                    }
                }         
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
           
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            try
            {
                if (Parameter) //for parameter setting
                {
                    if (!ReadOnly)
                    {
                        int count = CollectView.Rows.Count;
                        for (int i = count; i > 0; i--)
                        {
                            DataGridViewRow row = CollectView.Rows[i - 1];
                            if (row.Cells[1].Selected)
                            {
                                string ParameterName = this.CollectView.Rows[i - 1].Cells["Parameter name"].Value.ToString();

                                if (setDelete(ParameterName))
                                {
                                    CollectView.Rows.RemoveAt(row.Index);
                                    List<Parameter> send = lsCollectParameter;
                                    var sent = from n in send
                                               where n.ParamName == ParameterName
                                               select n;
                                    foreach (var s in sent)
                                    {
                                        // ParameterView.Rows.Add(new object[] { n.MailAddr, n.Division, n.Name, n.Title, n.Process });
                                        lsCollectParameter.Remove(s);
                                        lsParameter.Add(s);
                                        break;
                                    }
                                }
                            }
                        }
                        setParameterView();
                        setCollectView();
                        try
                        {
                            ParameterView.CurrentCell = ParameterView.Rows[ParameterIndex].Cells[0];
                            CollectView.CurrentCell = CollectView.Rows[CollectIndex].Cells[0];
                        }
                        catch (Exception ex)
                        {
                           // MessageBox.Show(ex.ToString());
                        }
                    }
                }
                else// for email setting
                {
                    if (!ReadOnly)
                    {
                        int count = CollectView.Rows.Count;
                        for (int i = count; i > 0; i--)
                        {
                            DataGridViewRow row = CollectView.Rows[i - 1];
                            if (row.Cells[1].Selected)
                            {
                                string eMailAddr = this.CollectView.Rows[i - 1].Cells["eMail Address"].Value.ToString();

                                CollectView.Rows.RemoveAt(row.Index);
                                List<Mail> temp = lsSentMail;
                                var tmp = from n in temp
                                          where n.MailAddr == eMailAddr
                                          select n;
                                foreach (var n in tmp)
                                {
                                    lsSentMail.Remove(n);
                                    lsMail.Add(n);
                                    break;
                                }
                            }
                        }

                        setParameterView();
                        setCollectView();
                        try
                        {
                            ParameterView.CurrentCell = ParameterView.Rows[ParameterIndex].Cells[0];
                            CollectView.CurrentCell = CollectView.Rows[CollectIndex].Cells[0];
                        }
                        catch (Exception ex)
                        {
                           // MessageBox.Show(ex.ToString());
                        }
                    }
                }
                
       
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
           
        }
    }
}
