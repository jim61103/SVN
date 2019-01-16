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
    public partial class eMailView : Form
    {
        OleDbDataReader reader;

        public List<Mail> lsMail = new List<Mail>();
        public List<string> lsDivision = new List<string>();
        public bool edit = false;
        public List<string> lsProcess = new List<string>();
        public string status = "";

        public eMailView()
        {
            InitializeComponent();
        }

        public eMailView(bool edit, List<Mail> lsMail, List<string> lsDiv,string status,Form F)
        {
            this.edit = edit;
            this.lsMail = lsMail;
            this.lsDivision = lsDiv;
            this.status = status;
            this.Tag = F;
            InitializeComponent();
            Icon myIcon = null;
            switch (status)
            {
                case "Add":
                    this.Text = status + " eMail List";
                    myIcon = Properties.Resources.add;
                    
                    break;
                case "Edit":
                    this.Text = status + " eMail List";
                    myIcon = Properties.Resources.edit1;                 
                    break;
                case "Delete":
                    this.Text = status + " eMail List";
                    myIcon = Properties.Resources.delete1;
                    break;
                default:
                    break;
            }
            this.Icon = myIcon;
        }

        private void eMailView_Load(object sender, EventArgs e)
        {
            fillProcess();
            fillGridView();
        }

        private void fillProcess()
        {
            lsProcess.Clear();
            DBConnect cn = new DBConnect();

            try
            {
                string sql = "select distinct OPCODE, OPDESC from GETOPERATION order by OPCODE asc";
                reader = cn.queryforResult(sql, cn.getConnect());
                while (reader.Read())
                {
                    string str = "";
                    string OPCODE = reader["OPCODE"].ToString();
                    string OPDESC = reader["OPDESC"].ToString();

                    str = OPCODE + " " + OPDESC;
                    lsProcess.Add(str);
                }
            }
            catch (Exception ex)
            {}
            finally
            {
                cn.closeData();
            }           
        }

        private void fillGridView()
        {
            DataTable View = new DataTable("View");
            View.Columns.Add("Item", typeof(string));
            View.Columns.Add("Value", typeof(string));

            

            if (lsMail != null)//edit and delete
            {
                View.Rows.Add(new object[] { "eMail Address", lsMail[0].MailAddr });
                View.Rows.Add(new object[] { "Dvivsion", lsMail[0].Division });
                View.Rows.Add(new object[] { "Name", lsMail[0].Name });
                View.Rows.Add(new object[] { "Title", lsMail[0].Title });
                View.Rows.Add(new object[] { "Process", lsMail[0].Process });
                View.Rows.Add(new object[] { "Created Time", lsMail[0].CreatedTime });
                View.Rows.Add(new object[] { "Created By", lsMail[0].CreatedBy });
                ViewMail.DataSource = View;

                DataGridViewComboBoxCell tCell = new DataGridViewComboBoxCell();
                tCell.DataSource = lsDivision;
                ViewMail["Value", 1] = tCell;

                DataGridViewComboBoxCell Process = new DataGridViewComboBoxCell();
                lsProcess.Add("");
                Process.DataSource = lsProcess;
                ViewMail["Value", 4] = Process;

                ViewMail.MultiSelect = false;
                ViewMail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                if (status == "Delete")
                {
                    ViewMail.ReadOnly = true;
                }
            }
            else// btnAdd
            {
                View.Rows.Add(new object[] { "eMail Address", "" });
                View.Rows.Add(new object[] { "Dvivsion",lsDivision[0] });
                View.Rows.Add(new object[] { "Name", "" });
                View.Rows.Add(new object[] { "Title", "" });
                View.Rows.Add(new object[] { "Process" });

                ViewMail.DataSource = View;

                DataGridViewComboBoxCell tCell = new DataGridViewComboBoxCell();
                lsDivision.Add("");
                tCell.DataSource = lsDivision;
                ViewMail["Value", 1] = tCell;

                DataGridViewComboBoxCell Process = new DataGridViewComboBoxCell();
                lsProcess.Add("");
                Process.DataSource = lsProcess;
                ViewMail["Value", 4] = Process;


                ViewMail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            switch (status)
            {
                case "Add":
                    eventAddOkBtnSelected();
                    break;
                case "Edit":
                    eventEditOkBtnSelected();
                    break;
                case "Delete":
                    eventDeleteOkBtnSelected();
                    break;
                default:
                    break;
            }
        }

        private void eventAddOkBtnSelected()
        {
            try
            {
                Mail vo = new Mail();
                vo.MailAddr = this.ViewMail.Rows[0].Cells["Value"].Value.ToString();
                vo.Division = this.ViewMail.Rows[1].Cells["Value"].Value.ToString();
                vo.Name = this.ViewMail.Rows[2].Cells["Value"].Value.ToString();
                vo.Title = this.ViewMail.Rows[3].Cells["Value"].Value.ToString();
                vo.Process = this.ViewMail.Rows[4].Cells["Value"].Value.ToString();
                string now = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                vo.CreatedTime = now;
                vo.CreatedBy = Global.UserID;

                if (vo.MailAddr.Trim().Equals("") || vo.Division.Trim().Equals(""))
                {
                    MessageBox.Show("It must be inputted eMail Address and Division!", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (isThereSameRecord(vo))
                    {
                        MessageBox.Show("Exist same record at server. Can't register it.", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        if (insertEmailSettingList(vo) == -1)
                        {
                            MessageBox.Show("Failed to add eMail list", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Succeed to add eMail list", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ((eMailSetting)this.Tag).find();
                            this.Close();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                
            }
        }

        private int insertEmailSettingList(Mail vo)
        {
            int iReturn = -1;
            DBConnect cn = new DBConnect();
            try
            {
                if (vo.CreatedTime != null) vo.CreatedTime = vo.CreatedTime.IndexOf("午") > -1 ? Convert.ToDateTime(vo.CreatedTime).ToString("yyyy/MM/dd HH:mm:ss") : vo.CreatedTime;
                string sql = "insert into ATT_ECIM.eMails values ('"+vo.MailAddr+"','"+vo.Division+"','"+vo.Name+"','"+vo.Title+"','"+vo.Process+ "', TO_DATE('" + vo.CreatedTime + "', 'YYYY/MM/DD HH24:MI:SS'),'" + vo.CreatedBy+"')";
                iReturn = cn.insert(sql, cn.getConnect());
                return iReturn;

            }
            catch (Exception ex)
            {
                return -1;
            }
            finally
            {
                cn.closeData();
            }
        }

        private bool isThereSameRecord(Mail vo)
        {
            DBConnect cn = new DBConnect();
            try
            {
                string sql = "select * from ATT_ECIM.eMails where MailAddr='"+vo.MailAddr+"' and Division='"+vo.Division+"'";
                reader = cn.queryforResult(sql, cn.getConnect());
                if (reader.HasRows)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return true;
            }
            finally { cn.closeData(); }
        }

        private void eventEditOkBtnSelected()
        {
            throw new NotImplementedException();
        }

        private void eventDeleteOkBtnSelected()
        {
            Mail vo = new Mail();
            vo.MailAddr = this.ViewMail.Rows[0].Cells["Value"].Value.ToString();
            vo.Division = this.ViewMail.Rows[1].Cells["Value"].Value.ToString();
            vo.Name = this.ViewMail.Rows[2].Cells["Value"].Value.ToString();
            vo.Title = this.ViewMail.Rows[3].Cells["Value"].Value.ToString();
            vo.Process = this.ViewMail.Rows[4].Cells["Value"].Value.ToString();

            DialogResult result = MessageBox.Show("Confirm deleting?", "User confirmation"
                                              , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    if (deleteEmailSettingList(vo) == 1)
                    {
                        MessageBox.Show("Succeed to delete eMail list", "Information for user", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ((eMailSetting)this.Tag).find();
                        this.Close();
                    }
                }
                catch (Exception ex)
                {
                    
                }
            }
        }

        private int deleteEmailSettingList(Mail vo)
        {
            int iReturn = -1;
            DBConnect cn = new DBConnect();
            try
            {
                string sql = "delete ATT_ECIM.eMails where MailAddr='" + vo.MailAddr + "' and Division='" + vo.Division + "'";
                iReturn = cn.insert(sql, cn.getConnect());
                return iReturn;
            }
            catch (Exception)
            {
                return -1;
            }
            finally { cn.closeData(); }
        }
    }
}
