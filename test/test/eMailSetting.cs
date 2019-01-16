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
    public partial class eMailSetting : Form
    {
        OleDbDataReader reader;
        public List<string> lsDiv = new List<string>();
        public List<Mail> lsMail = new List<Mail>();
        public int Index = 0;
        public string sDiv = "";


        public eMailSetting()
        {
            
        }

        public eMailSetting(List<string> lsDivision)
        {
            this.lsDiv = lsDivision;

            InitializeComponent();

        }

        private void eMailSetting_Load(object sender, EventArgs e)
        {
            fillDivision();
        }

        private void fillDivision()
        {
            cboDiv.Items.Add("");
            foreach (var item in lsDiv)
            {
                cboDiv.Items.Add(item);
            }
        }

        public void fillGridView()
        {
            try
            {
                DataTable View = new DataTable("View");
                View.Columns.Add("eMail Address", typeof(string));
                View.Columns.Add("Division", typeof(string));
                View.Columns.Add("Name", typeof(string));
                View.Columns.Add("Title", typeof(string));
                View.Columns.Add("Process", typeof(string));
                View.Columns.Add("Created Time", typeof(string));
                View.Columns.Add("Created By", typeof(string));

                foreach (var n in lsMail)
                {
                    View.Rows.Add(new object[] {n.MailAddr,n.Division,n.Name,n.Title,n.Process,n.CreatedTime,n.CreatedBy });
                    //View.Rows.Add(new object[] { n.ParameterName, true, secondflag, n.StepName, n.StepValue, n.EventId, n.WarmUpTime, n.StartDetectValue, "", "", checkmin, AllowMinEqual, n.MinValue, checkmax, AllowMaxEqual, n.MaxValue, RChartFlag, CheckRChartMin, n.RMinValue, CheckRChartMax, n.RMaxValue, TerminalMsg, Email, McInhibit, McHold, Eocap, LotHold });
                }
                VieweMail.MultiSelect = false;
                VieweMail.DataSource = View;
                VieweMail.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                VieweMail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            }
            catch (Exception ex)
            {
                
            }
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            find();
        }

        public void find()
        {
            sDiv = cboDiv.Text;
            if (sDiv.Equals(""))
            {
                getAllEmailSettingList();
            }
            else
            {
                getEmailSettingList(sDiv);
            }
        }

        private void getEmailSettingList(string div)
        {
            lsMail.Clear();

            DBConnect cn = new DBConnect();

            try
            {
                string sql = "select * from ATT_ECIM.eMails where Division='" + div + "'";
                reader = cn.queryforResult(sql, cn.getConnect());
                while (reader.Read())
                {
                    //FDCMater FDCmaster = new FDCMater
                    Mail mail = new Mail
                    {
                        MailAddr = reader["MailAddr"].ToString().Trim(),
                        Division = reader["Division"].ToString().Trim(),
                        Name = reader["Name"].ToString().Trim(),
                        Title = reader["Title"].ToString().Trim(),
                        Process = reader["Process"].ToString().Trim(),
                        CreatedTime = reader["CreatedTime"].ToString().Trim(),
                        CreatedBy = reader["CreatedBy"].ToString().Trim()
                    };
                    lsMail.Add(mail);
                }
            }
            catch (Exception ex)
            {
                
            }
           finally { cn.closeData(); }

            fillGridView();
        }

        private void getAllEmailSettingList()
        {
            lsMail.Clear();
            DBConnect cn = new DBConnect();
            try
            {
                
                string sql = "select * from ATT_ECIM.eMails";
                reader = cn.queryforResult(sql, cn.getConnect());
                while (reader.Read())
                {
                    //FDCMater FDCmaster = new FDCMater
                    Mail mail = new Mail
                    {
                        MailAddr = reader["MailAddr"].ToString().Trim(),
                        Division = reader["Division"].ToString().Trim(),
                        Name = reader["Name"].ToString().Trim(),
                        Title = reader["Title"].ToString().Trim(),
                        Process = reader["Process"].ToString().Trim(),
                        CreatedTime = reader["CreatedTime"].ToString().Trim(),
                        CreatedBy = reader["CreatedBy"].ToString().Trim()
                    };
                    lsMail.Add(mail);
                }
            }
            catch (Exception ex)
            {
               
            }
            finally
            {
                cn.closeData();
            }
            
            fillGridView();

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            eMailView view = new eMailView(true,null,lsDiv,"Add",this); // add 不用給lsMail
            view.Show();
        }

        private void VieweMail_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Index = e.RowIndex;
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (Index>=0)
            {
                //string Parameter = this.ViewLimit.Rows[e.RowIndex].Cells["Parameter name"].Value.ToString();
                string Address = this.VieweMail.Rows[Index].Cells["eMail Address"].Value.ToString();

                var tmp = from n in lsMail
                          where n.MailAddr == Address
                          select n;
                List<Mail> ls = new List<Mail>();
                foreach (var item in tmp)
                {
                    ls.Add(item);
                }
                eMailView view = new eMailView(true, ls, lsDiv, "Edit",this);
                view.Show();
            }
            else
            {
                MessageBox.Show("Select eMail list that will be edited", "ErrorMessage", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
          
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (Index>=0)
            {
                string Address = this.VieweMail.Rows[Index].Cells["eMail Address"].Value.ToString();

                var tmp = from n in lsMail
                          where n.MailAddr == Address
                          select n;
                List<Mail> ls = new List<Mail>();
                foreach (var item in tmp)
                {
                    ls.Add(item);
                }
                eMailView view = new eMailView(true, ls, lsDiv, "Delete",this);
                view.Show();
            }
            else
            {
                MessageBox.Show("Select eMail list that will be delete", "ErrorMessage", MessageBoxButtons.OK, MessageBoxIcon.Information);            
            }
           
        }
    }
}
