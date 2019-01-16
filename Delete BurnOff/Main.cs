using iEMS_Setting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Delete_BurnOff
{
    public partial class Main : Form
    {
        DataSet ds = new DataSet();
        OleDbDataAdapter da = new OleDbDataAdapter();
        OleDbDataReader reader ;
        string RSC = "", sql = "";
        public Main()
        {
            InitializeComponent();            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Global.ConnectDB = Global.BUMPConnectDB; //Global.BUMP_UAT_ConnectDB; //Global.BUMPConnectDB;

            GetData();
        }

        private void btn_Query_Click(object sender, EventArgs e)
        {
            
            GetData();
        }

        public void GetData()
        {
            ds.Clear();
            DBConnect cn = new DBConnect();
            sql = "select t.resourcename as RSC,t.trackindatetime from amkaas_lot_validate t where BUMPRECIPECODE = 'BUR1_P9'";
            da = cn.queryforAdapter(sql, cn.getConnect());
            da.Fill(ds,"view");
            dataGridView1.DataSource = ds.Tables["view"];
            cn.closeData();
            Count.Text = ds.Tables[0].Rows.Count.ToString();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void btn_Delete_Click(object sender, EventArgs e)
        {
            if (RSC != "")
            {
                DialogResult result = MessageBox.Show("Delete " + RSC + " BurnOff Data?", "User confirmation"
                                              , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    DBConnect cn = new DBConnect();
                    try
                    {
                        btn_Delete.Enabled = false;
                        sql = "Delete from amkaas_lot_validate t where resourcename = '" + RSC + "' and BUMPRECIPECODE = 'BUR1_P9'";
                        reader = cn.queryforResult(sql, cn.getConnect());
                        cn.closeData();
                        MessageBox.Show("Delete " + RSC + " success!!");
                    }
                    catch (Exception ex)
                    {
                        cn.closeData();
                        btn_Delete.Enabled = true;
                        RSC = "";
                    }

                }
                RSC = "";
                btn_Delete.Enabled = true;
                GetData();
            }
            else
            {
                MessageBox.Show("Please select RSC first");
            }
            
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                RSC = dataGridView1.Rows[e.RowIndex].Cells["RSC"].Value.ToString();
            }
            catch (Exception ex)
            {
                
            }
            
        }
    }
}
