using System;
using System.Windows.Forms;
using System.DirectoryServices.Protocols;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;

namespace iEMS_Setting
{
    public partial class LogInGUI : Form
    {
        public bool Success = false;
        public string LineCode = "";
        public string ErrorCode = "";
        public string sUserID = "";

        public LogInGUI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                sUserID = txt_ID.Text;
                DirectoryEntry ldap = new DirectoryEntry(string.Format("LDAP://10.185.32.103"), "tw\\"+ sUserID, txt_PWD.Text);
                object nativeObject = ldap.NativeObject;
                Success = true;
                this.Dispose();
            }
            catch (Exception ex)
            {
                ErrorCode = ex.Message.ToString();
                txt_PWD.Text = "";
                Success = false;
                MessageBox.Show(ErrorCode);
            }
                //LdapConnection conn = new LdapConnection("LDAP://10.185.32.103");
        }

        private void LogInGUI_Load(object sender, EventArgs e)
        {
            if (Global.isUAT)
            {                
                label_Env.Text += " UAT";
            }
            else
            {
                label_Env.Text += " Production";
            }

            try
            {
                label_Version.Text += System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            catch (Exception ex)
            {
            }
            
            txt_ID.Select();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            sUserID = "";
            Success = false;
            this.Close();
        }

        private void txt_ID_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if ((int)e.KeyChar < 48 | (int)e.KeyChar > 57)
            //{
            //    e.Handled = true;
            //}
        }

        private void txt_PWD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_OK.Focus();
                button1_Click(sender, e);
                //TextBox1.Focus();
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label_Env_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }
    }
}
