using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace iEMS_Setting
{
    public partial class Compare_Parameter : Form
    {
        List<FDCMater> lsMaster = new List<FDCMater>();
        List<ParameterLimitGroup> lsParameter = new List<ParameterLimitGroup>();
        List<FDCMater> lsMasterHis = new List<FDCMater>();
        List<ParameterLimitGroup> lsParameterHis = new List<ParameterLimitGroup>();
        public Dictionary<string, string> MailDic = new Dictionary<string, string>();

        public Compare_Parameter()
        {
            InitializeComponent();
        }

        public Compare_Parameter(List<FDCMater> lsFDCMaster,List<FDCMater> lsFDCMasterHistory,List<ParameterLimitGroup> lsParameter,List<ParameterLimitGroup> lsParameterHistory)
        {
            InitializeComponent();
            this.lsMaster = lsFDCMaster;
            this.lsMasterHis = lsFDCMasterHistory;
            this.lsParameter = lsParameter;
            this.lsParameterHis = lsParameterHistory;
 
            showData();
        }

        private void showData()
        {
            try
            {
                //label2.Text += lsParameter[0].Revision ;
                //label3.Text += lsParameterHis[0].Revision;
                FillGridView(CurrentView, lsMaster, lsParameter);
                FillGridView(PreviousView, lsMasterHis, lsParameterHis);
                               

            }
            catch (Exception ex)
            {
                throw;
            }           
        }

        private void FillGridView(DataGridView DGV,List<FDCMater> lsMaster,List<ParameterLimitGroup>lsParameter)
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
                View.Columns.Add("eMail List", typeof(string));
                View.Columns.Add("M/C inhibit", System.Type.GetType("System.Boolean"));
                View.Columns.Add("M/C hold", System.Type.GetType("System.Boolean"));
                View.Columns.Add("eOCAP", System.Type.GetType("System.Boolean"));
                View.Columns.Add("Lot hold", System.Type.GetType("System.Boolean"));

                foreach (var n in lsParameter)
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
                    if (n.eMailAddrs == null)
                    {
                        //MailDic.Add(n.ParameterName, "");
                        n.eMailAddrs = "";
                    }
                    //MailDic.Add(n.ParameterName, n.eMailAddrs);
                    View.Rows.Add(new object[] { n.ParameterName, true, secondflag, n.StepName, n.StepValue, n.EventId, n.WarmUpTime, n.MonitorTime, n.StartDetectValue, "", "", checkmin, AllowMinEqual, n.MinValue, checkmax, AllowMaxEqual, n.MaxValue, RChartFlag, CheckRChartMin, n.RMinValue, CheckRChartMax, n.RMaxValue, TerminalMsg, Email, n.eMailAddrs, McInhibit, McHold, Eocap, LotHold });
                }

                //DGV.MultiSelect = false;
                DGV.DataSource = View;
                //DataGridViewColumn btn = new DataGridViewButtonColumn();
                //btn.Name = "eMail List";
                //DGV.Columns.Insert(23, btn);

                DataGridViewComboBoxColumn cbo = new DataGridViewComboBoxColumn();
                cbo.Name = "eOCAP Templete";
                cbo.HeaderText = "eOCAP Templete";


                List<string> lseOCAPtmp = new List<string>();
                for (int i = 0; i < lsParameter.Count; i++)
                {
                    if (!lseOCAPtmp.Contains(lsParameter[i].eOCAPTemplete))
                    {
                        if (lsParameter[i].eOCAPTemplete.Equals(""))
                        {
                            cbo.Items.Add("");
                            lseOCAPtmp.Add("");
                        }
                        else
                        {
                            cbo.Items.Add(lsParameter[i].eOCAPTemplete.ToString());
                            lseOCAPtmp.Add(lsParameter[i].eOCAPTemplete.ToString());
                        }
                    }
                }
                cbo.DataSource = lseOCAPtmp;
                DGV.Columns.Insert(27, cbo);

                //DGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                DGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DGV.Columns[0].Frozen = true;
                for (int i = 0; i < DGV.ColumnCount; i++)
                {
                    DGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        private void CurrentView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex == 2) // for not edit parameter name
            {
                e.Cancel = true;
            }
        }

        private void PreviousView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex == 2) // for not edit parameter name
            {
                e.Cancel = true;
            }
        }

        private void CurrentView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void CurrentView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

        }

        private void CurrentView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                bool find = false;
                List<string> lsParameterName = new List<string>();

                for (int i = 0; i < CurrentView.RowCount; i++)
                {
                    find = false;
                    string Name = CurrentView.Rows[i].Cells["Parameter name"].Value.ToString();

                    lsParameterName.Add(Name);                   

                    for (int j = 0; j < PreviousView.RowCount; j++)
                    {                      
                                        
                        if (Name == PreviousView.Rows[j].Cells["Parameter name"].Value.ToString())
                        {
                            find = true;

                            ParameterLimitGroup Current  = lsParameter[i];//new ParameterLimitGroup();//
                            ParameterLimitGroup Previous = lsParameterHis[j]; //new ParameterLimitGroup();// 
                  
                            if (!Current.ParameterName.Equals(Previous.ParameterName))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Parameter name"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Parameter name"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.SecondFlag.Equals(Previous.SecondFlag))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Second flag"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Second flag"].Style.BackColor = Color.Yellow;
                            }

                            if (Current.StepName == null) Current.StepName = "";
                            if (Previous.StepName == null) Previous.StepName = "";
                            if (Current.StepValue == null) Current.StepValue = "";
                            if (Previous.StepValue == null) Previous.StepValue = "";

                            if (!Current.StepName.Equals(Previous.StepName))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Step name"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Step name"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.StepValue.Equals(Previous.StepValue))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Step value"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Step value"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.EventId.Equals(Previous.EventId))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Event id"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Event id"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.WarmUpTime.Equals(Previous.WarmUpTime))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Warm up time"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Warm up time"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.MonitorTime.Equals(Previous.MonitorTime))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Monitor time"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Monitor time"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.StartDetectValue.Equals(Previous.StartDetectValue))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Start detect value"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Start detect value"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.CheckMin.Equals(Previous.CheckMin))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Use lower limit"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Use lower limit"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.AllowMinEqual.Equals(Previous.AllowMinEqual))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Permit equal lower"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Permit equal lower"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.MinValue.Equals(Previous.MinValue))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Lower limit"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Lower limit"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.CheckMax.Equals(Previous.CheckMax))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Use upper limit"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Use upper limit"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.AllowMaxEqual.Equals(Previous.AllowMaxEqual))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Permit equal upper"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Permit equal upper"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.MaxValue.Equals(Previous.MaxValue))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Upper limit"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Upper limit"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.RChartFlag.Equals(Previous.RChartFlag))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Rchart flag"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Rchart flag"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.CheckRMin.Equals(Previous.CheckRMin))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Use Rchart lower limit"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Use Rchart lower limit"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.RMinValue.Equals(Previous.RMinValue))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Rchart lower limit"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Rchart lower limit"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.CheckRMax.Equals(Previous.CheckRMax))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Use Rchart upper limit"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Use Rchart upper limit"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.RMaxValue.Equals(Previous.RMaxValue))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Rchart upper limit"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Rchart upper limit"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.TerminalMsg.Equals(Previous.TerminalMsg))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Terminal msg"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Terminal msg"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.eMailAddrs.Equals(Previous.eMailAddrs))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["eMail List"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["eMail List"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.eMail.Equals(Previous.eMail))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["eMail"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["eMail"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.MCInhibit.Equals(Previous.MCInhibit))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["M/C inhibit"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["M/C inhibit"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.MCHold.Equals(Previous.MCHold))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["M/C hold"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["M/C hold"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.eOCAP.Equals(Previous.eOCAP))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["eOCAP"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["eOCAP"].Style.BackColor = Color.Yellow;
                            }

                            if (!Current.LotHold.Equals(Previous.LotHold))
                            {
                                CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].HeaderCell.Style.BackColor = Color.Yellow;
                                CurrentView.Rows[i].Cells["Lot hold"].Style.BackColor = Color.Yellow;
                                PreviousView.Rows[j].Cells["Lot hold"].Style.BackColor = Color.Yellow;
                            }

                            if (find) // find parameter, next i.
                            {
                                break;
                            }
                        }
                    }
                    if (!find)// can not find same parameter name in History, new parameter
                    {
                        CurrentView.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                        CurrentView.Rows[i].HeaderCell.Style.BackColor = Color.Yellow;
                    }
                }

                //find delete parameter
                foreach (var n in lsParameterHis)
                {
                    if (!lsParameterName.Contains(n.ParameterName))
                    {
                        for (int i = 0; i < lsParameterHis.Count; i++)
                        {
                            if (n.ParameterName == PreviousView.Rows[i].Cells["Parameter name"].Value.ToString())
                            {
                                PreviousView.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                                PreviousView.Rows[i].HeaderCell.Style.BackColor = Color.Red;
                                break;
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw;
            }            
        }
    }
}
