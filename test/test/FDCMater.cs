using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace iEMS_Setting
{
    public class FDCMater
    {
        private int SerialNo = 0;// { get; set; }
        private int revision = 0;// { get; set; }
        public string Status { get; set; }
        public string Division { get; set; }
        public string OperCode { get; set; }
        public string OperName { get; set; }
        public string MachModel { get; set; }
        public string MachId { get; set; }
        public string Customer { get; set; }
        public string Package { get; set; }
        public string Dimension { get; set; }
        public string Lead { get; set; }
        public string Device { get; set; }
        public string BondingNo { get; set; }
        public string BondingRev { get; set; }
        public string RecipeName { get; set; }
        public string CreatedTime { get; set; }
        private string Createdby = "";// { get; set; }
        public string SubmittedTime { get; set; }
        private string Submittedby = "";// { get; set; }
        public string ApprovedTime { get; set; }
        private string Approvedby = "";// { get; set; }
        public string EditedTime { get; set; }
        private string Editedby = "";// { get; set; }
        public string BackupTime { get; set; }
        private string Backupby = "";// { get; set; }

        public string Memo1 { get; set; }
        public string Memo2 { get; set; }
        public string Memo3 { get; set; }

        public string EFGPSN { get; set; }

        private static long serialVersionUID = 3783245561827049253L;

        public int Serialno
        {
            get { return SerialNo; }
            set { SerialNo = value; }
        }

        public int Revision
        {
            get { return revision; }
            set { revision = value; }
        }

        public string CreatedBy
        {
            get { return Createdby; }
            set { Createdby = value; }
        }

        public string SubmittedBy
        {
            get { return Submittedby; }
            set { Submittedby = value; }
        }

        public string ApprovedBy
        {
            get { return Approvedby; }
            set { Approvedby = value; }
        }

        public string EditedBy
        {
            get { return Editedby; }
            set { Editedby = value; }
        }

        public string BackupBy
        {
            get { return Backupby; }
            set { Backupby = value; }
        }

        public object Clone()
        {
            return this.MemberwiseClone();
        }

    }
}
