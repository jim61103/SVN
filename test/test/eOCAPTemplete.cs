using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iEMS_Setting
{
    public class eOCAPTemplete
    {
        public string Pkg { get; set;  }
        public string OperCode { get; set; }
        public string Templete { get; set; }
        public string CreatedTime { get; set; }
        private string Createdby = "";// { get; set; }

        public string CreateBy
        {
            get { return Createdby; }
            set { Createdby = value; }
        }
    }
}
