using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iEMS_Setting
{
    public class Parameter
    {
        public string MachModel { get; set; }
        public string MachId { get; set; }
        public string ParamName { get; set; }
        public string ParamType { get; set; }
        public byte DataType { get; set; }
        public string ParamUnit { get; set; }
        public string CreatedTime { get; set; }
        public string CreatedBy { get; set; }

        

    }    
}
