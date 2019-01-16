using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iEMS_Setting
{

    class Config
    {
        public string SiteID { get; set; }
        public string FactoryID { get; set; }
        public string LineID { get; set; }
        public string ProductionLine { get; set; }
        public string Environment { get; set; }
        public string ConfigServer { get; set; }
        public string ConfigServerID { get; set; }
        public string ConfigServerPWD { get; set; }
        public string ConfigFolder { get; set; }
        public string ConfigFile { get; set; }
        public string LocalConfigPath { get; set; }
    }
}
