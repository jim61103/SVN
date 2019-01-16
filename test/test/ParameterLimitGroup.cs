using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iEMS_Setting
{
    public class ParameterLimitGroup
    {
        public int SerialNo { get; set; }
        public string ParameterName { get; set; }
        public int Revision { get; set; }
        public byte CheckMin { get; set; }
        public byte AllowMinEqual { get; set; }
        public double MinValue { get; set; }
        public byte CheckMax { get; set; }
        public byte AllowMaxEqual { get; set; }
        public double MaxValue { get; set; }
        public byte TerminalMsg { get; set; }
        public byte eMail { get; set; }
        public string eMailAddrs { get; set; }
        public byte MCInhibit { get; set; }
        public byte MCHold { get; set; }
        public byte eOCAP { get; set; }
        private string EOCAPTemplete = "";
        public byte LotHold { get; set; }
        public string StepName { get; set; }
        public string StepValue { get; set; }
        public int EventId { get; set; }
        public int WarmUpTime { get; set; }
        private double StartDetectvalue = 0.0;// { get; set; }
        public byte SecondFlag { get; set; }
        public byte RChartFlag { get; set; }
        public byte CheckRMin { get; set; }
        private double RMinvalue = 0.0;// { get; set; }
        public byte CheckRMax { get; set; }
        private double RMaxvalue = 0.0;// { get; set; }
        private int Monitortime = 0; 

        public string eOCAPTemplete
        {
            get { return EOCAPTemplete; }
            set { EOCAPTemplete = value; }
        }
        public double StartDetectValue
        {
            get { return StartDetectvalue; }
            set { StartDetectvalue = value; }
        }
        public double RMinValue
        {
            get { return RMinvalue; }
            set { RMinvalue = value; }
        }
        public double RMaxValue
        {
            get { return RMaxvalue; }
            set { RMaxvalue = value; }
        }
        public int MonitorTime
        {
            get { return Monitortime; }
            set { Monitortime = value;}
        }

        public object Clone()
        {
            return this.MemberwiseClone();
        }

    }
}
