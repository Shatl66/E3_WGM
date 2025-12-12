using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    public class Part
    {
        //TODO можно использовать свойство с явным именем [DataMember(Name = "_oidMaster")] и тогда оставить только - public string oidMaster { get; set; }
        [DataMember]
        private string _oidMaster = "";
        internal string oidMaster
        {
            get { return _oidMaster; }
            set { _oidMaster = value; }
        }

        [DataMember]
        private string _oid = "";
        protected string oid
        {
            get { return _oid; }
            set { _oid = value; }
        }

        [DataMember]
        private bool _wchcheckout = true;
        public bool wchcheckout
        {
            get { return _wchcheckout; }
            set { _wchcheckout = value; }
        }

        [DataMember]
        private string _number = "";
        public string number
        {
            get { return _number; }
            set { _number = value; }
        }

        [DataMember]
        private string _name = "";
        public string name
        {
            get { return _name; }
            set { _name = value; }
        }

        [DataMember]
        private string _bomRS = "";
        public string ATR_BOM_RS
        {
            get { return _bomRS; }
            set { _bomRS = value; }
        }

        public Part()
        {
        }

        internal virtual object[] getDataForRow()
        {
            return new Object[] {oidMaster,
                                    number,
                                    name,
                                    ATR_BOM_RS };

        }

        internal void merge(Part tempPart)
        {
            this.number = tempPart.number;
            this.name = tempPart.name;
            this.ATR_BOM_RS = tempPart.ATR_BOM_RS;
        }
    }
}
