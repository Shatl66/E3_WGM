using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    [DataContract]
    public class Part
    {
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

        /// <summary>
        ///  если в проекте Изелию в "Раздел спецификации" указали - "Отсутствует", то в BOM такую СЧ не включаем
        /// </summary>
        [DataMember]
        private bool _isForBOM = true;
        public bool isForBOM
        {
            get { return _isForBOM; }
            set { _isForBOM = value; }
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
            if( String.IsNullOrEmpty( oidMaster))  // так будет у additionalPart, т.к. они созданы только на основании Обозначения занесенного в атрибут Изделия
            {
                this.oidMaster = tempPart.oidMaster;
                this.oid = tempPart.oid;
                this.wchcheckout = tempPart.wchcheckout;
            }
                

            this.number = tempPart.number;
            this.name = tempPart.name;
            this.ATR_BOM_RS = tempPart.ATR_BOM_RS;
        }
    }
}
