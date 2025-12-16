using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    [DataContract]
    class E3PartDescribe
    {
        public static string TypeOIDMaster = "typeOIDMaster";
        public static string TypeNumber = "typeNumber";

        [DataMember]
        private string _type;
        public string type
        {
            get { return _type; }
            set { _type = value; }
        }
        [DataMember]
        private string _value;
        public string value
        {
            get { return _value; }
            set { _value = value; }
        }


        public E3PartDescribe(string value, string type)
        {
            this.type = type;
            this.value = value;
        }

/*
        internal void updateDescribe(Document projectDoc2)
        {
            if (String.IsNullOrEmpty(projectDoc2.oidMaster))
            {
                this.type = TypeNumber;
                this.value = projectDoc2.number;
            }
            else
            {
                this.type = TypeOIDMaster;
                this.value = projectDoc2.oidMaster;
            }
        }
*/
    }
}
