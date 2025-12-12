using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    [DataContract]
    public class E3PartOccurrence
    {
        [DataMember]
        private string _oid;
        [DataMember]
        private string _refDes;
        public string refDes
        {
            get { return _refDes; }
            set { _refDes = value; }
        }
        public E3PartOccurrence(string refDes)
        {
            this._refDes = refDes;
        }

        public override string ToString()
        {
            return _refDes;
        }
    }
}
