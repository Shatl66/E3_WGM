using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    [DataContract]
    [KnownType(typeof(E3ProjectDocument))]
    [KnownType(typeof(E3Documentation))]
    public class Document
    {
        [DataMember]
        private string _oidMaster;
        internal string oidMaster
        {
            get { return _oidMaster; }
            set { _oidMaster = value; }
        }

        [DataMember]
        private string _oid;
        internal string oid
        {
            get { return _oid; }
            set { _oid = value; }
        }
        [DataMember]
        private string _contextOid;
        internal string contextOid
        {
            get { return _contextOid; }
            set { _contextOid = value; }
        }

        [DataMember]
        private string _folder;
        internal string folder
        {
            get { return _folder; }
            set { _folder = value; }
        }
        [DataMember]
        private string _number;
        internal string number
        {
            get { return _number; }
            set { _number = value; }
        }
        [DataMember]
        private string _name;
        internal string name
        {
            get { return _name; }
            set { _name = value; }
        }

        [DataMember]
        private string _filePath;
        internal string filePath
        {
            get { return _filePath; }
            set { _filePath = value; }
        }

        [DataMember]
        private string _fileName;
        internal string fileName
        {
            get { return _fileName; }
            set { _fileName = value; }
        }

        public Document()
        {
        }

        public Document(string docNumber, string docName, string filePath, string fileName)
        {
            this.number = docNumber;
            this.name = docName;
            this.filePath = filePath;
            this.fileName = fileName;
        }
    }
}
