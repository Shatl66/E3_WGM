using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    [DataContract]
    public class E3ProjectDocument : Document
    {
        internal bool differentNumbers = false;

        public E3ProjectDocument(string docNumber, string docName, string filePath, string fileName) : base(docNumber, docName, filePath, fileName)
        {
        }

        internal void updateDoc(E3ProjectDocument projectDoc2)
        {
            this.oidMaster = projectDoc2.oidMaster;
            this.oid = projectDoc2.oid;
            this.contextOid = projectDoc2.contextOid;
            this.folder = projectDoc2.folder;
            this.number = projectDoc2.number;
            this.name = projectDoc2.name;
        }

        internal void AddFilePath(string filePath)
        {
            this.filePath = filePath;
        }
    }
}
