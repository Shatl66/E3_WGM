using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Windows.Forms;
using e3;

namespace E3_WGM
{
    [DataContract]
    public class E3Cable : Part
    {
        [DataMember]
        private List<int> _ids = new List<int>();
        public List<int> IDs
        {
            get { return _ids; }
            set { }
        }

        [DataMember]
        private string _entry;
        public string ATR_E3_ENTRY
        {
            get { return _entry; }
            set { _entry = value; }
        }

        [DataMember]
        private string _wiretype;
        public string ATR_E3_WIRETYPE
        {
            get { return _wiretype; }
            set { _wiretype = value; }
        }

        /*[DataMember]
        private string _class;
        public string ATR_E3_CLASS
        {
            get { return _class; }
            set { _class = value; }
        }*/

        private SortedDictionary<object, E3PartUsage> _usages = new SortedDictionary<object, E3PartUsage>();
        public SortedDictionary<object, E3PartUsage> Usages
        {
            get { return _usages; }
            set { _usages = value; }
        }

        public E3Cable(e3Pin wire)
        {
            IDs.Add(wire.GetId());
            dynamic wiregrouptype = null, wiretype = null;
            wire.GetWireType(ref wiregrouptype, ref wiretype);
            ATR_E3_ENTRY = wiregrouptype;
            ATR_E3_WIRETYPE = wiretype;
            //ATR_E3_CLASS = wire.GetAttributeValue("Class");
            oidMaster = wire.GetAttributeValue("WCH_id");
            number = wire.GetAttributeValue("WCH_number");
            name = wire.GetAttributeValue("WCH_name");
            ATR_BOM_RS = wire.GetAttributeValue(AttrsName.getAttrsName("atrBomRs"));            
        }

        public E3Cable(DataGridViewRow row)
        {
            oidMaster = (string)row.Cells["oidMaster"].Value;
            // IDs = ((String)row.Cells["ID"].Value).Split(',').ToList;
            number = (string)row.Cells["number"].Value;
            name = (string)row.Cells["name"].Value;
            if (name == null || name == "")
            {
                throw new Exception("ОШИБКА: Наименование не заполнено.");
            }
            ATR_BOM_RS = (string)row.Cells["ATR_BOM_RS"].Value;
            if (ATR_BOM_RS == null || ATR_BOM_RS == "")
            {
                throw new Exception("ОШИБКА: Раздел спецификации не заполнен.");
            }
            ATR_E3_ENTRY = (string)row.Cells["ATR_E3_ENTRY"].Value;
            ATR_E3_WIRETYPE = (string)row.Cells["ATR_E3_WIRETYPE"].Value;
            // ATR_E3_CLASS = (string)row.Cells["ATR_E3_CLASS"].Value;
        }

        internal void merge(E3Cable tempE3Cable)
        {
            if (!String.IsNullOrEmpty(this.oidMaster) && !String.IsNullOrEmpty(tempE3Cable.oidMaster) && !String.Equals(this.oidMaster, tempE3Cable.oidMaster))
            {
                throw new Exception("ОШИБКА: oidMaster провода " + this.ATR_E3_ENTRY + " " + this.ATR_E3_WIRETYPE + " не совпадает с Windchill");
            }
            this.oidMaster = tempE3Cable.oidMaster;
            this.number = tempE3Cable.number;
            this.name = tempE3Cable.name;
            this.ATR_BOM_RS = tempE3Cable.ATR_BOM_RS;
            this.ATR_E3_ENTRY = tempE3Cable.ATR_E3_ENTRY;
            this.ATR_E3_WIRETYPE = tempE3Cable.ATR_E3_WIRETYPE;
            // this.ATR_E3_CLASS = tempE3Cable.ATR_E3_CLASS;

            e3Pin wire = null; // E3WGMForm.public_umens_e3project.getJob().CreatePinObject();
            dynamic wiregrouptype = null, wiretype = null;

            foreach (int itemId in IDs)
            {
                wire.SetId(itemId);
                wire.GetWireType(ref wiregrouptype, ref wiretype);

                if (this.ATR_E3_ENTRY != wiregrouptype || this.ATR_E3_WIRETYPE != wiretype)
                {
                    throw new Exception("ОШИБКА: Провод " + wire.GetName() + " " + wiregrouptype + " " + wiretype + " не синхронизирована с Библиотекой E3. ("+ this.number + " " + this.name + " "+ this.ATR_E3_ENTRY + " " + this.ATR_E3_WIRETYPE + ")");
                }

                wire.SetAttributeValue("WCH_id", this.oidMaster);
                wire.SetAttributeValue("WCH_number", this.number);
                wire.SetAttributeValue("WCH_name", this.name);
                wire.SetAttributeValue(AttrsName.getAttrsName("atrBomRs"), this.ATR_BOM_RS);
            }
        }

        public override bool Equals(object obj)
        {
            if (!(obj.GetType() == typeof(E3Cable)))
            {
                return false;
            }
            return this.IDs.Contains(((E3Part)obj).ID);
        }

        internal void Refresh()
        {

            MemoryStream stream = new MemoryStream();
            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
            settings.UseSimpleDictionaryFormat = true;
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(E3Cable), settings);
            ser.WriteObject(stream, this);
            stream.Position = 0;
            StreamReader sr = new StreamReader(stream);
            string jsonE3Cable = sr.ReadToEnd();
            jsonE3Cable = "{\"__type\":\"E3Cable:#E3WGM\"," + jsonE3Cable.Substring(1);
            string jsonE3CableFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonE3Cable, "netmarkets/jsp/by/iba/e3/http/findE3Cable.jsp");
            //string received = SocketClient.SendMessageFromSocket(SocketClient.ipString, 11000, 14, jsonE3Cable);
            MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonE3CableFromWindchill));
            DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Cable), settings);
            E3Cable tempE3Cable = (E3Cable)ser2.ReadObject(stream2);
            if (!String.IsNullOrEmpty(tempE3Cable.oidMaster))
            {
                this.merge(tempE3Cable);
            }
        }

        internal object[] getDataForRow()
        {
            return new Object[] {oidMaster,
                               //     IDs.ToString(),
                                    number,
                                    name,
                                    ATR_BOM_RS,
                                    ATR_E3_ENTRY,
                                    ATR_E3_WIRETYPE/*,
                                    ATR_E3_CLASS*/ };
        }

    }
}
