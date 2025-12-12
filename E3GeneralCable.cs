using e3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E3_WGM
{
    [DataContract]
    public class E3GeneralCable : E3Part
    {
        public E3GeneralCable(e3Component comp) : base(comp)
        {

        }

        public E3GeneralCable(DataGridViewRow row) : base(row)
        {

        }

        public E3GeneralCable(String ATR_E3_ENTRY)
        {
            this.ATR_E3_ENTRY = ATR_E3_ENTRY;
        }

        internal override void Refresh()
        {
            if (ID != 0)
            {
                base.Refresh();
            }
            else
            {
                MemoryStream stream = new MemoryStream();
                DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                settings.UseSimpleDictionaryFormat = true;
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(E3GeneralCable), settings);
                ser.WriteObject(stream, this);
                stream.Position = 0;
                StreamReader sr = new StreamReader(stream);
                string jsonE3GeneralCable = sr.ReadToEnd();
                jsonE3GeneralCable = "{\"__type\":\"E3GeneralCable:#E3WGM\"," + jsonE3GeneralCable.Substring(1);
                string jsonE3GeneralCableFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonE3GeneralCable, "netmarkets/jsp/by/iba/e3/http/findE3GeneralCable.jsp");
                //string received = SocketClient.SendMessageFromSocket(SocketClient.ipString, 11000, 13, jsonE3GeneralCable);

                MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonE3GeneralCableFromWindchill));
                DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3GeneralCable), settings);
                E3GeneralCable tempE3GeneralCable = (E3GeneralCable)ser2.ReadObject(stream2);

                if (!String.IsNullOrEmpty(tempE3GeneralCable.oidMaster))
                {
                    this.merge(tempE3GeneralCable);
                }
            }
        }
    }
}
