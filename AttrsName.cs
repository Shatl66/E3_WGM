using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    class AttrsName
    {

        public static Dictionary<string, String> dictionaryAttrsName = new Dictionary<string, string>();

        static AttrsName()
        {
            if (File.Exists("bomrs.json"))
            {
                String jsonBomRS = "";
                using (StreamReader streamReader = new StreamReader("attrsname.json"))
                {
                    jsonBomRS = streamReader.ReadToEnd();
                }

                using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonBomRS)))
                {
                    DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                    settings.UseSimpleDictionaryFormat = true;
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Dictionary<string, string>), settings);
                    dictionaryAttrsName = ser.ReadObject(stream) as Dictionary<string, string>;
                }
            }
        }

        public static String getAttrsName(string localName)
        {
            String value = "";
            dictionaryAttrsName.TryGetValue(localName, out value);
            return value;
        }
    }
}
