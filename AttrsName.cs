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
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string configFile = Path.Combine(appDir, "attrsname.json");

            if (File.Exists(configFile))
            {
                String jsonBomRS = "";
                using (StreamReader streamReader = new StreamReader(configFile))
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
