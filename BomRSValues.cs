using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    enum BomRSEnum
    {
        NO = -1,
        DOCUMENTATION = 0,
        ASSEMBLY = 1,
        PART = 2,
        STANDARD = 3,
        OTHER = 4,
        MATERIAL = 5
    }
    class BomRSValues
    {

        public static Dictionary<int, String> dictionaryBomRS = new Dictionary<int, string>();

        static BomRSValues()
        {
            if (File.Exists("bomrs.json"))
            {
                String jsonBomRS = "";
                using (StreamReader streamReader = new StreamReader("bomrs.json"))
                {
                    jsonBomRS = streamReader.ReadToEnd();
                }

                using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonBomRS)))
                {
                    DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                    settings.UseSimpleDictionaryFormat = true;
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Dictionary<int, string>), settings);
                    dictionaryBomRS = ser.ReadObject(stream) as Dictionary<int, string>;
                }
            }
        }

        public static String getBomRSValue(int index)
        {
            String value = "";
            dictionaryBomRS.TryGetValue(index, out value);
            return value;
        }
    }
}
