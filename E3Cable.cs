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

        /// <summary>
        /// Создание провода
        /// </summary>
        /// <param name="wire"></param>
        public E3Cable(e3Pin wire) 
        {
            //IDs.Add(wire.GetId());
            dynamic wiregrouptype = null, wiretype = null;
            wire.GetWireType(ref wiregrouptype, ref wiretype);
            ATR_E3_ENTRY = wiregrouptype;
            ATR_E3_WIRETYPE = wiretype;
            oidMaster = wire.GetAttributeValue("WCH_id");
            number = wire.GetAttributeValue("WCH_number");
            name = wire.GetAttributeValue("WCH_name");
            ATR_BOM_RS = wire.GetAttributeValue(AttrsName.getAttrsName("atrBomRs"));            
        }

        /// <summary>
        /// Создание кабеля
        /// </summary>
        /// <param name="dev"></param>
        public E3Cable(e3Device dev, e3Component comp)
        {
            String numberDev = dev.GetAttributeValue("WCH_number");
            String nameDev = dev.GetAttributeValue("WCH_name");

            // Если dev это часть изделия Сборка Е3
            if (String.IsNullOrEmpty(numberDev)) 
            {
                numberDev = comp.GetId().ToString();
            }
            
            if (String.IsNullOrEmpty(nameDev))
            {
                nameDev = comp.GetAttributeValue("Description"); // так же у E3Part
            }



            oidMaster = dev.GetAttributeValue("WCH_id");
            number = numberDev;
            name = nameDev;
            ATR_BOM_RS = dev.GetAttributeValue(AttrsName.getAttrsName("atrBomRs"));
        }


        internal void merge(E3Cable wchE3Cable, List<string> errorMessages, e3Job job)
        {
            if (String.IsNullOrEmpty( wchE3Cable.oidMaster))
            {
                if (!errorMessages.Contains($"Кабель/провод {number} не найден в Windchill"))
                    errorMessages.Add($"Кабель/провод {number} не найден в Windchill");
                //return;
            }
            else if (!String.IsNullOrEmpty(this.oidMaster) && !String.Equals(this.oidMaster, wchE3Cable.oidMaster))
            {
                if (!errorMessages.Contains($"У {this.number} значение oidMaster не совпадает с Windchill"))
                    errorMessages.Add($"У кабеля/провода {number} значение атрибута oidMaster не совпадает с Windchill");
                //return;
            }

            this.oidMaster = wchE3Cable.oidMaster;
            this.number = wchE3Cable.number;
            this.name = wchE3Cable.name;
            this.ATR_BOM_RS = wchE3Cable.ATR_BOM_RS;
            this.ATR_E3_ENTRY = wchE3Cable.ATR_E3_ENTRY;
            this.ATR_E3_WIRETYPE = wchE3Cable.ATR_E3_WIRETYPE;

            /* может и надо
            e3Pin wire = null;
            dynamic wiregrouptype = null, wiretype = null;

            foreach (int itemId in IDs)
            {
                wire.SetId(itemId);
                wire.GetWireType(ref wiregrouptype, ref wiretype);

                if (this.ATR_E3_ENTRY != wiregrouptype || this.ATR_E3_WIRETYPE != wiretype)
                {
                    throw new Exception("ОШИБКА: Провод " + wire.GetName() + " " + wiregrouptype + " " + wiretype + " не синхронизирована с Библиотекой E3. ("+ this.number + " " + this.name + " "+ this.ATR_E3_ENTRY + " " + this.ATR_E3_WIRETYPE + ")");
                }
            }
            */
        }

    }
}
