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
    public class E3Part : Part
    {
        [DataMember]
        private int _id = -1;
        public int ID
        {
            get { return _id; }
            set { _id = value; }
        }

        [DataMember]
        private string _entry;
        public string ATR_E3_ENTRY
        {
            get { return _entry; }
            set { _entry = value; }
        }

        [DataMember]
        private string _class;
        public string ATR_E3_CLASS
        {
            get { return _class; }
            set { _class = value; }
        }

        private SortedDictionary<object, E3PartUsage> _usages = new SortedDictionary<object, E3PartUsage>();
        public SortedDictionary<object, E3PartUsage> Usages
        {
            get { return _usages; }
            set { _usages = value; }
        }

        public E3Part()
        {

        }

        public E3Part(e3Component comp)
        {
            oidMaster = comp.GetAttributeValue("WCH_id");
            number = comp.GetAttributeValue("WCH_number");
            if (number == null || number == "")
            {
                number = comp.GetAttributeValue("DOCnum");
            }
            name = comp.GetAttributeValue("WCH_name");
            if (name == null || name == "")
            {
                name = comp.GetAttributeValue("Description");
            }
            ATR_BOM_RS = comp.GetAttributeValue(AttrsName.getAttrsName("atrBomRs"));
            ATR_E3_ENTRY = comp.GetName();
            ATR_E3_CLASS = comp.GetAttributeValue("Class");
            ID = comp.GetId();
        }

        internal void merge(E3Part tempE3Part)
        {
            this.oidMaster = tempE3Part.oidMaster;
            this.number = tempE3Part.number;
            this.name = tempE3Part.name;
            this.ATR_BOM_RS = tempE3Part.ATR_BOM_RS;
            this.ATR_E3_ENTRY = tempE3Part.ATR_E3_ENTRY;
            this.ATR_E3_CLASS = tempE3Part.ATR_E3_CLASS;
        }

        public E3Part(DataGridViewRow row)
        {
            oidMaster = (string)row.Cells["oidMaster"].Value;
            ID = (int)row.Cells["ID"].Value;
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
            ATR_E3_CLASS = (string)row.Cells["ATR_E3_CLASS"].Value;
        }

        public override bool Equals(object obj)
        {
            if (!(obj.GetType() == typeof(E3Part)))
            {
                return false;
            }
            return this.ID == ((E3Part)obj).ID;
        }

        internal virtual void Refresh()
        {
            /*
            e3Component comp = E3WGMForm.public_umens_e3project.getJob().CreateComponentObject();

            //У дополнительных частей _id=0
            if (_id == -1)
            {
                return;
            }

            comp.SetId(_id);
            if (this.oidMaster != null && this.oidMaster != "")
            {
                if (this.oidMaster != comp.GetAttributeValue("WCH_id") ||
                        this.number != comp.GetAttributeValue("WCH_number") ||
                        this.name != comp.GetAttributeValue("Description") ||
                        this.ATR_BOM_RS != comp.GetAttributeValue(AttrsName.getAttrsName("atrBomRs")) ||
                        this.ATR_E3_ENTRY != comp.GetName() ||
                        this.ATR_E3_CLASS != comp.GetAttributeValue("Class"))
                {
                    throw new Exception("ОШИБКА: Библиотека E3 (" + comp.GetAttributeValue("WCH_id") + ", " +
                        comp.GetAttributeValue("WCH_number") + ", " +
                        comp.GetAttributeValue("Description") + ", " +
                        comp.GetAttributeValue(AttrsName.getAttrsName("atrBomRs")) + ", " +
                        comp.GetName() + ", " +
                        comp.GetAttributeValue("Class") + ") не синхронизирована с Windchill (" + this.oidMaster + ", " +
                        this.number + ", " +
                        this.name + ", " +
                        this.ATR_BOM_RS + ", " +
                        this.ATR_E3_ENTRY + ", " +
                        this.ATR_E3_CLASS + ").");
                }
            }
            else
            {
                this.oidMaster = comp.GetAttributeValue("WCH_id");
                this.number = comp.GetAttributeValue("WCH_number");
                this.name = comp.GetAttributeValue("Description");
                this.ATR_BOM_RS = comp.GetAttributeValue(AttrsName.getAttrsName("atrBomRs"));
                this.ATR_E3_ENTRY = comp.GetName();
                this.ATR_E3_CLASS = comp.GetAttributeValue("Class");
            }
            */
        }

        internal object[] getDataForRow() //TODO добавить override ?
        {
            return new Object[] {oidMaster,
                                    ID,
                                    number,
                                    name,
                                    ATR_BOM_RS,
                                    ATR_E3_ENTRY,
                                    ATR_E3_CLASS };
        }
    }
}
