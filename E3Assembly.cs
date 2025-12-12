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
    public class E3Assembly : Part
    {

        [DataMember]
        protected List<E3PartUsage> _usages = new List<E3PartUsage>();
        internal List<E3PartUsage> Usages
        {
            get { return _usages; }
            set { }
        }
        [DataMember]
        private List<E3PartDescribe> _describes = new List<E3PartDescribe>();
        internal List<E3PartDescribe> Describes
        {
            get { return _describes; }
            set { }
        }
        [DataMember]
        private List<E3PartReference> _references = new List<E3PartReference>();
        internal List<E3PartReference> References
        {
            get { return _references; }
            set { }
        }

        public E3Assembly(String number, string name)
        {
            this.number = number;
            this.name = name;
        }

        public E3Assembly()
        {
        }

        public E3Assembly(DataGridViewRow row)
        {
            oidMaster = (string)row.Cells["oidMaster"].Value;
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
        }

        internal void merge(E3Assembly tempE3Asm)
        {
            if (String.IsNullOrEmpty(this.oidMaster))
            {
                E3WGMForm.public_umens_e3project.updateUsages(tempE3Asm);
            }
            else if (!String.IsNullOrEmpty(oidMaster)
                && !String.IsNullOrEmpty(tempE3Asm.oidMaster)
                && !String.Equals(oidMaster, tempE3Asm.oidMaster))
            {
                throw new Exception("ОШИБКА: oidMaster сборки не совпадает с Windchill");
            }
            update(tempE3Asm);
        }

        private void update(E3Assembly tempE3Asm)
        {
            if (String.IsNullOrEmpty(this.oidMaster))
            {
                this.oidMaster = tempE3Asm.oidMaster;
            }
            else if (!String.IsNullOrEmpty(oidMaster)
                && !String.IsNullOrEmpty(tempE3Asm.oidMaster)
                && !String.Equals(oidMaster, tempE3Asm.oidMaster))
            {
                throw new Exception("ОШИБКА: oidMaster сборки не совпадает с Windchill");
            }

            this.number = tempE3Asm.number;
            this.name = tempE3Asm.name;
            this.ATR_BOM_RS = tempE3Asm.ATR_BOM_RS;

        }

        internal object[] getDataForRow()
        {
            return new Object[] {oidMaster,
                                    number,
                                    name,
                                    ATR_BOM_RS };
        }
        internal void setParam(int columnIndex, object value)
        {
            switch (columnIndex)
            {
                case 0:
                    this.oidMaster = (string)value;
                    break;
                case 1:
                    this.number = (string)value;
                    break;
                case 2:
                    this.name = (string)value;
                    break;
                case 3:
                    this.ATR_BOM_RS = (string)value;
                    break;
            }
        }

        public override bool Equals(object obj)
        {
            if (!(obj.GetType() == typeof(E3Assembly)))
            {
                return false;
            }
            return this.number == ((E3Assembly)obj).number;
        }

        internal E3PartUsage AddUsage(e3Device dev, E3Part e3Part)
        {
            E3PartUsage usage;
            if (!_usages.Exists(x => x.idComp == e3Part.ID))
            {
                usage = new E3PartUsage(e3Part);
                _usages.Add(usage);
                _usages = _usages.OrderBy(o => o.number).ToList();
            }
            else
            {
                usage = _usages.Find(x => x.idComp == e3Part.ID);
            }
            usage.addID(dev.GetId());

            if (e3Part.ATR_BOM_RS.Equals(BomRSValues.getBomRSValue((int)BomRSEnum.MATERIAL)))
            {
                usage.unit = "m";
                double amount = 0;
                // Устройство должно быть настроено в БД как Кабель !!! if (dev.IsCable() == 1 && !String.IsNullOrEmpty(dev.GetAttributeValue(AttrsName.getAttrsName("length"))))
                if (!String.IsNullOrEmpty(dev.GetAttributeValue(AttrsName.getAttrsName("length"))))
                {
                    string length = dev.GetAttributeValue("Length");
                    if (!String.IsNullOrEmpty(length))
                    {
                        if (length.Contains(" "))
                        {
                            amount = Double.Parse(length.Split(' ')[0].Replace('.', ','));
                        }
                        else
                        {
                            amount = Double.Parse(length.Replace('.', ','));
                        }
                    }
                }
                else if (!String.IsNullOrEmpty(dev.GetAttributeValue(AttrsName.getAttrsName("dlina"))))
                {
                    amount = Double.Parse(dev.GetAttributeValue(AttrsName.getAttrsName("dlina")).Replace('.', ','));
                }

                amount = amount / 1000;
                usage.AddAmount(amount);
            }
            else
            {
                usage.AddOccurrence(dev);
            }

            return usage;
        }

        internal E3PartUsage AddUsage(Part part, String localUnit)
        {
            E3PartUsage usage;

            if (part.oidMaster != null && part.oidMaster != "")
            {
                if (!_usages.Exists(x => x.oidMaster == part.oidMaster))
                {
                    usage = new E3PartUsage(part, localUnit);
                    _usages.Add(usage);
                    _usages = _usages.OrderBy(o => o.number).ToList();
                }
                else
                {
                    usage = _usages.Find(x => x.oidMaster == part.oidMaster);
                }
            }
            else
            {
                return null;
            }

            return usage;
        }

        internal void AddUsage(e3Pin pin, E3Cable cable)
        {
            E3PartUsage usage;

            if (cable.oidMaster != null && cable.oidMaster != "")
            {
                if (!_usages.Exists(x => x.oidMaster == cable.oidMaster))
                {
                    usage = new E3PartUsage(cable);
                    _usages.Add(usage);
                    _usages = _usages.OrderBy(o => o.number).ToList();
                }
                else
                {
                    usage = _usages.Find(x => x.oidMaster == cable.oidMaster);
                }
            }
            else
            {
                if (!_usages.Exists(x => x.ATR_E3_ENTRY == cable.ATR_E3_ENTRY && x.ATR_E3_WIRETYPE == cable.ATR_E3_WIRETYPE))
                {
                    usage = new E3PartUsage(cable);
                    _usages.Add(usage);
                    _usages = _usages.OrderBy(o => o.number).ToList();
                }
                else
                {
                    usage = _usages.Find(x => x.ATR_E3_ENTRY == cable.ATR_E3_ENTRY && x.ATR_E3_WIRETYPE == cable.ATR_E3_WIRETYPE);
                }
            }


            double amount = pin.GetLength();

            if (amount == 0)
            {
                string length = pin.GetAttributeValue(AttrsName.getAttrsName("cuttingLength"));
                if (length != null && length != "")
                {
                    if (length.Contains(" "))
                    {
                        amount = Double.Parse(length.Split(' ')[0].Replace('.', ','));
                    }
                    else
                    {
                        amount = Double.Parse(length.Replace('.', ','));
                    }
                }
            }

            amount = amount / 1000;

            usage.AddAmount(amount);
            usage.addID(pin.GetId());
        }

        internal void AddUsage(E3Part part, int parentIDs)
        {
            E3PartUsage usage;

            if (!String.IsNullOrEmpty(part.oidMaster))
            {
                if (!_usages.Exists(x => x.oidMaster == part.oidMaster))
                {
                    usage = new E3PartUsage(part);
                    usage.AddAmount();
                    usage.addParentID(parentIDs);
                    _usages.Add(usage);
                    _usages = _usages.OrderBy(o => o.number).ToList();
                }
                else
                {
                    usage = _usages.Find(x => x.oidMaster == part.oidMaster);
                    usage.AddAmount();
                    usage.addParentID(parentIDs);
                }
            }
        }

        internal void AddDescribe(string value, string type)
        {
            E3PartDescribe describe;

            if (!_describes.Exists(x => x.value == value))
            {
                describe = new E3PartDescribe(value, type);
                _describes.Add(describe);
            }
            else
            {
                throw new NotImplementedException();
            }

        }

        internal void AddReference(string value, string type)
        {
            E3PartReference reference;

            if (!_references.Exists(x => x.value == value))
            {
                reference = new E3PartReference(value, type);
                _references.Add(reference);
            }
            else
            {
                throw new NotImplementedException();
            }

        }

        internal void Refresh()
        {
            MemoryStream stream = new MemoryStream();
            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
            settings.UseSimpleDictionaryFormat = true;
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(E3Assembly), settings);
            ser.WriteObject(stream, this);
            stream.Position = 0;
            StreamReader sr = new StreamReader(stream);
            string jsonE3Asm = sr.ReadToEnd();
            jsonE3Asm = "{\"__type\":\"E3Assembly:#E3WGM\"," + jsonE3Asm.Substring(1);
            string jsonE3AsmFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonE3Asm, "netmarkets/jsp/by/iba/e3/http/syncE3Assembly.jsp");
            //string received = SocketClient.SendMessageFromSocket(SocketClient.ipString, 11000, 2, jsonE3Asm);
            MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonE3AsmFromWindchill));
            DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Assembly), settings);
            E3Assembly tempE3Asm = (E3Assembly)ser2.ReadObject(stream2);
            E3WGMForm.public_umens_e3project.updateUsages(tempE3Asm);
            if (!String.IsNullOrEmpty(tempE3Asm.oidMaster))
            {
                this.merge(tempE3Asm);
            }
        }

    }
}
