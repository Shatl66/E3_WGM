using e3;
using E3SetAdditionalPart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace E3_WGM
{
    [DataContract]
    public class E3PartUsage
    {

        [DataMember]
        private List<int> _ids = new List<int>();
        public List<int> IDs
        {
            get { return _ids; }
            set { }
        }
        [DataMember]
        private List<int> _parentIds = new List<int>();
        public List<int> parentIDs
        {
            get { return _parentIds; }
            set { }
        }
        [DataMember]
        private string _oidMaster = "";
        public string oidMaster
        {
            get { return _oidMaster; }
            set { _oidMaster = value; }
        }

        [DataMember]
        private string _number = "";
        public string number
        {
            get { return _number; }
            set { _number = value; }
        }

        [DataMember]
        private string _unit = "ea";
        public string unit
        {
            get { return _unit; }
            set { _unit = value; }
        }

        [DataMember]
        private double _amount = 0;
        public double amount
        {
            get { return _amount; }
            set { _amount = value; }
        }

        [DataMember]
        private int _lineNumber = 0;
        public int lineNumber
        {
            get { return _lineNumber; }
            set { _lineNumber = value; }
        }

        [DataMember]
        private string _entry = "";
        public string ATR_E3_ENTRY
        {
            get { return _entry; }
            set { _entry = value; }
        }

        [DataMember]
        private string _wiretype = "";
        public string ATR_E3_WIRETYPE
        {
            get { return _wiretype; }
            set { _wiretype = value; }
        }
        [DataMember]
        private List<E3PartOccurrence> _occurrences = new List<E3PartOccurrence>();

        public int idComp { get; set; }

        public E3PartUsage(int idComp, string number, string unit)
        {
            this.number = number;
            this.idComp = idComp;
            this._unit = unit;
        }

        public E3PartUsage(Part part)
        {
            if (part is E3Part)
            {
                this.idComp = (part as E3Part).ID;
                this.ATR_E3_ENTRY = (part as E3Part).ATR_E3_ENTRY;
            }
            else
            {
                this.idComp = -1;
            }

            this.oidMaster = part.oidMaster;
            this.number = part.number;
        }

        public E3PartUsage(Part part, String localUnit)
        {
            if (part is E3Part)
            {
                this.idComp = (part as E3Part).ID;
                this.ATR_E3_ENTRY = (part as E3Part).ATR_E3_ENTRY;
            }
            else
            {
                this.idComp = -1;
            }
            this._unit = localUnit; // было "m";
            this.oidMaster = part.oidMaster;
            this.number = part.number;
        }

        public E3PartUsage(E3Cable cable)
        {
            this.idComp = -1;
            this._unit = "m";
            this.oidMaster = cable.oidMaster;
            this.number = cable.number;
            this.ATR_E3_ENTRY = cable.ATR_E3_ENTRY;
            this.ATR_E3_WIRETYPE = cable.ATR_E3_WIRETYPE;
        }

        internal void AddOccurrence(e3Device dev)
        {
            if (!_occurrences.Exists(x => x.refDes.Equals(dev.GetName())))
            {
                _occurrences.Add(new E3PartOccurrence(dev.GetName()));
                _amount++;
            }
        }

        internal void AddAmount()
        {
            _amount++;
        }

        internal void AddAmount(double localAmount)
        {
            _amount += localAmount;
        }

        public string RefDes
        {
            get
            {
                string refDes = "";
                foreach (E3PartOccurrence occurrence in _occurrences)
                {
                    if (refDes != "")
                    {
                        refDes += ", ";
                    }
                    refDes += occurrence.refDes;

                }
                return refDes;
            }
            set
            {
            }
        }

        internal void updateUsage(Part tempPart)
        {
            if (tempPart is E3Part)
            {
                this.ATR_E3_ENTRY = (tempPart as E3Part).ATR_E3_ENTRY;
            }
            if (tempPart is E3Cable)
            {
                this.ATR_E3_ENTRY = (tempPart as E3Cable).ATR_E3_ENTRY;
                this.ATR_E3_WIRETYPE = (tempPart as E3Cable).ATR_E3_WIRETYPE;
            }
            this.oidMaster = tempPart.oidMaster;
            this.number = tempPart.number;
        }

        internal void addParentID(int id)
        {
            parentIDs.Add(id);
        }

        internal void addID(int id)
        {
            IDs.Add(id);
        }

        internal void setLineNumber(int localLineNumber)
        {
            /*this.lineNumber = localLineNumber;

            foreach (int itemId in IDs)
            {
                if (ATR_E3_WIRETYPE != null && ATR_E3_WIRETYPE != "")
                {
                    e3Pin pin = E3WGMForm.project.getJob().CreatePinObject();
                    pin.SetId(itemId);
                    if (localLineNumber != 0)
                    {
                        pin.SetAttributeValue("WCH_lineNumber", "" + localLineNumber);
                    }
                    else
                    {
                        pin.SetAttributeValue("WCH_lineNumber", null);
                    }
                } else
                {
                    e3Device dev = E3WGMForm.project.getJob().CreateDeviceObject();
                    dev.SetId(itemId);

                    if (localLineNumber != 0)
                    {
                        dev.SetAttributeValue("WCH_lineNumber", "" + localLineNumber);
                    }
                    else
                    {
                        dev.SetAttributeValue("WCH_lineNumber", null);
                    }
                } 
            }*/

            foreach (int itemId in parentIDs)
            {
                e3Device dev = E3WGMForm.public_umens_e3project.getJob().CreateDeviceObject();
                dev.SetId(itemId);

                for (int ii = 1; ii <= AdditionalPart.additionPartsMaxCount; ii++)
                {
                    String valueId = dev.GetAttributeValue("WCH_AdditionalPart0" + ii + "_id");
                    if (valueId != null && valueId != "")
                    {
                        if (this.oidMaster == valueId)
                        {
                            if (localLineNumber != 0)
                            {
                                dev.SetAttributeValue("WCH_AdditionalPart0" + ii + "_lineNumber", "" + localLineNumber);
                            }
                            else
                            {
                                dev.SetAttributeValue("WCH_AdditionalPart0" + ii + "_lineNumber", null);
                            }
                            return;
                        }
                    }
                }
            }
        }

        public bool isUsageE3Cable()
        {
            if (ATR_E3_WIRETYPE != null && ATR_E3_WIRETYPE != "")
            {
                return true;
            }
            return false;
        }

        public bool isUsageE3Part()
        {
            if (ATR_E3_ENTRY != null && ATR_E3_ENTRY != "" && !isUsageE3Cable())
            {
                return true;
            }
            return false;
        }

        public bool isUsageAdditionalPart()
        {
            if (parentIDs.Count != 0)
            {
                return true;
            }
            return false;
        }

    }
}
