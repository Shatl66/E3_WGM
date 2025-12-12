using E3_WGM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace E3SetAdditionalPart
{
    [DataContract]
    public class AdditionalPart : Part
    {

        public static String[] additionPartsSuffix = { "id", "number", "name", "length", "lineNumber" };
        public static int additionPartsMaxCount = 6;

        private double _length = 0;
        internal double length
        {
            get { return _length; }
            set { _length = value; }
        }

        private int _lineNumber = 0;
        internal int lineNumber
        {
            get { return _lineNumber; }
            set { _lineNumber = value; }
        }
        public AdditionalPart() : base()
        {
        }

        internal override object[] getDataForRow()
        {
            return new Object[] {oidMaster,
                                    number,
                                    name,
                                    ATR_BOM_RS,
                                    length,
                                    lineNumber};
        }

        internal void setLength(string value)
        {
            if (value == "")
            {
                length = 0;
            }
            else
            {
                length = Double.Parse(value.Replace('.', ','));
            }

        }

        internal void setLineNumber(string value)
        {
            if (value == "")
            {
                lineNumber = 0;
            }
            else
            {
                lineNumber = Int32.Parse(value);
            }
        }
    }

}
