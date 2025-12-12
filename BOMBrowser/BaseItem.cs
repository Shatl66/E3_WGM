using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM.BOMBrowser
{
    public abstract class BaseItem
    {
        private Image _icon = new Bitmap(Properties.Resources.e3part);
        public Image Icon
        {
            get { return _icon; }
            set { _icon = value; }
        }

        // private string _number = "";
        public string NUMBER
        {
            get { return _part.number; }
            set { }
        }

        //private string _name = "";
        public string NAME
        {
            get { return _part.name; }
            set { }
        }

        //private string _entry = "";
        public string ATR_E3_ENTRY
        {
            get
            {
                if (_part.GetType() == typeof(E3Part))
                {
                    return ((E3Part)_part).ATR_E3_ENTRY;
                }

                if (_part is E3Cable)
                {
                    return ((E3Cable)_part).ATR_E3_ENTRY + " - " + ((E3Cable)_part).ATR_E3_WIRETYPE;
                }

                return _part.number;
            }
            set { }
        }

        //private string _bomRS = "";
        public string ATR_BOM_RS
        {
            get { return _part.ATR_BOM_RS; }
            set { }
        }

        private string _amount = "";
        public string Amount
        {
            get { return _amount; }
            set { _amount = value; }
        }

        private string _unit = "";
        public string Unit
        {
            get { return _unit; }
            set { _unit = value; }
        }

        private string _lineNumber = "";
        public string LineNumber
        {
            get { return _lineNumber; }
            set { _lineNumber = value; }
        }

        private BaseItem _parent;
        public BaseItem Parent
        {
            get { return _parent; }
            set { _parent = value; }
        }

        private E3BrowserModel _owner;
        public E3BrowserModel Owner
        {
            get { return _owner; }
            set { _owner = value; }
        }

        private Part _part;
        public Part Part
        {
            get { return _part; }
            set { _part = value; }
        }
    }
}
