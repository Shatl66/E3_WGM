using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM.BOMBrowser
{
    public class PartItem : BaseItem
    {
        public PartItem(Part part, BaseItem parent, E3BrowserModel owner)
        {
            Part = part;
            Parent = parent;
            Owner = owner;
        }

        public PartItem(E3Cable cable, BaseItem parent, E3BrowserModel owner)
        {
            Part = cable;
            Parent = parent;
            Owner = owner;
        }

        private string _refDes = "";
        public string RefDes
        {
            get { return _refDes; }
            set { _refDes = value; }
        }
    }
}
