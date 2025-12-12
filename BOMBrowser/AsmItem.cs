using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM.BOMBrowser
{
    public class AsmItem : BaseItem
    {
        public AsmItem(E3Assembly part, BaseItem parent, E3BrowserModel owner)
        {
            Part = part;
            Parent = parent;
            Owner = owner;
        }
    }
}
