using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM.BOMBrowser
{
    public class RootItem : BaseItem
    {
        public RootItem(E3Project project, E3BrowserModel owner)
        {
            Part = project;
            Owner = owner;
        }
    }
}
