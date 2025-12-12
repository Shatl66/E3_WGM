using Aga.Controls.Tree;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using E3_WGM.BOMBrowser;

namespace E3_WGM
{
    public partial class E3StructureBrowser : UserControl
    {
        public E3StructureBrowser()
        {
            InitializeComponent();
        }

        public override void Refresh()
        {
            base.Refresh();


            //cboxGrid.DataSource = System.Enum.GetValues(typeof(GridLineStyle));
            //cboxGrid.SelectedItem = GridLineStyle.HorizontalAndVertical;

            //cbLines.Checked = _treeView.ShowLines;

            _treeView.Model = new SortedTreeModel(new E3BrowserModel(E3WGMForm.public_umens_e3project));
        }
    }
}
