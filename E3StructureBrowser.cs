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

        /// <summary>
        /// <para>Удаляет из просмотра ЭСИ если она уже отображалась.</para> 
        /// Создает модель на основе рассчитанных нами данных (public_umens_e3project) для показа ЭСИ и Начинает показывать ЭСИ
        /// </summary>
        public override void Refresh()
        {
            base.Refresh();


            //cboxGrid.DataSource = System.Enum.GetValues(typeof(GridLineStyle));
            //cboxGrid.SelectedItem = GridLineStyle.HorizontalAndVertical;

            //cbLines.Checked = _treeView.ShowLines;

            _treeView.Model = new SortedTreeModel(new E3BrowserModel(E3WGMForm.public_umens_e3project));
        }

        private void _treeView_NodeMouseClick(object sender, TreeNodeAdvMouseEventArgs e)
        {
            string numbers = "";

            if (e.Node != null && e.Node.Tag is PartItem partItem)
            {
                // Теперь у нас есть прямой доступ к PartItem                
                if (partItem.Replacements.Count > 0)
                {
                    numbers = string.Join("\n", partItem.Replacements);
                }
                else
                {
                    numbers = "Подстановки отсутствуют";
                }
                MessageBox.Show( numbers, "Информация о подстановках");
            }

            /*
            else if (e.Node != null && e.Node.Tag is AsmItem asmItem)
            {
                // Аналогично для сборок
                MessageBox.Show($"Выбрана сборка: {asmItem.NUMBER}", "Информация о сборке");
            }
            else if (e.Node != null && e.Node.Tag is RootItem rootItem)
            {
                // Аналогично для корня проекта
                MessageBox.Show($"Выбран проект: {rootItem.NUMBER}", "Информация о проекте");
            }
            */
        }
    }
}
