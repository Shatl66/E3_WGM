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
using System.IO;
using System.Runtime.Serialization.Json;

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
                    MessageBox.Show(numbers, "Информация о подстановках");
                }
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

        private void buttonUploadStructure_Click(object sender, EventArgs e)
        {
            try
            {
                //    throw new Exception("ОШИБКА: Не все составные части созданы в Windchill") я им уже сообщил об этом. Тут повторно не сообщаю, Windchill сам вернет ошибку

                foreach (Part part in E3WGMForm.public_umens_e3project.Parts)
                {
                    if (part is E3Assembly)
                    {
                        MemoryStream stream = new MemoryStream();
                        DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                        settings.UseSimpleDictionaryFormat = true;
                        DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(E3Assembly), settings);
                        ser.WriteObject(stream, (E3Assembly)part);
                        stream.Position = 0;
                        StreamReader sr = new StreamReader(stream);
                        string jsonProject = sr.ReadToEnd();
                        //В классах Windchill Андреем прописано пространство имен E3WGM. Я пока использую эти же классы, поэтому нужно сопоставлять E3WGM и мое E3_WGM
                        jsonProject = "{\"__type\":\"E3Assembly:#E3WGM\"," + jsonProject.Substring(1);
                        string jsonProjectFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonProject, "netmarkets/jsp/by/iba/e3/http/updateStructureE3Assembly_Umens.jsp");
                        // Обратная замена при десериализации. Правильнее было бы прописать везде - [DataContract(Namespace = "E3WGM")]
                        //jsonProjectFromWindchill = jsonProjectFromWindchill.Replace("E3Assembly:#E3WGM", "E3Assembly:#E3_WGM");

                        UmensLogger.Log($"Выгрузка структуры {part.number} в Windchill завершена");
                    }
                }
                
                UmensLogger.Log($"Выгрузка структуры всего проекта {E3WGMForm.public_umens_e3project.number} в Windchill завершена");
               // buttonUploadStructure.Enabled = false; // не даем повторно отправить данные. После синхронизации кнопка опять станет доступной
            }
            catch (Exception ex)
            {
                UmensLogger.Log($"Выгрузить структуру. Сообщение Windchill: {ex.Message}");
            }

        }
    }
}
