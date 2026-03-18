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
using System.Reflection;
using E3SetAdditionalPart;
using e3;

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
        /*
        public override void Refresh()
        {
            base.Refresh();
            _treeView.Model = new SortedTreeModel(new E3BrowserModel(E3WGMForm.public_umens_e3project));

            _treeView.ExpandAll(); // Раскроет все узлы, включая корень, если есть дети
        }
        */

        private void _treeView_NodeMouseClick(object sender, TreeNodeAdvMouseEventArgs e)
        {
            string numbers = "";

            if (e.Node != null && e.Node.Tag is PartItem partItem)
            {
                // Теперь у нас есть прямой доступ к PartItem                
                if (partItem.Replacements.Count > 0)
                {
                    List<Part> Parts = E3WGMForm.UtilsInstance.umens_e3project.Parts;

                    foreach (String numberReplacement in partItem.Replacements) {
                        // у меня в Parts все замены занесены под типом AdditionalPart
                        AdditionalPart replacementPart = (AdditionalPart)Parts.Find(x => (x is AdditionalPart) && (x as AdditionalPart).number == numberReplacement);
                        numbers = numbers + "\n" + numberReplacement + " " + replacementPart.name;                        
                    }
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
                        string jsonAssemblyFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonProject, "netmarkets/jsp/by/iba/e3/http/updateStructureE3Assembly_Umens.jsp");
                        // Обратная замена при десериализации. Правильнее было бы прописать везде - [DataContract(Namespace = "E3WGM")]
                        jsonAssemblyFromWindchill = jsonAssemblyFromWindchill.Replace("E3Assembly:#E3WGM", "E3Assembly:#E3_WGM");

                        MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonAssemblyFromWindchill));
                        DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Assembly), settings);

                        // ниже все ради номеров позиций уже расчитанных в Windchill и возвращенных сюда
                        E3Assembly assmWch = (E3Assembly)ser2.ReadObject(stream2);
                        e3Job job = E3WGMForm.UtilsInstance.ConnectToE3Series(); // нужен только job
                        ((E3Assembly)part).merge(assmWch, E3WGMForm.UtilsInstance.errorMessages, job);

                        Refresh(); // перерисовываем таблицу с ЭСИ, сразу отрисуются номера позиций полученные из Windchill

                        UmensLogger.Log($"Выгрузка структуры {part.number} в Windchill завершена");
                        
                        E3WGMForm.UtilsInstance.DisconnectFromE3Series();
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

        private void E3StructureBrowser_Load(object sender, EventArgs e)
        {
            // Сортируемые колонки задал в дизайнере формы

            // Подписываемся на клик по колонке
            _treeView.ColumnClicked += _treeView_ColumnClicked;
        }


        //TODO - запоминать модель с сортировкой и восстанавливать ее при возврате к просмотру ЭСИ
        private E3BrowserModel _baseModel;


        /// <summary>
        /// <para>Удаляет из просмотра ЭСИ если она уже отображалась.</para> 
        /// Создает модель на основе рассчитанных нами данных (public_umens_e3project) для показа ЭСИ и Начинает показывать ЭСИ
        /// </summary>
        public override void Refresh()
        {
            base.Refresh();
            _baseModel = new E3BrowserModel(E3WGMForm.public_umens_e3project);
            _treeView.Model = new SortedTreeModel(_baseModel);
            _treeView.ExpandAll();
        }

        private void _treeView_ColumnClicked(object sender, TreeColumnEventArgs e)
        {
            // Получаем текущую модель (SortedTreeModel)
            var sortedModel = _treeView.Model as SortedTreeModel;
            if (sortedModel == null) return;


            // Определяем направление сортировки: переключаем или устанавливаем по умолчанию
            SortOrder newOrder;
            if (e.Column.SortOrder == SortOrder.None || e.Column.SortOrder == SortOrder.Descending)
                newOrder = SortOrder.Ascending;
            else
                newOrder = SortOrder.Descending;

            // Сбрасываем сортировку у всех колонок, устанавливаем только у текущей
            foreach (TreeColumn col in _treeView.Columns)
                col.SortOrder = SortOrder.None;
            e.Column.SortOrder = newOrder;


            // Создаём новую SortedTreeModel с компаратором
            var newSortedModel = new SortedTreeModel(_baseModel);
            newSortedModel.Comparer = new NodeComparer(e.Column.Header , newOrder);

            // Заменяем модель
            _treeView.Model = newSortedModel;

            _treeView.ExpandAll();

        }
    }
}
