using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aga.Controls.Tree;
using e3;

namespace E3_WGM
{
    internal class E3StructureTreeTraverser
    {
        private e3Application app;
        private e3Job job;
        private e3StructureNode structureNode;
        private e3Tree tree;
        private E3Project umens_e3project;
        private int k = 1; // временно, удалить !

        public E3StructureTreeTraverser(E3Project umens_e3project)
        {
            this.umens_e3project = umens_e3project;
            app = ConnectToE3();
            job = CreateJobObject();
            structureNode = job.CreateStructureNodeObject(); // это пока прото пустышка заданного типа
            tree = job.CreateTreeObject(); // это пока прото пустышка заданного типа
        }

        public void FindAllWTPartsFromSelectedFolder()
        {
            int rootNodeId = 0;
            int treeId = 0;
            String treeName;

            treeId = job.GetActiveTreeID();
            if (treeId == 0)
            {
                app.PutInfo(1, "Выделите узел в дереве листов !"); // при 1 - сообщение не только выводится в окне сообщений Е3, но и всплывает окно с этим сообщением.
                return;
            }
            else
            {
                tree.SetId(treeId); // инициализировали конкретный объект tree
                treeName = tree.GetName(); // может быть "Лист", "Изделия в проекте", "Чертеж жгута", "Цепи" и др.

                if (treeName != "Лист")
                {
                    app.PutInfo(1, "Сделайте активным узел в дереве листов !");
                    return;
                }

                dynamic structureNodeIds = null;
                int structureNodeCount = job.GetTreeSelectedStructureNodeIds(ref structureNodeIds);
                if (structureNodeCount > 1)
                {
                    app.PutInfo(1, "Сделайте активным только 1 узел в дереве листов !");
                    return;
                }
                else
                {
                    rootNodeId = structureNodeIds[1];
                }
            }

            FindWTPartsRecursive( rootNodeId, null);
        }

        private void FindWTPartsRecursive(int nodeId, E3Assembly parentAssembly)
        {
            String numberPart;
            String namePart; 
            E3Assembly assembly = null;
            E3Part part = null;            

            structureNode.SetId(nodeId);

            string typeName = structureNode.GetTypeName();
            string internName = structureNode.GetInternalName();

            if (typeName == "<Assignment>" || typeName == "<Project>" || typeName == "SubProj") // У папки со схемами - .DOCUMENT_TYPE
            {
                numberPart = structureNode.GetName();
                bool hasAttribute = structureNode.HasAttribute("Naimen_izdel") == 1 ? true : false;
                namePart = hasAttribute ? structureNode.GetAttributeValue("Naimen_izdel") : "Наименование пока не известно";

                app.PutInfo(0, $"Найден узел: {numberPart} {namePart}"); // тоже самое, что и "Найден WTPart: '" + nodeName + "' (ID: " + nodeId + ")"

                if (parentAssembly == null & typeName == "<Project>")
                {
                    // Настраиваем наш объект umens_e3project на работу с данными именно от всего проекта Е3. т.к. именно сам проект выбран в дереве листов Е3

                    // {} Определить имя и обозначение для СЧ из проекта E3
                    assembly = umens_e3project;
                }
                else if (parentAssembly == null)
                {
                    // Настраиваем наш объект umens_e3project на работу с данными от выбранной папки в дереве листов Е3
                    umens_e3project.number = numberPart;
                    umens_e3project.name = namePart;
                    assembly = umens_e3project;
                }
                else if(parentAssembly != null)
                {
                    assembly = new E3Assembly(numberPart, namePart);
                    umens_e3project.Parts.Add(assembly); // накапливаем все Part-ы по 1 разу
                    parentAssembly.AddUsage(assembly);
                }


                // Получаем дочерние узлы
                dynamic childNodeIds = null;
                int childCount = structureNode.GetStructureNodeIds(ref childNodeIds);

                for (int i = 1; i <= childCount; i++)
                {
                    int childNodeId = childNodeIds[i];
                    FindWTPartsRecursive(childNodeId, assembly);
                }
            }
            else
            {
                app.PutInfo(0, "Рассчитываем содержимое узла из схем");

                // временно !
                part = new E3Part();
                part.number = "0000" + k++;
                part.name = "Temp Part" + k++;

                umens_e3project.Parts.Add( part); // накапливаем все Part-ы по 1 разу
                parentAssembly.AddUsage( part);
            }
        }



        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Возвращает объект e3Application, либо закрывает программу
        /// </summary>
        e3Application ConnectToE3()
        {
            int processCount = Process.GetProcessesByName("E3.series").Length;

            switch (processCount)
            {
                case 0:
                    MessageBox.Show("Нет запущенных окон E3.Series\n\nВыход из программы");
                    Environment.Exit(0);
                    return null; // Не выполнится, но нужен для компиляции

                case 1:
                    try
                    {
                        return (e3Application)Activator.CreateInstance(
                            Type.GetTypeFromProgID("CT.Application"));
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Не установлен E3.Series\n\nВыход из программы");
                        Environment.Exit(0);
                        return null;
                    }

                default: // Несколько процессов E3
                    try
                    {
                        dynamic dispatcher = Activator.CreateInstance(
                            Type.GetTypeFromProgID("CT.DispatcherViewer"));

                        dynamic selectedApp = null;
                        dispatcher.ShowViewer(ref selectedApp);

                        if (selectedApp == null)
                        {
                            MessageBox.Show("Не выбрано окно E3.Series\n\nВыход из программы");
                            Environment.Exit(0);
                            return null;
                        }

                        return (e3Application)selectedApp;
                    }
                    catch (Exception)
                    {
                        MessageBox.Show(
                            "Открыто несколько окон E3.Series\n" +
                            "Закройте 'лишние' окна или установите E3.Dispatcher\n\n" +
                            "Выход из программы");
                        Environment.Exit(0);
                        return null;
                    }
            }
        }

        /// <summary>
        /// Возвращает объект проекта (job) или закрывает программу
        /// </summary>
        /// <returns></returns>
        e3Job CreateJobObject()
        {
            e3Job ret = null;

            try
            {
                ret = (e3Job)app.CreateJobObject();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Невозможно создать объект проект (Job) в E3.Series" +
                    "\nСообщение: " + ex.Message +
                    "\n\nВыход из программы");
                System.Environment.Exit(0);
            }

            // нет открытых проектов - выход
            if (ret.GetId() == 0)
            {
                MessageBox.Show("Нет открытых проектов" +
                    "\n\nВыход из программы");
                System.Environment.Exit(0);
            }

            return ret;
        }


        /// <summary>
        /// <para>выбрали данные и сразу отключились от Е3, чтобы пользователю Е3 полностью вернулось управление в самом Е3</para>
        /// Доработать !!!
        /// </summary>
        public void DisconnectFromE3Series()
        {
            if (app == null)
                return;

            // разрываем соединение с E3.Series.
            Marshal.ReleaseComObject(app);
            Marshal.ReleaseComObject(job);

            // Обнуляем поля, ссылавшиеся на COM-объекты
            app = null;
            job = null;
            //dev = null;
            //Cab = null;
            //Pin = null;
            //Core = null;

            // принудительно удаляем пустые (null) ссылки
            GC.Collect();
        }

    }
}
