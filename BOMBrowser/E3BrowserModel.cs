using Aga.Controls.Tree;
using E3SetAdditionalPart;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Menu;

namespace E3_WGM.BOMBrowser
{
    /* Поток данных при отображении дерева:
     * 1. Пользователь открывает TreeView
     * 2. TreeView вызывает GetChildren(null) → возвращаем RootItem
     * 3. Пользователь раскрывает RootItem
     * 4. TreeView вызывает GetChildren(RootItem) → создаем дочерние элементы и вычисляем свойства для каждого элемента
     * 5. TreeView обновляет соответствующий узел
     * 6. Процесс повторяется при раскрытии каждого узла
    */
    public class E3BrowserModel : ITreeModel
    {
        private E3Assembly project;
        private Dictionary<string, List<BaseItem>> _cache = new Dictionary<string, List<BaseItem>>();

        public E3BrowserModel(E3Assembly project)
        {
            this.project = project;
        }

        public System.Collections.IEnumerable GetChildren(TreePath treePath)
        {
            List<BaseItem> items = null;

            if (treePath.IsEmpty()) // Корневой уровень
            {
                if (_cache.ContainsKey("ROOT"))
                    return _cache["ROOT"];

                items = new List<BaseItem>();
                _cache.Add("ROOT", items);

                var rootItem = new RootItem(project, this);
                items.Add(rootItem);
            }
            else // Дочерние узлы
            {
                BaseItem parent = treePath.LastNode as BaseItem;
                if (parent == null) return Enumerable.Empty<BaseItem>();
                string cacheKey = parent.ATR_E3_ENTRY;

                if (_cache.ContainsKey(cacheKey))
                    return _cache[cacheKey];

                items = new List<BaseItem>();

                try
                {
                    //if (parent is RootItem)
                    //{
                    //   items = CreateChildItemsForRoot(parent);
                    //}
                    //else if (parent is AsmItem)
                    //{
                        items = CreateChildItemsForAssembly(parent);
                    //}
                    // PartItem - без детей (или другая логика)
                }
                catch (IOException ex)
                {
                    // Логирование ошибки
                    //Debug.WriteLine($"Ошибка при создании элементов: {ex.Message}");
                    return Enumerable.Empty<BaseItem>();
                }

                // вычисляем все свойства
                //CalculatePropertiesForItems(items, parent); // TODO ??? Уже вычислили в CalculateItemProperties()

                _cache.Add(cacheKey, items);
            }
            return items;
        }

     /*
        private List<BaseItem> CreateChildItemsForRoot(BaseItem parent)
        {
            var items = new List<BaseItem>();

            foreach (E3PartUsage usage in project.Usages)
            {
                BaseItem item = CreateItemFromUsage(usage, parent);
                if (item != null)
                {
                    // вычисляем свойства для этого элемента
                    CalculateItemProperties(item, parent, usage);
                    items.Add(item);
                }
            }
            return items;
        }
     */

        private List<BaseItem> CreateChildItemsForAssembly(BaseItem parent)
        {
            var items = new List<BaseItem>();
            var assembly = parent.Part as E3Assembly;
            if (assembly?.Usages == null) return items;

            foreach (E3PartUsage usage in assembly.Usages)
            {
                BaseItem item = CreateItemFromUsage(usage, parent);
                if (item != null)
                {
                    // вычисляем свойства для этого элемента
                    CalculateItemProperties(item, parent, usage);
                    items.Add(item);
                }
            }
            return items;
        }

        private BaseItem CreateItemFromUsage(E3PartUsage usage, BaseItem parent)
        {
            if (usage.idComp != -1)
            {
                E3Part part = (E3Part)project.Parts.Find(x => x is E3Part && ((E3Part)x).ID == usage.idComp);
                if (part != null & part.isForBOM == true)
                    return new PartItem(part, parent, this);
            }
            else if (string.IsNullOrEmpty(usage.number))
            {
                if (usage.ATR_E3_ENTRY != "")
                {
                    if (usage.ATR_E3_WIRETYPE != "")
                    {
                        E3Cable cable = (E3Cable)project.Parts.Find(x => (x is E3Cable) && ((x as E3Cable).ATR_E3_ENTRY == usage.ATR_E3_ENTRY) && (x as E3Cable).ATR_E3_WIRETYPE == usage.ATR_E3_WIRETYPE);
                        if (cable != null)
                            return new PartItem(cable, parent, this);
                    }
                    else
                    {
                        E3Part part = (E3Part)project.Parts.Find(x => (x is E3Part) && ((x as E3Part).ATR_E3_ENTRY == usage.ATR_E3_ENTRY));
                        if (part != null)
                            return new PartItem((E3Part)part, parent, this);
                    }
                }
            }
            else
            {
                Part part = project.Parts.Find(x => x.number == usage.number);

                if (part != null & part.isForBOM == true)
                {
                    if (part.GetType() == typeof(AdditionalPart))
                    {
                        return new PartItem(part, parent, this);
                    }
                    else if (part is E3Part)
                    {
                        return new PartItem((E3Part)part, parent, this);
                    }
                    else if (part is E3Cable)
                    {
                        return new PartItem((E3Cable)part, parent, this);
                    }
                    else if (part is E3Assembly)
                    {
                        return new AsmItem((E3Assembly)part, parent, this);
                    }
                }
            }

            return null;
        }

        /*
        private void CalculatePropertiesForItems(List<BaseItem> items, BaseItem parent)
        {
            foreach (var item in items)
            {
                CalculateItemProperties(item, parent, GetUsageForItem(item, parent));
            }
        }
        */

        private void CalculateItemProperties(BaseItem item, BaseItem parent, E3PartUsage usage)
        {
            if (usage == null) return;

            if (item is PartItem partItem) // проверка и инициализация переменной partItem = (PartItem)item;
            {
                // partItem доступна ТОЛЬКО здесь
                partItem.Amount = usage.amount.ToString();
                partItem.Unit = usage.unit;
                partItem.RefDes = usage.RefDes;
                partItem.LineNumber = usage.lineNumber.ToString();
                partItem.Replacement = usage.Replacements.Count > 0 ? "R" : "";
                partItem.Replacements = usage.Replacements;
            }
            else if (item is AsmItem asmItem)
            {
                asmItem.Amount = usage.amount.ToString();
                asmItem.Unit = usage.unit;
                asmItem.LineNumber = usage.lineNumber.ToString();
            }
        }

        /*
        private E3PartUsage GetUsageForItem(BaseItem item, BaseItem parent)
        {
            // Ваша логика поиска usage
            // (та же, что была в ReadFilesProperties)
            if (item is PartItem)
            {
                E3Assembly parentAssembly = parent.Part as E3Assembly;
                if (parentAssembly?.Usages == null) return null;

                E3PartUsage usage = null;
                // ... ваша логика поиска ...
                return usage;
            }
            // ... и т.д.
            return null;
        }
        */




        public bool IsLeaf(TreePath treePath)
        {
            return false; // Это значит все узлы могут иметь детей (даже если их нет) => Потенциальная проблема
        }

        public event EventHandler<TreeModelEventArgs> NodesChanged;

        internal void OnNodesChanged(BaseItem item)
        {
            if (NodesChanged != null)
            {
                TreePath path = GetPath(item.Parent);
                NodesChanged(this, new TreeModelEventArgs(path, new object[] { item }));
            }
        }

        private TreePath GetPath(BaseItem item)
        {
            if (item == null)
                return TreePath.Empty;
            else
            {
                Stack<object> stack = new Stack<object>();
                while (item != null)
                {
                    stack.Push(item);
                    item = item.Parent;
                }
                return new TreePath(stack.ToArray());
            }
        }

        public event EventHandler<TreeModelEventArgs> NodesInserted;

        public event EventHandler<TreeModelEventArgs> NodesRemoved;

        public event EventHandler<TreePathEventArgs> StructureChanged;
    }
}
