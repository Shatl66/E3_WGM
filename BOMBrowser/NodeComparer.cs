using Aga.Controls.Tree;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace E3_WGM.BOMBrowser
{
    public class NodeComparer : System.Collections.IComparer
    {
        private readonly int _columnIndex;
        private readonly String _columnHeader;
        private readonly SortOrder _sortOrder;

        public NodeComparer(String columnHeader, SortOrder sortOrder)
        {
            //_columnIndex = columnIndex;
            _columnHeader = columnHeader; // плюс перед columnIndex - если колонки поменяю местами в форме, минус - если переименую шапку колонки
            _sortOrder = sortOrder;
        }

        // SortedTreeModel при каждом запросе детей узла (метод GetChildren) получает список объектов от базовой модели (E3BrowserModel),
        // а затем сортирует этот список с помощью компаратора.
        public int Compare(object x, object y)
        {
            var itemX = x as PartItem;
            var itemY = y as PartItem;

            if (itemX == null && itemY == null) return 0;
            if (itemX == null) return -1;
            if (itemY == null) return 1;

            int result = 0;

            switch (_columnHeader)
            {
                case "Обозначение":
                    result = string.Compare(itemX.NUMBER, itemY.NUMBER, StringComparison.CurrentCulture);
                    break;
                case "Наименование":
                    result = string.Compare(itemX.NAME, itemY.NAME, StringComparison.CurrentCulture);
                    break;
                case "Раздел спецификации":
                    result = string.Compare(itemX.ATR_BOM_RS, itemY.ATR_BOM_RS, StringComparison.CurrentCulture);
                    break;
                case "Количество":
                    result = CompareAmount(itemX.Amount, itemY.Amount);
                    break;
                case "ЕИ":
                    result = string.Compare(itemX.Unit, itemY.Unit, StringComparison.CurrentCulture);
                    break;
                case "Позиция": // LineNumber
                    result = CompareNumeric(itemX.LineNumber, itemY.LineNumber);
                    break;
                case "Позиц. обозначение": 
                    result = string.Compare(itemX.RefDes, itemY.RefDes);
                    break;
                case "Имя в БД Е3":
                    result = string.Compare(itemX.ATR_E3_ENTRY, itemY.ATR_E3_ENTRY);
                    break;
                case "": // Подстановка
                    result = string.Compare(itemX.Replacement, itemY.Replacement);
                    break;
                default:
                    result = 0;
                    break;
            }


            // Учитываем направление сортировки
            if (_sortOrder == SortOrder.Descending)
                result = -result;

            return result;
        }

        private int CompareAmount(string a, string b)
        {
            // Попытка сравнить как double, иначе строковое сравнение
            if (double.TryParse(a, out double da) && double.TryParse(b, out double db))
                return da.CompareTo(db);

            return string.Compare(a, b, StringComparison.CurrentCulture);
        }

        private int CompareNumeric(string a, string b)
        {
            // Попытка сравнить как int, иначе строковое сравнение
            if (int.TryParse(a, out int ia) && int.TryParse(b, out int ib))
                return ia.CompareTo(ib);

            return string.Compare(a, b, StringComparison.CurrentCulture);
        }
    }
}
