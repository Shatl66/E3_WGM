using e3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E3_WGM
{
    [DataContract]
    public class E3Assembly : Part
    {
        internal List<Part> Parts = new List<Part>();
        internal List<Document> Docs = new List<Document>();

        [DataMember]
        protected List<E3PartUsage> _usages = new List<E3PartUsage>();
        internal List<E3PartUsage> Usages
        {
            get { return _usages; }
            set { }
        }
        [DataMember]
        private List<E3PartDescribe> _describes = new List<E3PartDescribe>();
        internal List<E3PartDescribe> Describes
        {
            get { return _describes; }
            set { }
        }
        [DataMember]
        private List<E3PartReference> _references = new List<E3PartReference>();
        internal List<E3PartReference> References
        {
            get { return _references; }
            set { }
        }

        public E3Assembly(String number, string name)
        {
            this.number = number;
            this.name = name;
            ATR_BOM_RS = BomRSValues.getBomRSValue((int)BomRSEnum.ASSEMBLY); // TODO А если Комплект ?
        }

        public E3Assembly()
        {
        }

        public E3Assembly(DataGridViewRow row)
        {
            oidMaster = (string)row.Cells["oidMaster"].Value;
            number = (string)row.Cells["number"].Value;
            name = (string)row.Cells["name"].Value;
            if (name == null || name == "")
            {
                throw new Exception("ОШИБКА: Наименование не заполнено.");
            }
            ATR_BOM_RS = (string)row.Cells["ATR_BOM_RS"].Value;
            if (ATR_BOM_RS == null || ATR_BOM_RS == "")
            {
                throw new Exception("ОШИБКА: Раздел спецификации не заполнен.");
            }
        }

        /// <summary>
        /// Обновляет полученными из Windchill данными все объекты нашей Е3 сборки.
        /// Наполняет список ошибок объектами которые не были найдены в Windchill
        /// Формирует значение атрибута "суммарная позиция" для каждого обекта нашей Е3 сборки.
        /// </summary>
        /// <param name="asmWch"></param>
        /// <param name="errorMessages"></param>
        /// <param name="job"></param>
        internal void merge(E3Assembly asmWch, List<string> errorMessages, e3Job job)
        {
            if (String.IsNullOrEmpty(asmWch.oidMaster))
            {
                if (!errorMessages.Contains($"Изделие {number} не найдено в Windchill"))
                    errorMessages.Add($"Изделие {number} не найдено в Windchill");
                //return;
            }
            else if (!String.IsNullOrEmpty(oidMaster) && !String.IsNullOrEmpty(asmWch.oidMaster) && !String.Equals(oidMaster, asmWch.oidMaster))
            {
                if (!errorMessages.Contains($"У {number} значение oidMaster не совпадает с Windchill"))
                    errorMessages.Add($"У {number} значение атрибута oidMaster не совпадает с Windchill");
                //return;
            }

            update(asmWch);
            updateUsages(asmWch, errorMessages);
            updateSumPosInUsages(job);
            updateSumPosInUsagesForNetSegments(job);
        }

        private void update(E3Assembly asmWch)
        {
            this.oidMaster = asmWch.oidMaster;
            this.number = asmWch.number;
            this.name = asmWch.name;
            this.ATR_BOM_RS = asmWch.ATR_BOM_RS;

        }


        /// <summary>
        /// Находит usage и наращивает количество
        /// </summary>
        /// <param name="dev"></param>
        /// <param name="e3Part"></param>
        /// <param name="segmentManufLength"></param>
        internal void AddUsageLenght(e3Device dev, E3Part e3Part, double segmentManufLength)
        {
            if (!e3Part.ATR_BOM_RS.Equals(BomRSValues.getBomRSValue((int)BomRSEnum.MATERIAL)))
                return;

            E3PartUsage usage = _usages.Find(x => x.idComp == e3Part.ID); // обязан уже быть
            double amount = segmentManufLength / 1000;
            usage.AddAmount(amount);

        }


        /// <summary>
        /// Создает для текущей assembly новый объект связи E3PartUsage (подобное у Windchill - WTPartUsageLink), если такого еще нет,
        /// или находит имеющийся и если надо, то наращивает у него Количество и Occurrence
        /// <para>deltaAmount возвращает на сколько увеличилось количество Изделия для ЭСИ. Нужна для расчета количества AdditionalParts этого Изделия</para>
        /// </summary>
        /// <param name="dev"></param>
        /// <param name="e3Part"></param>
        /// <param name="deltaAmount"></param>
        /// <param name="incrAmount">соответствует RS изделия. false - значит RS "отсутствует"</param>
        /// <returns></returns>
        internal E3PartUsage AddUsage(e3Device dev, E3Part e3Part, out double deltaAmount, Boolean incrAmount)
        {
            E3PartUsage usage;
            deltaAmount = 1;
            if (!_usages.Exists(x => x.idComp == e3Part.ID))
            {
                usage = new E3PartUsage(e3Part);
                _usages.Add(usage);
                _usages = _usages.OrderBy(o => o.number).ToList();
            }
            else
            {
                usage = _usages.Find(x => x.idComp == e3Part.ID);
            }

            usage.addID(dev.GetId());


            if (e3Part.ATR_BOM_RS.Equals(BomRSValues.getBomRSValue((int)BomRSEnum.MATERIAL)))
            {

            }
            else
            {
                double startAmount = usage.amount;

                // 1. Cначала определяем значение deltaAmount, а оно не зависит от RS ("Отсутствует" или другое)
                usage.AddOccurrence(dev);
                double currentAmount = usage.amount;
                deltaAmount = currentAmount - startAmount;

                // 2. Проверяем на RS "отсутствует" и если так, то возвращаем исходное количество изделия и occurence
                if (incrAmount == false)
                {
                    usage.amount = startAmount;
                    usage.RemoveOccurrence(dev);
                }
            }

            return usage;
        }

        internal E3PartUsage AddOrGetUsage(Part part, String localUnit)
        {
            E3PartUsage usage;

            if (!string.IsNullOrEmpty(part.number))
            {
                if (!_usages.Exists(x => x.number == part.number))
                {
                    usage = new E3PartUsage(part, localUnit);
                    _usages.Add(usage);
                    _usages = _usages.OrderBy(o => o.number).ToList();
                }
                else
                {
                    usage = _usages.Find(x => x.number == part.number);
                }
            }
            else
            {
                return null;
            }

            return usage;
        }

        /// <summary>
        /// Создает E3PartUsage если его нет и добавляет его в _usages сборки или находит его в _usages
        /// </summary>
        /// <param name="pin"></param>
        /// <param name="cable"></param>
        /// <returns>E3PartUsage вновь созданный или найденный</returns>
        internal E3PartUsage AddOrGetUsage(e3Pin pin, E3Cable cable)
        {
            E3PartUsage usage;

            if (!string.IsNullOrEmpty( cable.number))
            {
                if (!_usages.Exists(x => x.number == cable.number))
                {
                    usage = new E3PartUsage(cable);
                    _usages.Add(usage);
                    _usages = _usages.OrderBy(o => o.number).ToList();
                }
                else
                {
                    usage = _usages.Find(x => x.number == cable.number);
                }
            }
            else
            {
                if (!_usages.Exists(x => x.ATR_E3_ENTRY == cable.ATR_E3_ENTRY && x.ATR_E3_WIRETYPE == cable.ATR_E3_WIRETYPE))
                {
                    usage = new E3PartUsage(cable);
                    _usages.Add(usage);
                    _usages = _usages.OrderBy(o => o.number).ToList();
                }
                else
                {
                    usage = _usages.Find(x => x.ATR_E3_ENTRY == cable.ATR_E3_ENTRY && x.ATR_E3_WIRETYPE == cable.ATR_E3_WIRETYPE);
                }
            }

            return usage;
        }

        /*
        internal void AddUsage(E3Part part, int parentIDs)
        {
            E3PartUsage usage;

            if (!String.IsNullOrEmpty(part.oidMaster))
            {
                if (!_usages.Exists(x => x.oidMaster == part.oidMaster))
                {
                    usage = new E3PartUsage(part);
                    usage.AddAmount();
                    usage.addParentID(parentIDs);
                    _usages.Add(usage);
                    _usages = _usages.OrderBy(o => o.number).ToList();
                }
                else
                {
                    usage = _usages.Find(x => x.oidMaster == part.oidMaster);
                    usage.AddAmount();
                    usage.addParentID(parentIDs);
                }
            }
        }
        */

        internal void AddUsage(Part part)
        {
            E3PartUsage usage;

            usage = new E3PartUsage(part);
            usage.AddAmount();
            _usages.Add(usage);
            _usages = _usages.OrderBy(o => o.number).ToList();
        }

        internal void AddDescribe(string value, string type)
        {
            E3PartDescribe describe;

            if (!_describes.Exists(x => x.value == value))
            {
                describe = new E3PartDescribe(value, type);
                _describes.Add(describe);
            }
            else
            {
                throw new NotImplementedException();
            }

        }

        internal void AddReference(string value, string type)
        {
            E3PartReference reference;

            if (!_references.Exists(x => x.value == value))
            {
                reference = new E3PartReference(value, type);
                _references.Add(reference);
            }
            else
            {
                throw new NotImplementedException();
            }

        }


        /// <summary>
        /// Обновляет у нашей текущей Е3 сборки данные ее связей (все E3PartUsage) данными полученными из Windchill
        /// Выполняет проверки и наполняет errorMessages
        /// </summary>
        /// <param name="asmWch"></param>
        /// <param name="errorMessages"></param>
        internal void updateUsages(E3Assembly asmWch, List<string> errorMessages)
        {
            // У Андрея ключом между системами являлся винчиловский oidMaster, но электрики накопили в Е3 много компонентов созданных вручную (без интеграции), т.е. у них отсутствует oidMaster
            // Поэтому ПЫТАЮСЬ уйти от обязательного наличия у компонента Е3 атрибута oidMaster ! 
            // Ключом между системами буду использовать списки IDs у изделий, что располагаются на СБ чертеже и number у Доп. частей.

            // Создаем словарь для быстрого поиска по IDs
            Dictionary<string, E3PartUsage> wchUsagesDict = new Dictionary<string, E3PartUsage>();

            foreach (var wchUsage in asmWch._usages)
            {
                string key;
                if (wchUsage.IDs != null && wchUsage.IDs.Count > 0)
                {
                    // Создаем ключ из отсортированных IDs
                    var sortedIds = wchUsage.IDs.OrderBy(id => id).ToList();
                    key = string.Join(",", sortedIds);                    
                }
                else
                {
                    key = wchUsage.number; // У Доп.частей нет ID, т.к. они не расположены как самостоятельные Изделия на СБ чертеже, но обязательно имеют Number 
                }

                wchUsagesDict[key] = wchUsage; // добавляем в словарь
            }



            foreach (E3PartUsage currentE3PartUsage in _usages)
            {
                if (currentE3PartUsage.isForBOM == false)
                    continue;

                // Создаем ключ для поиска
                string currentKey;
                if (currentE3PartUsage.IDs != null && currentE3PartUsage.IDs.Count > 0)
                {
                    var sortedCurrentIds = currentE3PartUsage.IDs.OrderBy(id => id).ToList();
                    currentKey = string.Join(",", sortedCurrentIds);
                }
                else
                {
                    currentKey = currentE3PartUsage.number;
                }



                // Start. Выполняем проверки качества :
                if (wchUsagesDict.TryGetValue(currentKey, out E3PartUsage matchingUsageFromWch))
                {
                    String obj = !String.IsNullOrEmpty(matchingUsageFromWch.number) ? matchingUsageFromWch.number : matchingUsageFromWch.ATR_E3_ENTRY;
                    Part part = E3WGMForm.UtilsInstance.umens_e3project.Parts.Find(x => x.number == matchingUsageFromWch.number);

                    if (String.IsNullOrEmpty(matchingUsageFromWch.oidMaster))
                    {                        
                        if (!errorMessages.Contains($"Изделие {obj} {part.name} не найдено в Windchill"))
                            errorMessages.Add($"Изделие {obj} {part.name} не найдено в Windchill");

                        continue;
                    }

                    String wchState = matchingUsageFromWch.State;
                    if (wchState.Equals("Запрещено к применению") || wchState.Equals("Аннулировано") ||
                        wchState.Equals("Снято с производства")) // || wchState.Equals("Не в ограничительном перечне")
                    {
                        if (!errorMessages.Contains($"Изделие {obj} {part.name} {wchState} в Windchill"))
                            errorMessages.Add($"Изделие {obj} {part.name} {wchState} в Windchill");

                        //continue; поз. номер выводим !
                    }


                    if (matchingUsageFromWch.RS.Equals("Прочие изделия") || matchingUsageFromWch.RS.Equals("Стандартные изделия"))
                    {
                        String wchRestrict = matchingUsageFromWch.Restrict;
                        String prjRestrict = E3WGMForm.UtilsInstance.restrictProject;
                        //Part part = E3WGMForm.UtilsInstance.umens_e3project.Parts.Find(x => x.number == matchingUsageFromWch.number);

                        if (!prjRestrict.Equals("Без ограничений"))
                        {
                            if ( (prjRestrict.Equals("20") & !wchRestrict.Equals("20")) || (prjRestrict.Equals("30") & !wchRestrict.Equals("30")) )
                            {
                                String nameR;
                                E3WGMForm.UtilsInstance.restrictNames.TryGetValue(prjRestrict, out nameR);
                                if (!errorMessages.Contains($"{matchingUsageFromWch.number} {part.name} не входит в ограничительный перечень {nameR}"))
                                    errorMessages.Add($"{matchingUsageFromWch.number} {part.name} не входит в ограничительный перечень {nameR}");
                            }
                        }

                    }
                    // End. проверок качества


                    // Добавляем данные полученные в Windchill
                    currentE3PartUsage.oidMaster = matchingUsageFromWch.oidMaster; // если по заполненному в Е3 number нашли объект в Windchill, то перенесем вычисленный oidMaster в Е3, пусть будет. 
                    currentE3PartUsage.number = matchingUsageFromWch.number; // лишнее ?
                    currentE3PartUsage.unit = matchingUsageFromWch.unit;
                    //currentE3PartUsage.amount = matchingUsage.amount;
                    currentE3PartUsage.lineNumber = matchingUsageFromWch.lineNumber;
                    currentE3PartUsage.RS = matchingUsageFromWch.RS;

                    if (matchingUsageFromWch.isUsageE3Cable())
                    {
                        currentE3PartUsage.ATR_E3_WIRETYPE = matchingUsageFromWch.ATR_E3_WIRETYPE; // лишнее ?
                    }
                }
                else
                {
                    // такого по идее не может быть.
                }
            }
        }

        /// <summary>
        /// Для каждого Изделия входящего в нашу текущую Е3 сборку рассчитывает значение атрибута "Суммарная позиция".
        /// т.е. в этот атрибут заносятся позиции самого Изделия + позиции всех Дополнительных частей указанных в этом Изделии.
        /// Е3 "знает" об атрибуте "Суммарная позиция" и вынесет на СБ чертеж у Изделия одну общую выноску с несколькими полочками для позиций.
        /// </summary>
        private void updateSumPosInUsages(e3Job job)
        {
            e3Device dev = job.CreateDeviceObject();
            int devLineNumber = 0;
            String tempLineNumber;
            bool necesseryValue = false;
            List<int> listReplacementsAddPart; // список для взаимозамен в 150% BOM
            String currentDevRS;


            // 1. вычисляем данные для суммарной позиции
            foreach (E3PartUsage currentE3PartUsage in _usages)
            {
                //if (currentE3PartUsage.parentIDs.Count > 0 || currentE3PartUsage.netSegmentIds.Count > 0)
                //    continue; // имеем дело с E3PartUsage Дополнительных частей (доп.части, замены, cavity) или изделиями лежащими на сегменте, а нам нужны Изделия непосредственно расположенные на СБ чертеже от которых будет строиться выноска с номером позиции/й

                foreach (int itemId in currentE3PartUsage.IDs) // на СБ чертеже Изделие может встречаться несколько раз (разные Поз.обозначения) и могут быть с RS "Отсутствует"
                {
                    // сначала очищаем у всех изделий атрибут "суммарная позиция"
                    dev.SetId(itemId);
                    dev.SetAttributeValue(AttrsName.getAttrsName("lineNumber"), "");

                    if (currentE3PartUsage.parentIDs.Count > 0 || currentE3PartUsage.netSegmentIds.Count > 0)
                        continue; // имеем дело с E3PartUsage Дополнительных частей (доп.части, замены, cavity) или изделиями лежащими на сегменте, а нам нужны Изделия непосредственно расположенные на СБ чертеже от которых будет строиться выноска с номером позиции/й


                    tempLineNumber = "";
                    List<int> lineNumbers = new List<int>();
                                                            
                    currentDevRS = dev.GetAttributeValue(AttrsName.getAttrsName("atrBomRs"));

                    // Придумка Пеленга о 150% BOM. Проверяем их признак, т.е. в ЭСИ будет несколько записей с 0 количеством и ЕИ "По необходимости"
                    listReplacementsAddPart = new List<int>();
                    necesseryValue = dev.GetAttributeValue("IfNecessary") == "1" ? true : false;


                    //if (currentE3PartUsage.isForBOM == true)
                    if (!currentDevRS.Equals(BomRSValues.getBomRSValue((int)BomRSEnum.NO))) // т.е. у текущего изделия RS НЕ равен "Отсутствует"
                    {
                        devLineNumber = currentE3PartUsage.lineNumber;
                        lineNumbers.Add(devLineNumber); // Запомнили Позицию самого Изделия
                    }

                    foreach (E3PartUsage usageWithParent in _usages) // Ищем E3PartUsage Дополнительных частей у которых Parent-ом является изделие представленное itemId
                    {
                        if (usageWithParent.parentIDs.Contains(itemId))
                        {
                            if (necesseryValue == false) // у device обычные доп.части
                            {
                                if (!lineNumbers.Contains(usageWithParent.lineNumber))
                                {
                                    lineNumbers.Add(usageWithParent.lineNumber); // Добавили Позицию Дополнительной части 
                                }
                            }
                            else
                            {
                                listReplacementsAddPart.Add(usageWithParent.lineNumber); // накапливаем, чтобы потом сделать типа "15 или 16"
                            }
                        }
                    }


                    // 2. Заносим данные в атрибуты

                    if (necesseryValue == false)  // у device обычные доп.части
                    {
                        tempLineNumber = string.Join(" \r\n", lineNumbers.OrderBy(x => x)); // упорядочивает номера позиций по возрастанию и по правилу Е3, каждый номер должен располагаться строчкой ниже
                    }
                    else
                    {
                        tempLineNumber = string.Join(" или ", listReplacementsAddPart.OrderBy(x => x)); // Номера у самого Изделия нет, т.к. оно фэйковое и создано для 150% ЭСИ
                    }


                    // переходим к Изделию у которого в атрибут "Суммарная позиция" и занесем рассчитанное значение.
                    if (!String.IsNullOrEmpty(currentE3PartUsage.ATR_E3_WIRETYPE)) //TODO с таким не сталкивался, чтобы проводу назначали Доп.части
                    {
                        e3Pin pin = job.CreatePinObject();
                        pin.SetId(itemId);
                        pin.SetAttributeValue("BOMpos", devLineNumber.ToString()); //TODO это "Позиция спецификации". Надо ли вообще => добавить в attrsname.json
                        pin.SetAttributeValue(AttrsName.getAttrsName("lineNumber"), tempLineNumber);
                    }
                    else
                    {
                        //e3Device dev = job.CreateDeviceObject();
                        //dev.SetId(itemId);
                        dev.SetAttributeValue("BOMpos", devLineNumber.ToString()); //TODO это "Позиция спецификации". Надо ли вообще => добавить в attrsname.json
                        dev.SetAttributeValue(AttrsName.getAttrsName("lineNumber"), tempLineNumber); // В Е3 у "родительского" компонента будет выведена общая выноска с позициями 
                    }
                }
            }

        }

        /// <summary>
        /// Заполняем атрибут "Суммарная позиция" у сегмента.
        /// В интерфейсе Е3 чтобы его посмотреть нужно выделить сермент на СБ чертеже нажать ПКМ и выбрать "Параметры соединения"
        /// </summary>
        /// <param name="job"></param>
        private void updateSumPosInUsagesForNetSegments(e3Job job)
        {
            e3Device dev = job.CreateDeviceObject();
            e3Pin pin = job.CreatePinObject();
            e3NetSegment netSegment = job.CreateNetSegmentObject();
            int devLineNumber = 0;
            String tempLineNumber;
            List<int> listWireAndDevIds = null;

            List<int> lineNumbers = null;
            var dictionaryLineNumbersOnSegments = new Dictionary<int, List<int>>(); // мап Id сегментов - перечень Номеров позиций для каждого сегмента


            // 1. вычисляем данные для суммарной позиции сегмента
            Dictionary<int, List<int>> dictionaryIdDevsOnSegment = E3WGMForm.UtilsInstance.dictionaryIdDevsOnSegment;

            foreach (E3PartUsage currentE3PartUsage in _usages)
            {
                if ( currentE3PartUsage.netSegmentIds.Count == 0) // признак, что номер позиции из currentE3PartUsage должен выводиться на выноске от сегмента
                    continue;

                foreach (int segmentId in currentE3PartUsage.netSegmentIds) //
                {
                    if (segmentId == -1)
                        continue;

                    tempLineNumber = "";

                    // сначала очищаем у всех сегментов атрибут "суммарная позиция"
                    netSegment.SetId(segmentId);
                    netSegment.SetAttributeValue(AttrsName.getAttrsName("lineNumber"), "");


                    //если сегмент уже встречался в предыдущих currentE3PartUsage, то продолжаем дописывать ему номера позиций в lineNumbers
                    dictionaryLineNumbersOnSegments.TryGetValue(segmentId, out lineNumbers);
                    if(lineNumbers == null)
                    {
                        lineNumbers = new List<int>();
                        dictionaryLineNumbersOnSegments.Add(segmentId, lineNumbers);
                    }
                        


                    dictionaryIdDevsOnSegment.TryGetValue(segmentId, out listWireAndDevIds); // получили ранее наполненный перечень id объектов лежащих на сегменте

                    foreach (int itemId in currentE3PartUsage.IDs)
                    {
                        //E3WGMForm.UtilsInstance.app.PutInfo(0, $"itemId {itemId}", itemId);
                        if (!listWireAndDevIds.Contains(itemId)) // проверяем относится ли текущий itemId к текущему сегменту
                            continue;

                        if (currentE3PartUsage.isForBOM == true)
                        {
                            devLineNumber = currentE3PartUsage.lineNumber;
                            if (!lineNumbers.Contains(devLineNumber))
                                lineNumbers.Add(devLineNumber); // Запомнили Позицию самого Изделия
                        }

                        foreach (E3PartUsage usageWithParent in _usages) //Протестировать!!!  Ищем E3PartUsage Дополнительных частей у которых Parent-ом является изделие представленное itemId
                        {
                            if (usageWithParent.parentIDs.Contains(itemId)) // TODO проверить, может в parentIDs сидит другой itemId одного и того же изделия ?
                            {
                                if (!lineNumbers.Contains(usageWithParent.lineNumber))
                                    lineNumbers.Add(usageWithParent.lineNumber); // Добавили Позицию Дополнительной части 
                            }
                        }
                    }
                }
            }

            // 2. Заносим данные в атрибут сегмента
            
            foreach (KeyValuePair<int, List<int>> kvp in dictionaryLineNumbersOnSegments)
            {
                int segmentId = kvp.Key;
                lineNumbers = kvp.Value;

                netSegment.SetId(segmentId);
                //E3WGMForm.UtilsInstance.app.PutInfo(0, "segment", segmentId);

                tempLineNumber = string.Join(" \r\n", lineNumbers.OrderBy(x => x)); // упорядочивает номера позиций по возрастанию и по правилу Е3, каждый номер должен располагаться строчкой ниже
                netSegment.SetAttributeValue(AttrsName.getAttrsName("lineNumber"), tempLineNumber);
            }


        }
    }
}
