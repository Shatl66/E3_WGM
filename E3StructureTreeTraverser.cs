using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Runtime.Serialization.Json;
using System.Windows.Forms;
using Aga.Controls.Tree;
using e3;
using E3SetAdditionalPart;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using System.IO.Ports;

namespace E3_WGM
{
    internal class E3StructureTreeTraverser
    {
        private e3Application app;
        private e3Job job;
        private e3StructureNode structureNode;
        private e3Tree tree;
        //private E3Project umens_e3project;
        private E3Assembly umens_e3project;
        private List<String> folders = new List<String>(); // для проверки от зацикливания, когда папка с именем "Х" может входить в себя же ниже по структуре папок.
        public List<string> errorMessages = new List<string>();
        private E3Assembly assemblyForPartsFromShemas = null;


        public E3StructureTreeTraverser(E3Assembly umens_e3project)
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
            //errorMessages = new List<string>();

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

            FindSBRecursive( rootNodeId, null);
        }

        /// <summary>
        /// Рекурсивный метод поиска документов типа "Сборочный чертеж".
        /// Обходит вниз все дерево листов начиная от папки выделенной пользователем.
        /// </summary>
        /// <param name="nodeId"></param>
        /// <param name="parentAssembly"></param>
        private void FindSBRecursive(int nodeId, E3Assembly parentAssembly)
        {
            String numberPart;
            String namePart; 
            E3Assembly assembly = null;            

            structureNode.SetId(nodeId);

            string typeName = structureNode.GetTypeName();
            string internName = structureNode.GetInternalName();

            if (typeName == "<Assignment>" || typeName == "<Project>" || typeName == "SubProj") // У папки со схемами - .DOCUMENT_TYPE
            { // тут каждую папку дерева листов превращаем в Сборочную единицу (СЧ).
                numberPart = structureNode.GetName();

                if (folders.Contains(numberPart))
                {
                    app.PutError(1, $"Папка {numberPart} входит сама в себя !");
                    return;
                }
                else
                {
                    folders.Add(numberPart);
                }

                bool hasAttribute = structureNode.HasAttribute("Naimen_izdel") == 1 ? true : false;
                namePart = hasAttribute ? structureNode.GetAttributeValue("Naimen_izdel") : "Наименование пока не известно";

                app.PutInfo(0, $"Найден узел: {numberPart} {namePart}"); // тоже самое, что и "Найден WTPart: '" + nodeName + "' (ID: " + nodeId + ")"

                if (parentAssembly == null & typeName == "<Project>")
                {
                    // Настраиваем наш центральный объект umens_e3project на работу с данными именно от всего проекта Е3. т.к. именно сам проект выбран в дереве листов Е3

                    // {} Определить имя и обозначение для СЧ из проекта E3
                    assembly = umens_e3project;
                }
                else if (parentAssembly == null)
                {
                    // Настраиваем наш центральный объект umens_e3project на работу с данными от выбранной папки в дереве листов Е3
                    umens_e3project.number = numberPart;
                    umens_e3project.name = namePart;
                    assembly = umens_e3project;
                }
                else if(parentAssembly != null)
                {
                    assembly = new E3Assembly(numberPart, namePart);
                    parentAssembly.AddUsage(assembly);
                }

                umens_e3project.Parts.Add(assembly); // накапливаем в кэше все Part-ы по 1 разу

                // Получаем дочерние узлы
                dynamic childNodeIds = null;
                int childCount = structureNode.GetStructureNodeIds(ref childNodeIds);

                for (int i = 1; i <= childCount; i++)
                {
                    int childNodeId = childNodeIds[i];
                    FindSBRecursive(childNodeId, assembly);
                }
            }
            else if(internName == "Сборочный чертеж")
            {   // дошли до папки Тип документа (.DOCUMENT_TYPE) в нее уже вложены схемы. Теперь нужно достать все СЧ из схем
                e3Sheet sheet = job.CreateSheetObject();

                dynamic sAllSheetIds = null;
                int nAllSheet = structureNode.GetSheetIds(ref sAllSheetIds);
                app.PutInfo(0, $"имеем дочерних схем - {nAllSheet}");
                for (int i = 1; i <= nAllSheet; i++)
                {
                    sheet.SetId(sAllSheetIds[i]); // sheet.IsEmbedded() == 1 это листы области чертежа внутри формбоард
                    managerFindParts( sheet, parentAssembly);                    
                }                
            }
        }



        //-----------------------------------------Start вычисления СЧ-ей из схем ------------------------------------------------------------------------------------

        private void managerFindParts( e3Sheet sheet, E3Assembly assembly)
        {            
            assemblyForPartsFromShemas = assembly; // установили глобальную assemblyForPartsFromShemas, к ней и будем в методах что ниже добавлять используя E3PartUsage все входящие изделия
            e3Symbol symbol = job.CreateSymbolObject();
            e3NetSegment netSegment = job.CreateNetSegmentObject();
            e3Pin pin = job.CreatePinObject();
            e3Device dev = job.CreateDeviceObject(); // Это Изделие и оно является экземпляром Компонента БД Е3 в текущем проекте Е3
            int symbolId = 0;
            int devId = 0;
            int netSegmentId = 0;
            int pinId = 0;

            // 1.
            dynamic sAllSimbolIds = null;
            int symbolCount = sheet.GetSymbolIds(ref sAllSimbolIds);            

            String typeSheet = sheet.GetAttributeValue(".DOCUMENT_TYPE");
            Console.WriteLine(" Вид документа " + typeSheet + ". Символов " + symbolCount);
            app.PutInfo(0, $" Вид документа {typeSheet}. Символов {symbolCount}", sheet.GetId());

            for (int i = 1; i <= symbolCount; i++)
            {
                try
                {
                    symbolId = sAllSimbolIds[i];
                    symbol.SetId(symbolId);                   

                    // 1. Получаем изделие
                    devId = dev.SetId( symbolId);

                    // 2. Изделие не найдено. Это когда за одним символом скрывается сразу несколько изделий !
                    if (devId == 0)
                    { 
                        app.PutInfo(0, $"not found device - symbol {symbol.GetName()}", symbolId);
                        // код далее это вместо закомменченного ниже способа через - sheet.GetNetSegmentIds

                        dynamic symbolDevicePinIds = null;
                        int symbolDevicePinCount = symbol.GetDevicePinIds( ref symbolDevicePinIds); // поиск изделий через их пины

                        for (int k = 1; k <= symbolDevicePinCount; k++)
                        {
                            devId = dev.SetId(symbolDevicePinIds[k]);
                            ProcessDevice(dev); //TODO надо выполнять ProcessPin( pin); ??????
                        }
                    }
                    else
                    {
                        ProcessDevice(dev);
                    }                    
                }
                catch (Exception ex)
                {
                    errorMessages.Add($"Ошибка при обработке символа {symbol.GetName()}, ID={symbolId}: {ex.Message}"); //TODO может выводить про dev ?
                }
            }

            // 2.
            
            dynamic sAllSegmentIds = null;
            int segmentCount = sheet.GetNetSegmentIds(ref sAllSegmentIds);

            for (int i = 1; i <= segmentCount; i++)
            {
                try
                {
                    netSegmentId = sAllSegmentIds[i];
                    netSegment.SetId(netSegmentId);

                    dynamic sAllCoreIds = null;
                    int coreCount = netSegment.GetCoreIds(ref sAllCoreIds);

                    app.PutInfo(0, $" Net Segment {netSegment.GetName()}. Core {coreCount}", netSegment.GetId());

                    for (int j = 1; i <= coreCount; j++)
                    {
                        pinId = sAllCoreIds[j];
                        pin.SetId(pinId);

                        app.PutInfo(0, $"     CoreId {pinId}, Имя провода {pin.GetName()}", pin.GetId());

                        ProcessPin( pin);

                    }
                }
                catch (Exception ex)
                {
                    errorMessages.Add($"Ошибка при обработке NetSegment {netSegment.GetName()}, ID={netSegmentId}: {ex.Message}"); 
                }
            }
            
        }

        /// <summary>
        /// Обрабатывает одно изделие с целью определения из него 1 или нескольких СЧ (доп. части и замены) для Windchill
        /// </summary>
        private void ProcessDevice(e3Device dev)
        {
            // Уточняем изделие, только у оригинала будут найдены атрибуты
            if (dev.IsFormboard() == 1)
            {
                dev.SetId(dev.GetOriginalId());  // переходим к оригиналу, у него уже будут найдены атрибуты.
            }
            else if (dev.IsView() == 1)
            {
                dev.SetId(dev.GetOriginalId());
            }


            e3Component comp = job.CreateComponentObject(); // Компонент это объект в БД
            comp.SetId(dev.GetId());

            Console.WriteLine(dev.GetId() + "\t" + dev.GetName() + "\t" + dev.GetLocation() + "\t" + comp.GetId() + "\t " + comp.GetName());

            if (comp.GetId() == 0)
            {
                ProcessDeviceWithoutComponent(dev);
            }
            else
            {
                ProcessDeviceWithComponent(dev, comp);
            }
        }

        /// <summary>
        /// Обрабатывает устройство, отсутствующее в библиотеке компонентов. ? Провода
        /// </summary>
        private void ProcessDeviceWithoutComponent(e3Device dev)
        {
            app.PutInfo(0, $"Изделие {dev.GetName()}", dev.GetId());

            if (dev.IsCable() == 1) // По dev.GetId() E3 переходит к папке "Провода" в дереве изделий.
            {
                string assignment = dev.GetAssignment().Trim();

                if (string.IsNullOrEmpty(assignment))
                {
                    errorMessages.Add($"Для Компонента ДСЕ Жгута '{dev.GetName()}' не задано поле \"Устройство\".");
                    return;
                }

                //E3Assembly assembly = GetOrCreateE3Assembly(assignment);
                ProcessPinsForCableDevice(dev, assemblyForPartsFromShemas);

            }
            else if (dev.IsCable() == 0)
            {
                if (dev.IsWireGroup() == 1)
                {
                    ProcessWireGroup(dev);
                }
                else if (dev.IsTerminalBlock() == 1)
                {
                    ProcessTerminalBlock(dev);
                }
                else if (dev.IsTerminal() == 1)
                {
                    app.PutInfo(0, $"Доработать Наконечник !", dev.GetId());
                    errorMessages.Add("Доработать код !");
                }

            }
            else
            {
                errorMessages.Add($"Компонент {dev.GetName()} отсутствует в библиотеке");
            }
        }


        /// <summary>
        /// Обрабатывает группу проводов
        /// </summary>
        private void ProcessWireGroup(e3Device dev)
        {
            try
            {
                e3Pin pin = job.CreatePinObject();
                dynamic sAllPinIds = null;
                int nAllPin = dev.GetAllPinIds(ref sAllPinIds);

                for (int j = 1; j <= nAllPin; j++)
                {
                    pin.SetId(sAllPinIds[j]);
                    ProcessPin(pin);
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке группы проводов '{dev.GetName()}': {ex.Message}");
            }
        }

        /// <summary>
        /// Обрабатывает клеммную колодку
        /// </summary>
        private void ProcessTerminalBlock(e3Device dev)
        {
            try
            {
                e3Device localDev = job.CreateDeviceObject();
                e3Component localComp = job.CreateComponentObject();
                dynamic sAllLocalDevIds = null;
                int nAllLocalDev = dev.GetDeviceIds(ref sAllLocalDevIds);

                if (nAllLocalDev >= 1)
                {
                    localDev.SetId(sAllLocalDevIds[1]);
                    localComp.SetId(localDev.GetId());

                    if (localComp.GetId() != 0)
                    {
                        ProcessTerminalBlockComponent(localDev, localComp, dev, sAllLocalDevIds, nAllLocalDev);
                    }
                    else
                    {
                        errorMessages.Add($"Компонент IsTerminalBlock() '{localDev.GetName()}' отсутствует в библиотеке");
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке клеммной колодки '{dev.GetName()}': {ex.Message}");
            }
        }

        /// <summary>
        /// Обрабатывает компонент клеммной колодки
        /// </summary>
        private void ProcessTerminalBlockComponent(e3Device localDev, e3Component localComp, e3Device parentDev,
            dynamic sAllLocalDevIds, int nAllLocalDev)
        {
            string assignment = localDev.GetAssignment().Trim();
            E3Part part = new E3Part(localComp);

            if (!umens_e3project.Parts.Contains(part))
            {
                umens_e3project.Parts.Add(part);
                //LogInfo("Компонент добавлен");
            }
            else
            {
                part = (E3Part)umens_e3project.Parts.Find(x => x.oidMaster == part.oidMaster);
                //LogInfo("Компонент был добавлен ранее");
            }

            ProcessTerminalBlockUsage(parentDev, part, assignment, sAllLocalDevIds, nAllLocalDev);
        }

        /// <summary>
        /// Обрабатывает использование терминального блока
        /// </summary>
        private void ProcessTerminalBlockUsage(e3Device dev, E3Part part, string assignment,
            dynamic sAllLocalDevIds, int nAllLocalDev)
        {
            try
            {
                E3PartUsage usage = null;

                if (!string.IsNullOrEmpty(assignment) && !string.Equals(umens_e3project.number, assignment)) //TODO Проверить umens_e3project.number
                {
                    E3Assembly assembly = GetOrCreateE3Assembly(assignment); // не assemblyForPartsFromShemas ?
                    usage = assembly.AddUsage(dev, part, out _);
                }
                else
                {
                    usage = assemblyForPartsFromShemas.AddUsage(dev, part, out _);
                }

                // Добавляем все идентификаторы локальных устройств
                for (int k = 1; k <= nAllLocalDev; k++)
                {
                    usage?.addID(sAllLocalDevIds[k]);
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке использования терминального блока '{dev.GetName()}': {ex.Message}");
            }
        }

        /// <summary>
        /// Обрабатывает устройство с существующим компонентом в библиотеке
        /// </summary>
        private void ProcessDeviceWithComponent(e3Device dev, e3Component comp)
        {
            if (dev.IsTerminal() == 1)
            {
                ProcessTerminalDevice(dev, comp);
                return;
            }

            ProcessStandardDevice(dev, comp);
        }

        /// <summary>
        /// Обрабатывает терминальное устройство
        /// </summary>
        private void ProcessTerminalDevice(e3Device dev, e3Component comp)
        {
            try
            {
                E3Part part = new E3Part(comp);

                if (!umens_e3project.Parts.Contains(part))
                {
                    errorMessages.Add($"Предупреждение: Обнаружен компонент Terminal '{dev.GetName()}'");
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке терминального устройства '{dev.GetName()}': {ex.Message}");
            }
        }

        /// <summary>
        /// Обрабатывает стандартное устройство
        /// </summary>
        private void ProcessStandardDevice(e3Device dev, e3Component comp)
        {
            try
            {
                E3Part part = null;

                // ищем сначала в кэше
                part = (E3Part)umens_e3project.Parts.Find(x => (x is E3Part) && (x as E3Part).ID == comp.GetId());

                if ( part == null)
                {
                    part = new E3Part(comp);

                    // ИЗДЕЛИЕ без раздела спецификации, не компонента - у него заполнен RS !, добавляем ради нахождения Доп частей ! А в показе в WGM и передаче в Windchill будем исключать такие СЧ !!!
                    if (dev.GetAttributeValue(AttrsName.getAttrsName("atrBomRs")).Equals(BomRSValues.getBomRSValue((int)BomRSEnum.NO))) // if (part.ATR_BOM_RS.Equals(BomRSValues.getBomRSValue((int)BomRSEnum.NO)))
                    {
                        Console.WriteLine(part.number + " RS Отсутствует !");
                        part.isForBOM = false;
                    }

                    umens_e3project.Parts.Add(part);
                }
            
                double deltaAmount = 0.0;
                E3PartUsage usage = assemblyForPartsFromShemas.AddUsage(dev, part, out deltaAmount);
                AddReplacements(dev, usage);
                AddAdditionalParts(dev, assemblyForPartsFromShemas, deltaAmount);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке стандартного устройства '{dev.GetName()}': {ex.Message}");
            }
        }



        /// <summary>
        /// Обрабатывает в цикле все пины (жилы) кабеля, создает из них объекты E3Cable и добавляет их в сборку 
        /// </summary>
        private void ProcessPinsForCableDevice(e3Device dev, E3Assembly assembly)
        {
            try
            {
                e3Pin pin = job.CreatePinObject();
                E3Cable e3cable = null;

                dynamic sAllPinIds = null;
                int nAllPin = dev.GetAllPinIds(ref sAllPinIds);

                for (int j = 1; j <= nAllPin; j++)
                {
                    pin.SetId(sAllPinIds[j]);

                    dynamic wiregrouptype = null, wiretype = null;
                    pin.GetWireType(ref wiregrouptype, ref wiretype);

                    //Зачем он нужен ?            EnsureGeneralCableExists(wiregrouptype);

                    e3cable = GetOrCreateCable(pin, wiregrouptype, wiretype);

                    assembly.AddUsage(pin, e3cable); // у pin определяет его длинну и суммирует ее к такой же марке провода. ID каждого провода запоминает

                    ProcessCavityesOfPin(pin, assembly);
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке пинов кабеля '{dev.GetName()}': {ex.Message}");
            }
        }


        /// <summary>
        /// Обрабатывает наконечеики
        /// </summary>
        /// <param name="pin"></param>
        /// <param name="assembly"></param>
        private void ProcessCavityesOfPin(e3Pin pin, E3Assembly assembly)
        {            
            e3CavityPart cavity = job.CreateCavityPartObject();

            dynamic cavities = null;
            // ? pin.GetCavityPartsFromPinByCore();
            //int nAllCavityes = pin.GetEndCavityPartIds(0, ref cavities, 0); // последний параметр может быть &h01 - Only connector pin terminals 
            int nAllCavityes = pin.GetCavityPartIds( out cavities, 0); // 2 - это all CavityParts of type Wire Seal, т.е. уплотнители 

            for (int k = 1; k <= nAllCavityes; k++)
            {
                int cavId = cavity.SetId(cavities[k]);
                if( cavId != 0)
                    app.PutInfo(0, "Cavity", cavId);
            }
        }

        /// <summary>
        /// Гарантирует существование общего кабеля, если его еще нет, то создает
        /// </summary>
        private void EnsureGeneralCableExists(string wiregrouptype)
        {
            try
            {
                if (!umens_e3project.Parts.Exists(x => (x is E3GeneralCable) && ((x as E3GeneralCable).ATR_E3_ENTRY == wiregrouptype)))
                {
                    // Поиск в библиотеке компонентов
                    e3Component generalCableComp = job.CreateComponentObject();
                    dynamic sAllCompIds = null;
                    int nAllComp = job.GetAllComponentIds(ref sAllCompIds);

                    bool found = false;
                    for (int k = 0; k <= nAllComp; k++)
                    {
                        generalCableComp.SetId(sAllCompIds[k]);

                        if (generalCableComp.GetName().Equals(wiregrouptype))
                        {
                            E3GeneralCable generalCable = new E3GeneralCable(generalCableComp);
                            umens_e3project.Parts.Add(generalCable); // Разобраться зачем он нужен. Ведь это не СЧ в ЭСИ
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        E3GeneralCable generalCableTemp = new E3GeneralCable(wiregrouptype);
                        umens_e3project.Parts.Add(generalCableTemp);
                        //LogInfo($"Общий провод {generalCableTemp.ATR_E3_ENTRY} добавлен");
                    }

                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при создании общего кабеля '{wiregrouptype}': {ex.Message}");
            }
        }


        /// <summary>
        /// Получает или создает кабель
        /// </summary>
        private E3Cable GetOrCreateCable(e3Pin pin, string wiregrouptype, string wiretype)
        {
            try
            {
                E3Cable cable = null;

                if (!umens_e3project.Parts.Exists(x => (x is E3Cable) &&
                                       ((x as E3Cable).ATR_E3_ENTRY == wiregrouptype) &&
                                       ((x as E3Cable).ATR_E3_WIRETYPE == wiretype)))
                {
                    cable = new E3Cable(pin);
                    umens_e3project.Parts.Add(cable);
                    //LogInfo("Провод добавлен");
                }
                else
                {
                    cable = (E3Cable)umens_e3project.Parts.Find(x => (x is E3Cable) &&
                                                     ((x as E3Cable).ATR_E3_ENTRY == wiregrouptype) &&
                                                     ((x as E3Cable).ATR_E3_WIRETYPE == wiretype));
                    //LogInfo("Провод был добавлен ранее");
                }

                return cable;
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при создании кабеля '{wiregrouptype}/{wiretype}': {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Получает или создает сборку
        /// </summary>
        private E3Assembly GetOrCreateE3Assembly(string assemblyNumber)
        {
            E3Assembly assembly;

            if (!umens_e3project.Parts.Exists(x => x.number == assemblyNumber))
            {
                assembly = new E3Assembly(assemblyNumber, null);
                umens_e3project.Parts.Add(assembly);
                assemblyForPartsFromShemas.AddUsage(assembly);
                //LogInfo($"Сборка {assemblyNumber} добавлена");
            }
            else
            {
                assembly = (E3Assembly)umens_e3project.Parts.Find(x => x.number == assemblyNumber);
            }

            return assembly;
        }

        /// <summary>
        /// Обрабатывает пин (для группы проводов)
        /// </summary>
        private void ProcessPin(e3Pin pin)
        {
            try
            {
                dynamic wiregrouptype = null, wiretype = null;
                pin.GetWireType(ref wiregrouptype, ref wiretype);


                // TODO зачем он нужен EnsureGeneralCableExists(wiregrouptype);

                E3Cable cable = GetOrCreateCable(pin, wiregrouptype, wiretype);
                assemblyForPartsFromShemas.AddUsage(pin, cable);


                e3CavityPart cavity = job.CreateCavityPartObject();
                e3Component comp = job.CreateComponentObject();

                dynamic cavities = null;
                //int nAllCavityes = pin.GetCavityPartsFromPinByCore(pin.GetId(), out cavities, 0);
                //int nAllCavityes = pin.GetEndCavityPartIds(0, ref cavities, 0); // последний параметр может быть &h01 - Only connector pin terminals 
                int nAllCavityes = pin.GetCavityPartIds(out cavities, 0); // 2 - это all CavityParts of type Wire Seal, т.е. уплотнители 

                for (int k = 1; k <= nAllCavityes; k++)
                {
                    int cavId = cavity.SetId(cavities[k]);
                    if (cavId != 0)
                    {
                        int devId = comp.SetId( cavId);
                        if (devId == 0)
                            continue;

                        Console.WriteLine( comp.GetId() + "\t " + comp.GetName());


                    }
                }

            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке пина: {ex.Message}");
            }
        }

        /// <summary>
        /// Добавляем Подстановки. Используем E3PartUsage для их хранения
        /// </summary>
        /// <param name="dev"></param>
        /// <param name="usage"></param>
        private void AddReplacements(e3Device dev, E3PartUsage usage)
        {
            try
            {
                String numberReplacement = "";
                e3Attribute attribute = job.CreateAttributeObject();

                if (!E3WGMForm.wchHTTPClient.isAuthorization())
                {
                    WindchillLoginForm wchLogin = new WindchillLoginForm(E3WGMForm.wchHTTPClient);
                    wchLogin.ShowDialog();
                    if (wchLogin.DialogResult.Equals(DialogResult.Cancel))
                    {
                        return;
                    }
                }


                dynamic attributeIds = null;
                int nAllValues = dev.GetAttributeIds(ref attributeIds, "AdditionalReplacement");
                for (int i = 1; i <= nAllValues; i++) // считываем по очереди все значения многозначного атрибута "AdditionalReplacement"
                {
                    attribute.SetId(attributeIds[i]);
                    numberReplacement = attribute.GetValue();

                    if (usage.Replacements.Contains(numberReplacement))
                        continue; // символ этого же изделия в этой сборке уже мог встретиться на СБ чертеже

                    if (!string.IsNullOrEmpty(numberReplacement))
                    {
                        try
                        {
                            AdditionalPart replacementPart = new AdditionalPart(); //TODO использую пока тип AdditionalPart() не по назначению
                            replacementPart.number = numberReplacement;
                            synchronizationAdditionalPartWithWindchill(replacementPart); //TODO Я пока использую чисто для проверки наличия такой СЧ в Windchill 
                            usage.Replacements.Add(numberReplacement);
                        }
                        catch (Exception e)
                        {
                            if (!errorMessages.Contains($"Подстановка {e.Message}"))
                            {
                                errorMessages.Add($"Подстановка {e.Message}");
                                app.PutError(0, $"Подстановка {e.Message}", dev.GetId()); // чтобы в самом Е3 быстро найти Изделие у которого назначена эта подстановка
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при добавлении подстановок: {ex.Message}");
            }

        }


        /// <summary>
        /// Метод для добавления дополнительных частей
        /// </summary>
        private void AddAdditionalParts(e3Device dev, E3Assembly assembly, double deltaAmountHostDevice)
        {
            try
            {
                String numberAddPart = "";
                double amountAddPart = 1.0;
                e3Attribute attribute = job.CreateAttributeObject();

                if (!E3WGMForm.wchHTTPClient.isAuthorization())
                {
                    WindchillLoginForm wchLogin = new WindchillLoginForm(E3WGMForm.wchHTTPClient);
                    wchLogin.ShowDialog();
                    if (wchLogin.DialogResult.Equals(DialogResult.Cancel))
                    {
                        return;
                    }
                }


                dynamic attributeIds = null;
                int nAllValues = dev.GetAttributeIds(ref attributeIds, "AdditionalPart");                
                for (int i = 1; i <= nAllValues; i++) // считываем по очереди все значения многозначного атрибута "AdditionalPart"
                {
                    attribute.SetId( attributeIds[i]);
                    String valueAttr = attribute.GetValue(); // обозначение Доп.СЧ или обозначение Доп.СЧ;количество                    

                    if (!string.IsNullOrEmpty(valueAttr))
                    {
                        try
                        {
                            ParseValueAttr(valueAttr, out numberAddPart, out amountAddPart);
                        }
                        catch (Exception ex)
                        {
                            if (!errorMessages.Contains(dev.GetName() + " дополнительная часть " + ex.Message))
                                errorMessages.Add(dev.GetName() + " дополнительная часть " + ex.Message);
                        }

                        AdditionalPart additionalPart = new AdditionalPart();
                        additionalPart.number = numberAddPart; // этот объект с заполненным только number будем передавать в Windchill для его там поиска

                        try
                        {                             

                            if (!umens_e3project.Parts.Exists(x => x.number == numberAddPart))
                            {
                                synchronizationAdditionalPartWithWindchill(additionalPart); // additionalPart дополняется значениями из Windchill
                                umens_e3project.Parts.Add(additionalPart);
                            }
                            else
                            {
                                additionalPart = (AdditionalPart)umens_e3project.Parts.Find(x => x.number == numberAddPart);
                            }

                            E3PartUsage usage = null;;

                            if (!additionalPart.ATR_BOM_RS.Equals("Материалы"))
                            {
                                usage = assembly.AddUsage(additionalPart, "ea");                                
                            }
                            else
                            {
                                usage = assembly.AddUsage(additionalPart, "m");

                                /*
                                amount = 0;                                
                                if (additionalPartLength != null && additionalPartLength != "")
                                {
                                    if (additionalPartLength.Contains(" "))
                                    {
                                        amount = Double.Parse(additionalPartLength.Split(' ')[0].Replace('.', ','));
                                    }
                                    else
                                    {
                                        amount = Double.Parse(additionalPartLength.Replace('.', ','));
                                    }

                                    amount = amount / 1000;
                                }
                                */
                            }

                            usage.amount = usage.amount + deltaAmountHostDevice * amountAddPart;
                            usage.addParentID(dev.GetId());
                        }
                        catch (Exception e)
                        {
                            if (!errorMessages.Contains($"Доп.часть {e.Message}"))
                            {
                                errorMessages.Add($"Доп.часть {e.Message}");
                                app.PutError(0, $"Дополнительная часть {e.Message}", dev.GetId()); // чтобы в самом Е3 быстро найти Изделие у которого назначена эта Доп.часть
                            }
                        }

                    }
                }


                // Вместо MessageBox.Show:
                // errorMessages.Add($"ОШИБКА: {сообщение}");
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при добавлении дополнительных частей: {ex.Message}");
            }
        }



        /////////////////// Технологические методы /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Для каждой найденной сборки идем в Windchill за номерами позиций
        /// </summary>
        public void SyncronizeE3ProjectDataWithWindchill()
        {
            foreach (Part part in umens_e3project.Parts)
            {
                if (part is E3Assembly)
                    syncE3Assembly((E3Assembly)part);
            }
        }

        public void syncE3Assembly( E3Assembly assm)
        {
            MemoryStream stream = new MemoryStream();
            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
            settings.UseSimpleDictionaryFormat = true;
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(E3Assembly), settings);
            ser.WriteObject(stream, assm);
            stream.Position = 0;
            StreamReader sr = new StreamReader(stream);
            string jsonProject = sr.ReadToEnd();

            //В классах Windchill Андреем прописано пространство имен E3WGM. Я пока использую эти же классы, поэтому нужно сопоставлять E3WGM и мое E3_WGM
            jsonProject = "{\"__type\":\"E3Assembly:#E3WGM\"," + jsonProject.Substring(1);
            string jsonAssemblyFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonProject, "netmarkets/jsp/by/iba/e3/http/syncE3Assembly.jsp");
            // Обратная замена при десериализации. Правильнее было бы прописать везде - [DataContract(Namespace = "E3WGM")]
            jsonAssemblyFromWindchill = jsonAssemblyFromWindchill.Replace("E3Assembly:#E3WGM", "E3Assembly:#E3_WGM");


            MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonAssemblyFromWindchill));
            DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Assembly), settings);
            E3Assembly assmWch = (E3Assembly)ser2.ReadObject(stream2);
            assm.merge( assmWch, errorMessages, job);
        }


        private void synchronizationAdditionalPartWithWindchill(AdditionalPart additionalPart)
        {
            MemoryStream stream = new MemoryStream();
            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
            settings.UseSimpleDictionaryFormat = true;
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(AdditionalPart), settings);
            ser.WriteObject(stream, additionalPart);
            stream.Position = 0;
            StreamReader sr = new StreamReader(stream);

            string jsonAdditionalPart = sr.ReadToEnd();
            jsonAdditionalPart = "{\"__type\":\"AdditionalPart:#E3SetAdditionalPart\"," + jsonAdditionalPart.Substring(1);
            string received = E3WGMForm.wchHTTPClient.getJSON(jsonAdditionalPart, "netmarkets/jsp/by/iba/e3/http/syncAdditionalPart.jsp");
            MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(received));
            DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(AdditionalPart), settings);
            Part tempPart = (Part)ser2.ReadObject(stream2);

            if (String.IsNullOrEmpty(tempPart.oidMaster))
            {
                throw new Exception(additionalPart.number + " не найдена в Windchill");
            }

            additionalPart.merge(tempPart);
        }



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
            tree = null;
            structureNode = null;
            //dev = null;
            //Cab = null;
            //Pin = null;
            //Core = null;

            // принудительно удаляем пустые (null) ссылки
            GC.Collect();
        }


        /// <summary>
        /// <para>В атрибуте "Дополнительная честь" может содержаться значение по шаблону "number;количество"</para>
        /// Возвращает: Number доп.части, количество этой доп.части на единицу изделия где задана эта доп.часть.
        /// </summary>
        /// <param name="valueAttr"></param>
        /// <param name="numberAddPart"></param>
        /// <param name="amount"></param>
        /// <exception cref="FormatException"></exception>
        private void ParseValueAttr(string valueAttr, out string numberAddPart, out double amount)
        {
            // Проверка количества разделителей
            int separatorIndex = valueAttr.IndexOf(';');
            int lastSeparatorIndex = valueAttr.LastIndexOf(';');

            if (separatorIndex != lastSeparatorIndex)
                throw new FormatException("Допускается только один разделитель ';'");

            // Нет разделителя
            if (separatorIndex == -1)
            {
                numberAddPart = valueAttr;
                amount = 1.0;
                return;
            }

            // Проверка текста перед разделителем
            string text = valueAttr.Substring(0, separatorIndex);
            if (string.IsNullOrWhiteSpace(text))
                throw new FormatException("Текст перед ';' обязателен");

            // Нет числа после разделителя
            if (separatorIndex == valueAttr.Length - 1)
            {
                numberAddPart = text;
                amount = 1.0;
                return;
            }

            string numberStr = valueAttr.Substring(separatorIndex + 1);
            if (string.IsNullOrWhiteSpace(numberStr))
            {
                numberAddPart = text;
                amount = 1.0;
                return;
            }

            // Парсинг числа
            if (!double.TryParse(numberStr.Replace(',', '.'),
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double number))
            {
                throw new FormatException($"Неверный числовой формат: '{numberStr}'");
            }

            numberAddPart = text;
            amount = number;
            return;
        }


    }
}
