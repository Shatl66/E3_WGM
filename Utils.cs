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
using System.Net.NetworkInformation;
using System.Reflection.Emit;

namespace E3_WGM
{
    public class Utils
    {
        public e3Application app;
        public e3Job job;
        private e3StructureNode structureNode;
        private e3Tree tree;
        private e3Sheet sheet;
        private e3Bundle bundle;

        private List<string> folders = null; // для проверки от зацикливания, когда папка с именем "Х" может входить в себя же ниже по структуре папок.        
        private E3Assembly assemblyForPartsFromShemas = null;
        private int selectedRootNodeId = 0;

        public E3Assembly umens_e3project;
        public List<string> errorMessages = new List<string>();
        public List<string> numberPartsForE3ProjectDocument = new List<string>(); // с этими СЧ в Windchill будем связывать WTDocument проекта Е3


        public string tempPathForDoc { get; internal set; }
        public string restrictProject { get; internal set; } // Значение атрибута проекта «WCH Ограничительный перечень ПКИ проекта» 

        public Dictionary<string, string> typeDocuments = null; // устанавливается из E3WGMForms.cs

        public List<int> pinIDsTerminal = new List<int>(); // перечень ID пинов изделий у которых есть cavity объекты типа Заглушка
        public Dictionary<string, List<int>> pairesNumberTerminalCavityToPinIDs = new Dictionary<string, List<int>>(); //пары: обозначение Заглушки - ID пинов где назначена Заглушка

        public List<int> pinIDsSeal = new List<int>(); // перечень ID пинов изделий у которых есть cavity объекты типа Уплотнитель
        public Dictionary<string, List<int>> pairesNumberSealCavityToPinIDs = new Dictionary<string, List<int>>(); //пары: обозначение Уплотнителя - ID пинов где назначен Уплотнитель

        public Utils()
        {
            //ConnectToE3Series();
            /*
            app = CreateAppObject();
            job = CreateJobObject();
            structureNode = job.CreateStructureNodeObject(); // это пока прото пустышка заданного типа
            tree = job.CreateTreeObject(); // это пока прото пустышка заданного типа
            sheet = job.CreateSheetObject();
            */
        }

        public bool FindAllWTPartsFromSelectedFolder()
        {            
            bool found = false;
            int treeId = 0;
            String treeName;
            //errorMessages = new List<string>();

            treeId = job.GetActiveTreeID();
            if (treeId == 0)
            {
                MessageBox.Show("Выделите узел в дереве листов !", "",  MessageBoxButtons.OK, MessageBoxIcon.Information);

                DisconnectFromE3Series();
                return found;
            }
            else
            {
                tree.SetId(treeId); // инициализировали конкретный объект tree
                treeName = tree.GetName(); // может быть "Лист", "Изделия в проекте", "Чертеж жгута", "Цепи" и др.

                if (treeName != "Лист")
                {
                    MessageBox.Show("Сделайте активным узел в дереве листов !", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 
                    DisconnectFromE3Series();
                    return found;
                }

                dynamic structureNodeIds = null;
                int structureNodeCount = job.GetTreeSelectedStructureNodeIds(ref structureNodeIds);
                if (structureNodeCount > 1)
                {
                    MessageBox.Show("Сделайте активным только 1 узел в дереве листов !", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    DisconnectFromE3Series();
                    return found;
                }
                else if( structureNodeCount == 0) {
                    MessageBox.Show("Сделайте активным узел в дереве листов !", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    DisconnectFromE3Series();
                    return found;
                }
                else
                {
                    structureNode.SetId(structureNodeIds[1]);
                    
                    if( structureNode.IsLocked() == 1)
                    {
                        MessageBox.Show("Сделайте активным незаблокированный узел в дереве листов !", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        DisconnectFromE3Series();
                        return found;
                    }

                    string typeName = structureNode.GetTypeName();
                    if( typeName == ".DOCUMENT_TYPE")
                    {
                        MessageBox.Show("Сделайте активным узел в дереве листов !", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        DisconnectFromE3Series();
                        return found;
                    }

                    selectedRootNodeId = structureNodeIds[1];
                }
            }

            folders = new List<String>();
            FindSBRecursive( selectedRootNodeId, null);

            return true;
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

            if (structureNode.IsLocked() == 1) // если папка в проекте заблокирована, то не обрабатываем включенные в нее объекты.
                return;

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

                //app.PutInfo(0, $"Найден узел: {numberPart} {namePart}"); // тоже самое, что и "Найден WTPart: '" + nodeName + "' (ID: " + nodeId + ")"

                if (parentAssembly == null & typeName == "<Project>")
                {
                    // Настраиваем наш центральный объект umens_e3project на работу с данными именно от всего проекта Е3. т.к. именно сам проект выбран в дереве листов Е3

                    umens_e3project.number = job.GetName();
                    assembly = umens_e3project;
                }
                else if (parentAssembly == null)
                {
                    // Настраиваем наш центральный объект umens_e3project на работу с данными от выбранной папки в дереве листов Е3
                    umens_e3project.number = numberPart;
                    umens_e3project.name = namePart;
                    assembly = umens_e3project;

                    numberPartsForE3ProjectDocument.Add(numberPart);
                }
                else if(parentAssembly != null)
                {
                    assembly = new E3Assembly(numberPart, namePart);
                    parentAssembly.AddUsage(assembly);

                    numberPartsForE3ProjectDocument.Add(numberPart);
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

                //numberPartsForE3ProjectDocument.Add(parentAssembly.number);

                dynamic sAllSheetIds = null;
                int nAllSheet = structureNode.GetSheetIds(ref sAllSheetIds);
                //app.PutInfo(0, $"имеем дочерних схем - {nAllSheet}");
                for (int i = 1; i <= nAllSheet; i++)
                {
                    sheet.SetId(sAllSheetIds[i]); // sheet.IsEmbedded() == 1 это листы области чертежа внутри формбоард
                    managerFindParts( sheet, parentAssembly);                    
                }                
            }
        }



        //-----------------------------------------Start вычисления СЧ-ей из схем для BOM Windchill  ------------------------------------------------------------------------------------

        private void managerFindParts( e3Sheet sheet, E3Assembly assembly)
        {            
            assemblyForPartsFromShemas = assembly; // assemblyForPartsFromShemas и assembly указывают на один и тот же объект в памяти
            // глобальную assemblyForPartsFromShemas просто удобно использовать, к ней и будем в методах что ниже добавлять используя E3PartUsage все входящие изделия

            e3Symbol symbol = job.CreateSymbolObject();
            e3NetSegment netSegment = job.CreateNetSegmentObject();
            e3Pin pin = job.CreatePinObject();
            e3Device dev = job.CreateDeviceObject(); // Это Изделие и оно является экземпляром Компонента БД Е3 в текущем проекте Е3
            e3Component comp = job.CreateComponentObject(); // Компонент это объект в БД
            e3ConnectLine e3ConnectLine = job.CreateConnectLineObject();

            int symbolId = 0;
            int devId = 0;
            int netSegmentId = 0;
            int pinId = 0;

            // 1.
            dynamic sAllSimbolIds = null;
            int symbolCount = sheet.GetSymbolIds(ref sAllSimbolIds);            

            String typeSheet = sheet.GetAttributeValue(".DOCUMENT_TYPE");
            Console.WriteLine(" Вид документа " + typeSheet + ". Символов " + symbolCount);

            for (int i = 1; i <= symbolCount; i++)
            {
                try
                {
                    symbolId = sAllSimbolIds[i];
                    symbol.SetId(symbolId);                   

                    // 1.1 Получаем изделие
                    devId = dev.SetId( symbolId);

                    // 1.2 Изделие не найдено. Это когда за одним символом скрывается сразу несколько изделий !
                    if (devId == 0)
                    { 
                        dynamic symbolDevicePinIds = null;
                        int symbolDevicePinCount = symbol.GetDevicePinIds( ref symbolDevicePinIds); // поиск изделий через их пины

                        for (int k = 1; k <= symbolDevicePinCount; k++)
                        {
                            devId = dev.SetId(symbolDevicePinIds[k]);
                            ProcessDevice(dev);  //TODO не встретилось, поэтому не понятно что это такое, разобраться !
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

            // 2. Кабели ( 2 и более одиночных провода (жилы) в общей изоляции) не располагаются на СБ чертеже в виде собственного символа, поэтому в 1. их не нашло.
            // В этом блоке кода сначала будут найдены только жилы. Для каждой жилы будем определять это одиночный провод, или кабель и дальше их уже обрабатывть.

            var dictionarySymbolsOnSegments = new Dictionary<int, e3Component>();
            dynamic sAllSegmentIds = null;
            int segmentCount = sheet.GetNetSegmentIds(ref sAllSegmentIds);
            double tolerance;

            for (int i = 1; i <= segmentCount; i++)
            {
                try
                {
                    netSegmentId = sAllSegmentIds[i];
                    netSegment.SetId(netSegmentId);
                    
                    tolerance = double.TryParse(netSegment.GetAttributeValue("Tolerance"),
                                   System.Globalization.NumberStyles.Any,
                                   System.Globalization.CultureInfo.InvariantCulture,
                                   out double parsed) ? parsed / 1000 : 0.0;

                    // 2.1 Обрабатываем жилы проходящие через сегмент. Учитываем, что одна и таже жила может проходить через несколько сегментов
                    dynamic sAllCoreIds = null;
                    int coreCount = netSegment.GetCoreIds(ref sAllCoreIds);
                    
                    //app.PutInfo(0, $" Net Segment {netSegment.GetName()}. Core {coreCount}", netSegment.GetId());

                    // 2.1.2 Определяем к чему относится жила: к одиночному проводу, или к кабелю
                    for (int j = 1; j <= coreCount; j++)
                    {
                        pinId = sAllCoreIds[j];
                        pin.SetId(pinId);
                        //app.PutInfo(0, $"     Имя провода {pin.GetName()}", pin.GetId());
                        devId = dev.SetId( pinId); // где dev это Кабель к которому принадлежит жила.

                        if(dev.IsCable() == 1 & dev.IsOverbraid() == 0) // отбросили искуственный Кабель "Провода", где dev.IsOverbraid() тоже 1
                        {
                            ProcessCable(dev, pin, tolerance, netSegment.GetId()); // обрабатываем кабель
                        }
                        else {
                            ProcessCore(pin, tolerance); // обрабатываем одиночный провод
                        }

                    }

                    // 2.2 Обрабатываем изделия лежащие на сегменте для определения их производственной длинны
                    // использовать просто int symCount = netSegment.GetSymbolIds(ref sSymIds); к сожалению недостаточно. Код Данилы
                    E3Part part = null;
                    E3PartUsage usage = null;
                    double segmentManufLength = netSegment.GetManufacturingLength();
                    var symbolsOnLines = new List<int>();

                    dynamic connectLineIds = null;
                    int connectLineCount = netSegment.GetConnectLineIds(ref connectLineIds);
                    for (int l = 1; l <= connectLineCount; l++)
                    {
                        e3ConnectLine.SetId(connectLineIds[l]);

                        dynamic protectionSymbolIds = null;
                        int protectionSymbolCount = e3ConnectLine.GetProtectionSymbolIds(ref protectionSymbolIds);
                        for (int s = 1; s <= protectionSymbolCount; s++)
                        {
                            int protectionSymbolId = protectionSymbolIds[s];

                            if (symbolsOnLines.Contains(protectionSymbolId))
                                continue;

                            symbolsOnLines.Add(protectionSymbolId);

                            if (dictionarySymbolsOnSegments.ContainsKey(protectionSymbolId))
                            {
                                dictionarySymbolsOnSegments.TryGetValue(protectionSymbolId, out comp);
                                part = (E3Part)umens_e3project.Parts.Find(x => (x is E3Part) && (x as E3Part).ID == comp.GetId());
                                assemblyForPartsFromShemas.AddUsageLenght(dev, part, segmentManufLength);
                            }                                
                            else
                            {
                                devId = dev.SetId(protectionSymbolId);
                                if (dev.IsView() == 1) //TODO бывает ли тут View? Оставил код, пусть будет
                                {
                                    devId = dev.SetId(dev.GetOriginalId());
                                }

                                if (devId == 0)
                                    continue;

                                comp.SetId(devId);
                                part = (E3Part)umens_e3project.Parts.Find(x => (x is E3Part) && (x as E3Part).ID == comp.GetId());
                                assemblyForPartsFromShemas.AddUsageLenght(dev, part, segmentManufLength);
                                
                                dictionarySymbolsOnSegments.Add(protectionSymbolId, comp);
                            }

                            usage = assemblyForPartsFromShemas.Usages.Find(x => x.number == part.number);
                            usage.Tolerance += tolerance;
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorMessages.Add($"Ошибка при обработке NetSegment {netSegment.GetName()}, ID={netSegmentId}: {ex.Message}"); 
                }
            }


            symbol = null;
            netSegment = null;
            pin = null;
            dev = null;
            comp = null;
            e3ConnectLine = null;
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

            if (comp.GetId() == 0) // Удалили компонент из проекта и осталось только устройство и символ на чертеже ?
            {
               /// ProcessDeviceWithoutComponent(dev); //TODO не встретилось, поэтому не понятно что это такое, разобраться !
            }
            else
            {
                ProcessDeviceWithComponent(dev, comp); // т.е. компонент (comp) БД размещен (загружет в проект) на схеме, где уже считается Изделием ( dev)
            }
        }


        /// <summary>
        /// Обрабатывает устройство полученное из загруженного в проект компонента из библиотеки
        /// </summary>
        private void ProcessDeviceWithComponent(e3Device dev, e3Component comp)
        {
            if (dev.IsTerminal() == 1) //TODO не встретилось, поэтому не понятно что это такое, разобраться !
            {
                ProcessTerminalDevice(dev, comp);
                return;
            }

            ProcessStandardDevice(dev, comp);
        }


        /// <summary>
        /// Обрабатывает стандартное устройство
        /// </summary>
        private void ProcessStandardDevice(e3Device dev, e3Component comp)
        {
            try
            {
                E3Part part = null;
                Boolean incrAmount = true;

                // ищем сначала в кэше
                part = (E3Part)umens_e3project.Parts.Find(x => (x is E3Part) && (x as E3Part).ID == comp.GetId());

                if ( part == null)
                {
                    part = new E3Part(comp);

                    // ИЗДЕЛИЕ с RS "Отсутствует", не КОМПОНЕНТА ! Добавляем в Parts ! А в показе в WGM и передаче в Windchill будем исключать такие СЧ !!!
                    if (dev.GetAttributeValue(AttrsName.getAttrsName("atrBomRs")).Equals(BomRSValues.getBomRSValue((int)BomRSEnum.NO)))
                    {
                        Console.WriteLine(part.number + " RS Отсутствует !");
                        part.isForBOM = false;
                        incrAmount = false;
                    }

                    umens_e3project.Parts.Add(part);
                }
                else
                {
                    String currentDevRS = dev.GetAttributeValue(AttrsName.getAttrsName("atrBomRs"));
                    if( !currentDevRS.Equals(BomRSValues.getBomRSValue((int)BomRSEnum.NO))) // т.е. RS НЕ равен "Отсутствует"
                    {
                        part.isForBOM = true; // первое такое изделие встретившееся на СБ чертеже и попавшее в Parts, возможно было isForBOM = false. А теперь его всеже нужно показывать в ЭСИ
                        //E3PartUsage usage = umens_e3project.Usages.Find(x => x.number == part.number);
                    }
                    else
                    {
                        incrAmount = false;
                    }
                }
            
                double deltaAmount = 0.0;
                E3PartUsage usage = assemblyForPartsFromShemas.AddUsage(dev, part, out deltaAmount, incrAmount);
                usage.isForBOM = part.isForBOM; // если СЧ все же показывать (смотри коммент чуть выше), то это же значение нужно задать и тут

                AddReplacements(dev, usage);
                AddAdditionalParts(dev, assemblyForPartsFromShemas, deltaAmount);

                ProcessPins(dev);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке стандартного устройства '{dev.GetName()}': {ex.Message}");
            }
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

                    assembly.AddOrGetUsage(pin, e3cable); // у pin определяет его длинну и суммирует ее к такой же марке провода. ID каждого провода запоминает

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
                    usage = assembly.AddUsage(dev, part, out _, true);
                }
                else
                {
                    usage = assemblyForPartsFromShemas.AddUsage(dev, part, out _, true);
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
                    ProcessPins(dev);
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
        /// <para>Обрабатывает все пины изделия с целью нахождения всех жил подключенных к ним.</para>
        /// <para>Цель - обнаружить провода которых нет возможности вынести на СБ чертеж (заземление), но которые нужны в ЭСИ</para>
        /// При обработке сегментов цепи, как это делается в другом месте этого класса, эти провода не обнаруживаются.
        /// </summary>
        private void ProcessPins(e3Device devIzdelie)
        {
            e3Pin pinIzdelie = job.CreatePinObject();
            e3Pin core = job.CreatePinObject();
            e3Device dev = job.CreateDeviceObject(); //это Кабель или Одиночный провод к которому принадлежит жила.
            int coreId, pinIdIzdelie, devId;
            e3CavityPart cavity = job.CreateCavityPartObject();

            try
            {                
                dynamic sAllPinIds = null;
                int nAllPin = devIzdelie.GetAllPinIds(ref sAllPinIds);

                for (int j = 1; j <= nAllPin; j++)
                {
                    pinIdIzdelie = pinIzdelie.SetId(sAllPinIds[j]);

                    // сначала гарантированно обнуляем эти атрибуты у всех пинов изделий
                    pinIzdelie.SetAttributeValue("TipPosition", "0"); 
                    pinIzdelie.SetAttributeValue("SealerPosition", "0");

                    // Start наполнения списков pairesNumberTerminalCavityToPinIDs и pairesNumberSealCavityToPinIDs. 
                    // Эти списки будут позже использоваться для заполнения атрибутов пинов изделий номерами позиций Заглушек и Уплотнителей из ЭСИ

                    // Ситуация - 21 провод входит в вилку на 20 контактов.
                    // ВАЖНО. Тут cavity находятся от контактов, в отличие от кода в AddCavityOfCable(), где они находятся от жил. В итоге тут и там находится разное количество                    
                    // cavity объектов (в тестовом проекте разница на 1 шт.).
                    // Ответ Данилы - Правильно так как говорит заказчик)) уплотнители по количеству проводов... UniSP можно настроить и на провода, а не на выводы.
                    dynamic cavities = null;
                    int nTerminalCavityes = pinIzdelie.GetCavityPartIds(out cavities, 1); // ,1 - Ищем Заглушки назначенные пину изделия
                    
                    for (int k = 1; k <= nTerminalCavityes; k++)
                    {
                        int cavId = cavity.SetId(cavities[k]);
                        if (cavId != 0)
                        {
                            //Console.WriteLine("TERMINAL cavity" + "\t " + cavity.GetValue());
                            app.PutInfo(0, $" TERMINAL cavity {cavity.GetValue()}, пин {pinIzdelie.GetName()}", pinIdIzdelie);

                            if (String.IsNullOrEmpty(cavity.GetValue()))
                                continue;

                            if (pairesNumberTerminalCavityToPinIDs.TryGetValue(cavity.GetValue(), out var terminalIds))
                            {
                                if( !terminalIds.Contains( cavId))
                                    terminalIds.Add(cavId);
                            }
                            else
                            {
                                terminalIds = new List<int> { cavId };
                                pairesNumberTerminalCavityToPinIDs.Add(cavity.GetValue(), terminalIds);
                            }
                        }
                    }

                    cavities = null;
                    int nSealCavityes = pinIzdelie.GetCavityPartIds(out cavities, 2); // ,2 - Ищем Уплотнители назначенные пину изделия

                    for (int k = 1; k <= nSealCavityes; k++)
                    {
                        int cavId = cavity.SetId(cavities[k]);
                        if (cavId != 0)
                        {                            
                            app.PutInfo(0, $" SEAL cavity {cavity.GetValue()}, пин {pinIzdelie.GetName()}", pinIdIzdelie);

                            if (String.IsNullOrEmpty(cavity.GetValue()))
                                continue;

                            if (pairesNumberSealCavityToPinIDs.TryGetValue(cavity.GetValue(), out var sealIds))
                            {
                                if( !sealIds.Contains( cavId))
                                    sealIds.Add(cavId);
                            }
                            else
                            {
                                sealIds = new List<int> { cavId };
                                pairesNumberSealCavityToPinIDs.Add(cavity.GetValue(), sealIds);
                            }
                        }
                    }
                 // End наполнения списков


                    dynamic sAllCoreIds = null;
                    int coreCount = pinIzdelie.GetCoreIds(ref sAllCoreIds);

                    //app.PutInfo(0, $" Net Segment {netSegment.GetName()}. Core {coreCount}", netSegment.GetId());

                    // Определяем к чему относится жила: к одиночному проводу, или к кабелю
                    for (int k = 1; k <= coreCount; k++)
                    {
                        coreId = sAllCoreIds[k];
                        core.SetId(coreId);
                     app.PutInfo(0, $" pinIzdelie {pinIzdelie.GetName()}    Имя провода {core.GetName()}", core.GetId());
                        devId = dev.SetId(coreId); // где dev это Кабель к которому принадлежит жила.


                        if (dev.IsCable() == 1 & dev.IsOverbraid() == 0) // отбросили искуственный Кабель "Провода", где dev.IsOverbraid() тоже 1
                        {
                            ProcessCable(dev, core, 0, -1); // обрабатываем кабель
                        }
                        else if (dev.IsCable() == 1 & dev.IsOverbraid() == 1)
                        {
                            ProcessCore(core, 0); // обрабатываем одиночный провод
                        }
                    }
                }

                core = null;
                dev = null;
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке группы проводов '{dev.GetName()}': {ex.Message}");

                core = null;
                dev = null;
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
        /// Обрабатывает одиночный провод:
        /// <para>Если Раздел спецификации "Отсутствует", то пропускаем обработку этого провода</para>
        /// <para>Иначе. Добавляет провод к нашей ЭСИ и определяет его длину. Находит Cavity объекты связанные с проводом и добавляет их в нашу ЭСИ</para>
        /// </summary>
        private void ProcessCore(e3Pin pin, double tolerance)
        {
            if (pin.GetAttributeValue(AttrsName.getAttrsName("atrBomRs")).Equals(BomRSValues.getBomRSValue((int)BomRSEnum.NO)))
            {
                return; // cable.isForBOM = false; тут нам не нужен
            }

            try
            {
                dynamic wiregrouptype = null, wiretype = null;
                pin.GetWireType(ref wiregrouptype, ref wiretype);

                E3Cable cable = GetOrCreateCable(pin, wiregrouptype, wiretype); // при создании cable закомментировал занесение в него devId
                E3PartUsage usage = assemblyForPartsFromShemas.AddOrGetUsage(pin, cable);
                usage.Tolerance += tolerance;

                if (!cable.IDs.Contains(pin.GetId())) // проверяем обработали ли мы уже эту жилу
                {
                    // 1. Обрабатываем у жилы возможные объекты Cavity
                    AddCavityOfCable(pin, cable);
                    cable.IDs.Add(pin.GetId());

                    // 2.                     
                    double amount = pin.GetLength();
                    if (amount == 0)
                    {
                        string length = pin.GetAttributeValue(AttrsName.getAttrsName("cuttingLength")); //  в настройках Пеленга ->  .CORE_MANUFACTURING_LENGHT

                        if (length != null && length != "")
                        {
                            if (length.Contains(" ")) // TODO с таким не столкнулся
                                amount = Double.Parse(length.Split(' ')[0].Replace('.', ','));
                            else
                                amount = Double.Parse(length.Replace('.', ','));
                        }
                    }

                    amount = amount / 1000;

                    // проверяем входит ли провод в витую пару 
                    int bundleId = bundle.SetId( pin.GetId());
                    if( bundleId != 0 && bundle.IsTwisted() == 1)
                    {
                        amount = amount * 1.3;
                    }

                    usage.AddAmount(amount);
                    usage.addID(pin.GetId());
                }                
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке провода: {ex.Message}");
            }
        }

        private void ProcessCable(e3Device devCable, e3Pin pin, double tolerance, int netSegmentId)
        {
            try
            {
                E3Cable cable = GetOrCreateCable( devCable); // при создании cable закомментировал занесение в него devId
                E3PartUsage usage = assemblyForPartsFromShemas.AddOrGetUsage(pin, cable);                

                // 1. Обрабатываем у жилы возможные объекты Cavity
                if (!cable.IDs.Contains( pin.GetId())) // проверяем обработали ли мы уже эту жилу у нашего кабеля
                {                    
                    AddCavityOfCable(pin, cable);
                    cable.IDs.Add( pin.GetId()); // !!! в IDs кабеля заношу как Id самого кабеля, так и ID его жил, так и ID сегментов сети !!!
                }

                // 2. Проверяем обработали ли мы уже этот кабель. Длина всего кабеля определяется сразу по длинне 1 жилы
                if (!cable.IDs.Contains( devCable.GetId()))
                {
                    cable.IDs.Add( devCable.GetId());

                    double amount = pin.GetLength(); // длинну кабеля определяем по длинне жилы
                    if (amount == 0)
                    {
                        string length = pin.GetAttributeValue(AttrsName.getAttrsName("cuttingLength")); //  в настройках Пеленга ->  .CORE_MANUFACTURING_LENGHT

                        if (length != null && length != "")
                        {
                            if (length.Contains(" ")) // TODO с таким не столкнулся
                                amount = Double.Parse(length.Split(' ')[0].Replace('.', ','));
                            else
                                amount = Double.Parse(length.Replace('.', ','));
                        }
                    }

                    amount = amount / 1000;
                    usage.AddAmount(amount);

                    usage.addID( devCable.GetId());
                }

                // 3. Допуск на длину всего кабеля определяется суммированием допусков назначенных на каждый сегмент сети где проходит кабель.
                if (!cable.IDs.Contains( netSegmentId))
                {
                    usage.Tolerance += tolerance;
                    cable.IDs.Add( netSegmentId);
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при обработке провода: {ex.Message}");
            }
        }



        /// <summary>
        /// Получает или создает наш объект E3Cable для одиночного провода E3.
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
                errorMessages.Add($"Ошибка при создании провода '{wiregrouptype}/{wiretype}': {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Получает или создает наш объект E3Cable для кабеля Е3.
        /// </summary>
        private E3Cable GetOrCreateCable(e3Device dev)
        {
            String number = dev.GetAttributeValue("WCH_number");

            try
            {
                E3Cable cable = null;

                if (!umens_e3project.Parts.Exists(x => (x is E3Cable) &&
                                       ((x as E3Cable).number == number)))                                       
                {
                    cable = new E3Cable( dev);
                    umens_e3project.Parts.Add(cable);
                    // "Кабель добавлен"
                }
                else
                {
                    cable = (E3Cable)umens_e3project.Parts.Find(x => (x is E3Cable) &&
                                                     ((x as E3Cable).number == number));
                    // "Кабель был добавлен ранее"
                }

                return cable;
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Ошибка при создании кабеля '{number}': {ex.Message}");
                return null;
            }
        }


        /// <summary>
        /// Для каждого отдельного куска одного и того же провода находит его cavity (уплотнители, наконечники, заглушки, ...)
        /// и добавляет их в нашу ЭСИ
        /// </summary>
        /// <param name="pin"></param>
        /// <param name="cable"></param>
        private void AddCavityOfCable( e3Pin pin, E3Cable cable)
        {
            e3CavityPart cavity = job.CreateCavityPartObject();
            String numberCavityPart = "";

            dynamic cavities = null;
            int nAllCavityes = pin.GetCavityPartIds(out cavities, 0); // 2 - это all CavityParts of type Wire Seal, т.е. уплотнители 

            for (int k = 1; k <= nAllCavityes; k++)
            {
                int cavId = cavity.SetId(cavities[k]);
                if (cavId != 0)
                {
                    Console.WriteLine("cavity" + "\t " + cavity.GetValue());
                    if (string.IsNullOrEmpty(cavity.GetValue()))
                        continue;

                    AdditionalPart cavityPart = new AdditionalPart(); //TODO использую пока тип AdditionalPart() не по назначению
                    numberCavityPart = cavity.GetValue();
                    cavityPart.number = numberCavityPart; // этот объект с заполненным только number будем передавать в Windchill для его там поиска

                    // для реализации п.20 доп.требований (у контакта атрибут для позиции уплотнителя - "SealerPosition", Атрибут для позиции наконечника -"TipPosition"


                    try
                    {

                        if (!umens_e3project.Parts.Exists(x => x.number == numberCavityPart))
                        {
                            synchronizationAdditionalPartWithWindchill(cavityPart); // cavityPart дополняется значениями из Windchill
                            umens_e3project.Parts.Add(cavityPart);
                        }
                        else
                        {
                            cavityPart = (AdditionalPart)umens_e3project.Parts.Find(x => x.number == numberCavityPart);
                        }

                        E3PartUsage usage = null;

                        if (!cavityPart.ATR_BOM_RS.Equals("Материалы"))
                        {
                            usage = assemblyForPartsFromShemas.AddOrGetUsage(cavityPart, "ea");
                        }
                        else
                        {
                            // материала в качестве CAVITY наверное не бывает и сюда код не зайдет
                            usage = assemblyForPartsFromShemas.AddOrGetUsage(cavityPart, "m");
                        }

                        usage.amount++;
                        usage.addParentID( pin.GetId());
                    }
                    catch (Exception e)
                    {
                        if (!errorMessages.Contains($"Наконечник/Уплотнитель {e.Message}"))
                        {
                            errorMessages.Add($"Наконечник/Уплотнитель {e.Message}");
                            app.PutError(0, $"Наконечник/Уплотнитель {e.Message}", pin.GetId()); // чтобы в самом Е3 быстро найти провод у которого назначена эта cavity
                        }
                    }


                }
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
                        DisconnectFromE3Series();
                        Environment.Exit(0);
                    }
                }


                dynamic attributeIds = null;
                int nAllValues = dev.GetAttributeIds(ref attributeIds, "AdditionalReplacement");
                for (int i = 1; i <= nAllValues; i++) // считываем по очереди все значения многозначного атрибута "AdditionalReplacement"
                {
                    attribute.SetId(attributeIds[i]);
                    numberReplacement = attribute.GetValue();

                    if (usage.Replacements.Contains(numberReplacement))
                        continue;

                    if (!string.IsNullOrEmpty(numberReplacement))
                    {
                        try
                        {
                            AdditionalPart replacementPart = new AdditionalPart(); //TODO использую пока тип AdditionalPart() не по назначению
                            replacementPart.number = numberReplacement;
                            
                            synchronizationAdditionalPartWithWindchill(replacementPart); //TODO Я пока использую чисто для проверки наличия такой СЧ в Windchill 
                            usage.Replacements.Add(numberReplacement);

                            if (!umens_e3project.Parts.Exists(x => x.number == numberReplacement))
                            {
                                umens_e3project.Parts.Add(replacementPart);
                            }


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
                        DisconnectFromE3Series();
                        Environment.Exit(0);
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
                            continue;
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
                                Part testPart = umens_e3project.Parts.Find(x => x.number == numberAddPart); //защита от дурака, если назначили в качестве Доп.части реальное устройство Е3
                                if( testPart is AdditionalPart)
                                    additionalPart = (AdditionalPart)testPart;
                                else
                                {
                                    errorMessages.Add($"Как Доп.часть назначено устройство Е3 {numberAddPart}");
                                    app.PutError(0, $"Как Доп.часть назначено устройство Е3 {numberAddPart}", dev.GetId()); // чтобы в самом Е3 быстро найти Изделие у которого назначена эта Доп.часть
                                    continue;
                                }
                            }

                            E3PartUsage usage = null;

                            if (!additionalPart.ATR_BOM_RS.Equals("Материалы"))
                            {
                                usage = assembly.AddOrGetUsage(additionalPart, "ea");                                
                            }
                            else
                            {
                                usage = assembly.AddOrGetUsage(additionalPart, "m");
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







        internal void FindAllWTDocsFromSelectedFolder()
        {
            FindDocsRecursive(selectedRootNodeId, null, tempPathForDoc);
        }

        /// <summary>
        /// Рекурсивный метод поиска документов.
        /// Обходит вниз все дерево листов начиная от папки выделенной пользователем.
        /// </summary>
        /// <param name="nodeId"></param>
        /// <param name="parentAssembly"></param>
        private void FindDocsRecursive(int nodeId, E3Assembly parentAssembly, String tempPathForDoc)
        {
            String numberPart;
            String namePart;
            E3Assembly assembly = null;

            structureNode.SetId(nodeId);

            if (structureNode.IsLocked() == 1) // если папка в проекте заблокирована, то не обрабатываем включенные в нее объекты.
                return;

            string typeName = structureNode.GetTypeName();
            string internName = structureNode.GetInternalName();
            numberPart = structureNode.GetName();

            if (typeName == "<Assignment>" || typeName == "<Project>" || typeName == "SubProj") // У папки со схемами - .DOCUMENT_TYPE
            { // тут каждую папку дерева листов превращаем в Сборочную единицу (СЧ).

                bool hasAttribute = structureNode.HasAttribute("Naimen_izdel") == 1 ? true : false;
                namePart = hasAttribute ? structureNode.GetAttributeValue("Naimen_izdel") : "Наименование пока не известно";

                if (parentAssembly == null)
                {
                    assembly = umens_e3project; //  umens_e3project уже настроен (заданы number и name) на выделенную пользователем папку
                }
                else if (parentAssembly != null)
                {
                    assembly = (E3Assembly)umens_e3project.Parts.Find(x => x.number == numberPart); // эта сборка уже тоже создана и занесена в Parts
                }

                // Получаем дочерние узлы
                dynamic childNodeIds = null;
                int childCount = structureNode.GetStructureNodeIds(ref childNodeIds);

                for (int i = 1; i <= childCount; i++)
                {
                    int childNodeId = childNodeIds[i];
                    FindDocsRecursive(childNodeId, assembly, tempPathForDoc);
                }
            }
            else
            {   // дошли до папки Тип документа (.DOCUMENT_TYPE) в нее уже вложены схемы              
                string docType = "", docTypeSuff = "", docNumber = "", docName = "", docFormat = "";

                numberPart = parentAssembly.number;

                dynamic sAllSheetIds = null;
                int nAllSheet = structureNode.GetSheetIds(ref sAllSheetIds);

                for (int i = 1; i <= nAllSheet; i++)
                {
                    sheet.SetId(sAllSheetIds[i]); // sheet.IsEmbedded() == 1 это листы области чертежа внутри формбоард

                    //  Все значения должны быть заданы в свойствах листа !!! иначе выдавать ошибку !
                    if (sheet.IsEmbedded() == 0)
                    {
                        docType = sheet.GetAttributeValue(AttrsName.getAttrsName("docType"));
                        docType = docType.Replace("\r\n", "");
                        docTypeSuff = calcDocTypeSuffix(docType, sheet); //sheet.GetAttributeValue(AttrsName.getAttrsName("docTypeSuff")); 
                                                                                                      
                        docNumber = numberPart + docTypeSuff; // именно без пробела.                        

                        docFormat = sheet.GetAttributeValue(AttrsName.getAttrsName("docFormat"));
                        docFormat = docFormat.Replace("А", "A"); // на всякий случай русскую "А" меняем на английскую
                        docFormat = docFormat.Replace("х", "x"); // на всякий случай

                        docName = parentAssembly.name; // TODO всегда совпадает ? или читать атрибут листа "Наименование изделия" (он назначен надписи в форматке)
                    }


                    E3Documentation doc;
                    if (umens_e3project.Docs.Exists(x => x.number.Equals(docNumber)))
                    {
                        doc = (E3Documentation)umens_e3project.Docs.Find(x => x.number.Equals(docNumber));
                    }
                    else
                    {
                        doc = new E3Documentation(parentAssembly, docNumber, docName, job, docNumber, docType, tempPathForDoc); // job.GetPath()
                        umens_e3project.Docs.Add(doc);
                    }

                    doc.AddSheet( sheet, docFormat);
                }
            }
        }

        /// <summary>
        /// По типу документа определяю его суффикс.
        /// </summary>
        /// <param name="docType"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private string calcDocTypeSuffix( string docType, e3Sheet sheet)
        {
            String docSuff = "";
            typeDocuments.TryGetValue( docType, out docSuff);

            if (docSuff == null) // пусть будет пока и эта логика.
            {
                docSuff = sheet.GetAttributeValue(AttrsName.getAttrsName("docTypeSuff"));
            }

            return docSuff;
        }




        /////////////////// Технологические методы /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Для каждой найденной сборки идем в Windchill за недостающей информацией, в том числе - номерами позиций.
        /// </summary>
        public void SyncronizeE3ProjectDataWithWindchill()
        {
            foreach (Part part in umens_e3project.Parts)
            {
                if (part is E3Assembly)
                    syncE3Assembly((E3Assembly)part);
 ///               else if( part is E3Cable)
 ///                   syncE3Cable((E3Cable)part);
            }
        }

        /// <summary>
        /// Синхронизируются данные всех объектов входящих в E3Assembly
        /// </summary>
        /// <param name="assm"></param>
        public void syncE3Assembly( E3Assembly assm)
        {
            if (!E3WGMForm.wchHTTPClient.isAuthorization())
            {
                WindchillLoginForm wchLogin = new WindchillLoginForm(E3WGMForm.wchHTTPClient);
                wchLogin.ShowDialog();
                if (wchLogin.DialogResult.Equals(DialogResult.Cancel))
                {
                    DisconnectFromE3Series();
                    Environment.Exit(0);
                }
            }

            try
            {
                string jsonAssemblyFromWindchill = "";

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

                try
                {
                    jsonAssemblyFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonProject, "netmarkets/jsp/by/iba/e3/http/syncE3Assembly.jsp");
                }
                catch (Exception ex)
                {
                    UmensLogger.Log($"Синхронизация сборки {assm.number}. Сообщение Windchill: {ex.Message}");
                    return;
                }

                // Обратная замена при десериализации. Правильнее было бы прописать везде - [DataContract(Namespace = "E3WGM")]
                jsonAssemblyFromWindchill = jsonAssemblyFromWindchill.Replace("E3Assembly:#E3WGM", "E3Assembly:#E3_WGM");

                MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonAssemblyFromWindchill));
                DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Assembly), settings);
                E3Assembly assmWch = (E3Assembly)ser2.ReadObject(stream2);
                assm.merge( assmWch, errorMessages, job); // в подметоде наполняется список ошибок, формируется значение атрибута "суммарная позиция"
                fillPinAttrsByLineNumber( assmWch);
            }
            catch (Exception ex)
            {
                UmensLogger.Log($"Синхронизация сборки {assm.number}. Сообщение Windchill: {ex.Message}");
            }
        }

        /// <summary>
        /// Заполняет у контактов(выводов) изделий атрибуты "Позиция наконечника" и "Позиция уплотнителя" значением номера позиции наконечника или уплотнителя в спецификации
        /// </summary>
        /// <param name="assmWch"></param>
        private void fillPinAttrsByLineNumber(E3Assembly assmWch)
        {
            e3Pin pinIzdelie = job.CreatePinObject();
            String number;
            int lineNumber;

            foreach (var wchUsage in assmWch.Usages)
            {
                number = wchUsage.number;
                lineNumber = wchUsage.lineNumber;

                if (pairesNumberTerminalCavityToPinIDs.TryGetValue(number, out var terminalIds))
                {   
                    foreach( int pinID in terminalIds)
                    {
                        int id = pinIzdelie.SetId(pinID);
                        pinIzdelie.SetAttributeValue("TipPosition", lineNumber.ToString()); 
                    }
                    
                }

                if (pairesNumberSealCavityToPinIDs.TryGetValue(number, out var sealIds))
                {
                    foreach (int pinID in sealIds)
                    {
                        int id = pinIzdelie.SetId(pinID);
                        pinIzdelie.SetAttributeValue("SealerPosition", lineNumber.ToString());
                    }

                }

            }
        }

        public void syncE3Cable(E3Cable cable)
        {
            if (!E3WGMForm.wchHTTPClient.isAuthorization())
            {
                WindchillLoginForm wchLogin = new WindchillLoginForm(E3WGMForm.wchHTTPClient);
                wchLogin.ShowDialog();
                if (wchLogin.DialogResult.Equals(DialogResult.Cancel))
                {
                    DisconnectFromE3Series();
                    Environment.Exit(0);
                }
            }

            try
            {
                string jsonCableFromWindchill = "";

                MemoryStream stream = new MemoryStream();
                DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                settings.UseSimpleDictionaryFormat = true;
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(E3Cable), settings);
                ser.WriteObject(stream, cable);
                stream.Position = 0;
                StreamReader sr = new StreamReader(stream);
                string jsonE3Cable = sr.ReadToEnd();

                //В классах Windchill Андреем прописано пространство имен E3WGM. Я пока использую эти же классы, поэтому нужно сопоставлять E3WGM и мое E3_WGM
                jsonE3Cable = "{\"__type\":\"E3Cable:#E3WGM\"," + jsonE3Cable.Substring(1);

                try
                {
                    jsonCableFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonE3Cable, "netmarkets/jsp/by/iba/e3/http/findE3Cable.jsp");
                }
                catch (Exception ex)
                {
                    UmensLogger.Log($"Синхронизация провода/кабеля {cable.number}. Сообщение Windchill: {ex.Message}");
                    return;
                }

                // Обратная замена при десериализации. Правильнее было бы прописать везде - [DataContract(Namespace = "E3WGM")]
                jsonCableFromWindchill = jsonCableFromWindchill.Replace("E3Cable:#E3WGM", "E3Cable:#E3_WGM");

                MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonCableFromWindchill));
                DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Cable), settings);
                E3Cable cableWch = (E3Cable)ser2.ReadObject(stream2);
                cable.merge( cableWch, errorMessages, job);
            }
            catch (Exception ex)
            {
                UmensLogger.Log($"Синхронизация провода/кабеля {cable.number}. Сообщение Windchill: {ex.Message}");
            }
        }



        /// <summary>
        /// Для каждой найденного документа идем в Windchill за недостающей информацией, в том числе - есть такой документ в Windchill или нет
        /// </summary>
        public void SyncronizeE3DocsWithWindchill()
        {
            foreach (E3Documentation doc in umens_e3project.Docs)
            {
                SyncE3Document( doc);
            }
        }

        public void SyncE3Document(E3Documentation doc)
        {
            if (!E3WGMForm.wchHTTPClient.isAuthorization())
            {
                WindchillLoginForm wchLogin = new WindchillLoginForm(E3WGMForm.wchHTTPClient);
                wchLogin.ShowDialog();
                if (wchLogin.DialogResult.Equals(DialogResult.Cancel))
                {
                    DisconnectFromE3Series();
                    Environment.Exit(0);                   
                }
            }

            string jsonDocumentationFromWindchill = "";
            MemoryStream stream = new MemoryStream();
            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
            settings.UseSimpleDictionaryFormat = true;
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(E3Documentation), settings);
            ser.WriteObject(stream, doc);
            stream.Position = 0;
            StreamReader sr = new StreamReader(stream);
            string jsonDocumentation = sr.ReadToEnd();

            //В классах Windchill Андреем прописано пространство имен E3WGM. Я пока использую эти же классы, поэтому нужно сопоставлять E3WGM и мое E3_WGM
            jsonDocumentation = "{\"__type\":\"E3Documentation:#E3WGM\"," + jsonDocumentation.Substring(1);

            try
            {
                jsonDocumentationFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonDocumentation, "netmarkets/jsp/by/iba/e3/http/syncE3Documentation.jsp");
            }
            catch (Exception ex)
            {
                UmensLogger.Log($"Синхронизация документа {doc.number}. Сообщение Windchill: {ex.Message}");
                return;
            }

            // Обратная замена при десериализации. Правильнее было бы прописать везде - [DataContract(Namespace = "E3WGM")]
            jsonDocumentationFromWindchill = jsonDocumentationFromWindchill.Replace("E3Documentation:#E3WGM", "E3Documentation:#E3_WGM");

            MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonDocumentationFromWindchill));
            DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Documentation), settings);
            E3Documentation docWch = (E3Documentation)ser2.ReadObject(stream2);
            
            // было у Андрея (doc as E3Documentation).merge(job, docWch); 
            doc.updateDoc(docWch);
        }

        private void synchronizationAdditionalPartWithWindchill(AdditionalPart additionalPart)
        {
            string received = "";
            MemoryStream stream = new MemoryStream();
            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
            settings.UseSimpleDictionaryFormat = true;
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(AdditionalPart), settings);
            ser.WriteObject(stream, additionalPart);
            stream.Position = 0;
            StreamReader sr = new StreamReader(stream);

            string jsonAdditionalPart = sr.ReadToEnd();
            jsonAdditionalPart = "{\"__type\":\"AdditionalPart:#E3SetAdditionalPart\"," + jsonAdditionalPart.Substring(1);

            try
            {
                received = E3WGMForm.wchHTTPClient.getJSON(jsonAdditionalPart, "netmarkets/jsp/by/iba/e3/http/syncAdditionalPart.jsp");
            }
            catch (Exception ex)
            {
                UmensLogger.Log($"Синхронизация дополнительной части {additionalPart.number}. Сообщение Windchill: {ex.Message}");
                return;
            }

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
        public e3Application CreateAppObject()
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
        public e3Job CreateJobObject()
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


        public e3Job ConnectToE3Series()
        {
            app = CreateAppObject();
            job = CreateJobObject();
            structureNode = job.CreateStructureNodeObject(); // это пока прото пустышка заданного типа
            tree = job.CreateTreeObject(); // это пока прото пустышка заданного типа
            sheet = job.CreateSheetObject();
            bundle = job.CreateBundleObject();

            return job;
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
            sheet = null;
            bundle = null;
            //dev = null;
            //Cab = null;
            //pin = null;
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

        public E3Assembly getFilled_Umens_e3project()
        {
            return umens_e3project;
        }

        internal void getRestrictivProject()
        {
            restrictProject = job.GetAttributeValue("WCH_project_restrictive_list");

            if (String.IsNullOrEmpty( restrictProject))
            {
                MessageBox.Show("Заполните атрибут «WCH Ограничительный перечень проекта» допустимым значением из выпадающего списка в параметрах проекта.", "",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }
    }
}
