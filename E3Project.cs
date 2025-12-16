using e3;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using E3SetAdditionalPart;
using System.Xml.Linq;

namespace E3_WGM
{
    /// <summary>
    /// Содержит все данные (СЧ, WTDocuments, взаимосвязи) вычисленные от выбранной папки в дереве листов.
    /// <para>Эти данные отражаются в таблицах и передаются в Windchill</para>
    /// </summary>
    [DataContract]
    [KnownType(typeof(AdditionalPart))]
    [KnownType(typeof(E3Cable))]
    [KnownType(typeof(E3Part))]
    [KnownType(typeof(E3GeneralCable))]
    [KnownType(typeof(E3Assembly))]
    [KnownType(typeof(E3Project))]

    public class E3Project : E3Assembly
    {
        [DataMember]
        private List<Part> _parts = new List<Part>();
        internal List<Part> Parts
        {
            get { return _parts; }
            set { }
        }
        
        /*
        [DataMember]
        private List<Document> _docs = new List<Document>();
        internal List<Document> Docs
        {
            get { return _docs; }
            set { }
        }
        */
        
        private e3Job job;

        public e3Job getJob()
        {
            return job;
        }

        public E3Project(string projectNumber, string projectName) : base(projectNumber, projectName)
        {
            ATR_BOM_RS = BomRSValues.getBomRSValue((int)BomRSEnum.ASSEMBLY);
        }
        public E3Project(e3Job job) : base()
        {
            this.job = job;
            if (job.GetId() != 0)
            {
                int save = this.job.Save();
                getJobAttributes();
            }

            /*
            AddProjectDoc();
            if (!String.IsNullOrEmpty(oidMaster)
               && !String.IsNullOrEmpty(getE3ProjectDocument().oidMaster))
            {
                ReadDoc();
                ReadPart();
            }
            */
        }

        /*
        public void AddProjectDoc()
        {
            String E3prjOidMaster = "";
            String E3prjNumber = "";
            String E3prjName = "";

            bool differentNumbers = false;

            if (job.GetId() != 0)
            {
                E3prjOidMaster = job.GetAttributeValue("WCH_e3prj_id");
                E3prjNumber = job.GetAttributeValue("WCH_e3prj_number");
                E3prjName = job.GetAttributeValue("WCH_e3prj_name");
            }

            if (!String.IsNullOrWhiteSpace(E3prjNumber) && !String.IsNullOrWhiteSpace(this.number) && !E3prjNumber.StartsWith(this.number))
            {
                differentNumbers = true;
            }


            if (!differentNumbers)
            {
                E3prjNumber = this.number + "-E3";
                E3prjName = this.name;
            }
            E3ProjectDocument e3ProjectDocument = new E3ProjectDocument(E3prjNumber, E3prjName, job.GetPath(), job.GetName() + ".e3s");
            e3ProjectDocument.differentNumbers = differentNumbers;

            if (!String.IsNullOrEmpty(E3prjOidMaster))
            {
                e3ProjectDocument.oidMaster = E3prjOidMaster;
                Docs.Add(e3ProjectDocument);
                AddDescribe(e3ProjectDocument.oidMaster, E3PartDescribe.TypeOIDMaster);

            }
            else
            {
                Docs.Add(e3ProjectDocument);
                AddDescribe(e3ProjectDocument.number, E3PartDescribe.TypeNumber);
            }
        }
        

        internal void merge(E3Project project2)
        {
            E3ProjectDocument projectDoc = (E3ProjectDocument)Docs.Find(x => (x is E3ProjectDocument));
            E3ProjectDocument projectDoc2 = (E3ProjectDocument)project2.Docs.Find(x => (x is E3ProjectDocument));

            E3PartDescribe describe = null;
            if (String.IsNullOrEmpty(projectDoc.oidMaster))
            {
                describe = Describes.Find(x => x.value == projectDoc.number);
                describe.updateDescribe(projectDoc2);
                projectDoc.updateDoc(projectDoc2);
            }
            else
            {
                if (String.Equals(projectDoc.oidMaster, projectDoc2.oidMaster))
                {
                    describe = Describes.Find(x => x.value == projectDoc.oidMaster);
                    describe.updateDescribe(projectDoc2);
                    projectDoc.updateDoc(projectDoc2);
                }
                else
                {
                    throw new Exception("ОШИБКА: oidMaster документа проекта не совпадает с Windchill");
                }
            }
            updateJobAttribute(project2, projectDoc);


            updateLineNumber(this, project2);

            foreach (Part part in E3WGMForm.public_umens_e3project.Parts)
            {
                if (part is E3Assembly)
                {
                    (part as E3Assembly).Refresh();
                }
            }

            foreach (Part part2 in project2._parts)
            {
                if (part2 is E3Assembly)
                {
                    E3Assembly assembly2 = (E3Assembly)part2;

                    if (this._parts.Exists(x => x.oidMaster == assembly2.oidMaster))
                    {
                        E3Assembly assembly = (E3Assembly)this._parts.Find(x => x.oidMaster == assembly2.oidMaster);

                        updateLineNumber(assembly, assembly2);
                    }

                }

            }
        }
        */

        private void updateLineNumber(E3Assembly localAssebly, E3Assembly wchAssebly)
        {
            foreach (E3PartUsage usage2 in wchAssebly.Usages)
            {
                if (localAssebly.Usages.Exists(x => x.oidMaster == usage2.oidMaster))
                {
                    E3PartUsage usage = localAssebly.Usages.Find(x => x.oidMaster == usage2.oidMaster);
                    usage.lineNumber = usage2.lineNumber;
                }
            }

            foreach (E3PartUsage usage in localAssebly.Usages)
            {
                String tempLineNumber = "";

                foreach (int itemId in usage.IDs)
                {
                    tempLineNumber = "";
                    List<int> lineNumbers = new List<int>();
                    lineNumbers.Add(usage.lineNumber); // Запомнили Позицию основного ("родительского") компонента

                    foreach (E3PartUsage usageWithParent in localAssebly.Usages)
                    {
                        if (usageWithParent.parentIDs.Contains(itemId))
                        {
                            if (!lineNumbers.Contains(usageWithParent.lineNumber))
                            {
                                lineNumbers.Add(usageWithParent.lineNumber); // Запомнили Позицию компонента назначенного как дополнительный к "родительскому" 
                            }
                        }

                    }


                    foreach (int localLineNumber in lineNumbers.OrderBy(x => x)) // сортируем массив чисел по возрастанию
                    {
                        if (String.IsNullOrEmpty(tempLineNumber))
                        {
                            tempLineNumber = "" + localLineNumber;
                        }
                        else
                        {
                            tempLineNumber = tempLineNumber + " \r\n" + localLineNumber;
                        }
                    }
                    // !!! Цикл выше можно заменить одной строкой - tempLineNumber = string.Join(" \r\n", lineNumbers.OrderBy(x => x));

                    if (!String.IsNullOrEmpty(usage.ATR_E3_WIRETYPE))
                    {
                        e3Pin pin = job.CreatePinObject();
                        pin.SetId(itemId);
                        pin.SetAttributeValue(AttrsName.getAttrsName("lineNumber"), tempLineNumber);
                    }
                    else
                    {
                        e3Device dev = job.CreateDeviceObject();
                        dev.SetId(itemId);
                        dev.SetAttributeValue(AttrsName.getAttrsName("lineNumber"), tempLineNumber); // В Е3 у "родительского" компонента будет выведена общая выноска с позициями 
                    }
                }
            }
        }

        public void getJobAttributes()
        {
            this.oidMaster = job.GetAttributeValue("WCH_id");
            this.number = job.GetAttributeValue("WCH_number");
            this.name = job.GetAttributeValue("WCH_name");
            this.ATR_BOM_RS = job.GetAttributeValue(AttrsName.getAttrsName("atrBomRs"));
            if (this.ATR_BOM_RS == null || this.ATR_BOM_RS == "")
            {
                this.ATR_BOM_RS = BomRSValues.getBomRSValue((int)BomRSEnum.ASSEMBLY);//"Сборочные единицы";
            }
        }

/*
        public void updateJobAttribute(E3Project project2, E3ProjectDocument projectDoc)
        {
            this.wchcheckout = project2.wchcheckout;
            this.oid = project2.oid;
            updateJobAttribute(job, project2, projectDoc);
            getJobAttributes();

        }

        public static void updateJobAttribute(e3Job job, E3Project project2, E3ProjectDocument projectDoc)
        {
            job.SetAttributeValue("WCH_id", project2.oidMaster);
            job.SetAttributeValue("WCH_number", project2.number);
            job.SetAttributeValue("WCH_name", project2.name);
            job.SetAttributeValue(AttrsName.getAttrsName("atrBomRs"), project2.ATR_BOM_RS);

            job.SetAttributeValue("WCH_e3prj_id", projectDoc.oidMaster);
            job.SetAttributeValue("WCH_e3prj_number", projectDoc.number);
            job.SetAttributeValue("WCH_e3prj_name", projectDoc.name);
        }
*/


        internal void AddUsage(Part part)
        {
            E3PartUsage usage;
            if (!_usages.Exists(x => x.number == part.number))
            {
                usage = new E3PartUsage(part);
                usage.AddAmount();
                _usages.Add(usage);
                _usages = _usages.OrderBy(o => o.number).ToList();
            }
        }
        private void ReadPart()
        {
            e3Device dev = job.CreateDeviceObject(); // Это конкретный экземпляр Компонента в проекте Е3
            e3Component comp = job.CreateComponentObject(); // Компонент это объект в БД
            dynamic sAllDevIds = null;

            int nAllDev = job.GetAllDeviceIds(ref sAllDevIds);
            for (int i = 1; i <= nAllDev; i++)
            {
                dev.SetId(sAllDevIds[i]);

                if (dev.IsView() == 1)
                {
                    // Запрос идентификатора Оригинала и установка
                    dev.SetId(dev.GetOriginalId());
                }

                comp.SetId(dev.GetId());
                Console.WriteLine(dev.GetId() + "\t" + dev.GetName() + "\t" + dev.GetLocation() + "\t" + comp.GetId() + "\t " + comp.GetName());

                {
                    e3Pin pin = job.CreatePinObject();

                    dynamic sAllPinIds = null;
                    int nAllPin = dev.GetAllPinIds(ref sAllPinIds);
                    for (int j = 1; j <= nAllPin; j++)
                    {
                        pin.SetId(sAllPinIds[j]);
                        Console.WriteLine("=========================START Pin=======================================");
                        Console.WriteLine("0: " + dev.GetName());
                        Console.WriteLine("1: " + dev.GetAssignment());
                        Console.WriteLine("2: " + dev.GetLocation());
                        Console.WriteLine("8: " + pin.GetFitting());
                        Console.WriteLine("=========================FINISH Pin=======================================");
                    }
                }
                if (comp.GetId() == 0)
                {
                    if (dev.IsCable() == 1)
                    {
                        Console.WriteLine("0: " + dev.GetName());
                        Console.WriteLine("1: " + dev.GetAssignment());
                        Console.WriteLine("2: " + dev.GetLocation());
                        Console.WriteLine("3: " + dev.GetComponentName());
                        Console.WriteLine("WireGroup: " + dev.IsWireGroup());
                        Console.WriteLine("TerminalBlock: " + dev.IsTerminalBlock());

                        string assignment = dev.GetAssignment().Trim();
                        if (assignment != "")
                        {
                            E3Assembly assembly = null;
                            if (!_parts.Exists(x => x.number == assignment))
                            {
                                assembly = new E3Assembly(assignment, null);
                                _parts.Add(assembly);
                                AddUsage(assembly);
                                Console.WriteLine("ИНФО: Сборка " + assignment + " добавлена");
                            }
                            else
                            {
                                assembly = (E3Assembly)_parts.Find(x => x.number == assignment);
                            }

                            // addAdditionalParts(dev, assembly);

                            e3Pin pin = job.CreatePinObject();
                            dynamic sAllPinIds = null;
                            int nAllPin = dev.GetAllPinIds(ref sAllPinIds);
                            for (int j = 1; j <= nAllPin; j++)
                            {
                                pin.SetId(sAllPinIds[j]);
                                dynamic wiregrouptype = null, wiretype = null;
                                pin.GetWireType(ref wiregrouptype, ref wiretype);

                                if (!_parts.Exists(x => (x is E3GeneralCable) && ((x as E3GeneralCable).ATR_E3_ENTRY == wiregrouptype)))
                                {
                                    e3Component generalCableComp = job.CreateComponentObject();
                                    dynamic sAllCompIds = null;
                                    int nAllComp = job.GetAllComponentIds(ref sAllCompIds);

                                    for (int k = 0; k <= nAllComp; k++)
                                    {

                                        generalCableComp.SetId(sAllCompIds[k]);
                                        Console.WriteLine(generalCableComp.GetId() + " " + generalCableComp.GetName());
                                        if (generalCableComp.GetName().Equals(wiregrouptype))
                                        {
                                            Console.WriteLine("E3GeneralCable: " + generalCableComp.GetName() + " - " + generalCableComp.GetAttributeValue("Class"));

                                            E3GeneralCable generalCable = new E3GeneralCable(generalCableComp);
                                            _parts.Add(generalCable);
                                            Console.WriteLine("ИНФО: Общий провод " + generalCable.ATR_E3_ENTRY + " добавлен");
                                        }
                                    }
                                    if (!_parts.Exists(x => (x is E3GeneralCable) && ((x as E3GeneralCable).ATR_E3_ENTRY == wiregrouptype)))
                                    {
                                        E3GeneralCable generalCable = new E3GeneralCable(wiregrouptype);
                                        _parts.Add(generalCable);
                                        Console.WriteLine("ИНФО: Общий провод " + generalCable.ATR_E3_ENTRY + " добавлен");
                                    }
                                }

                                E3Cable cable = null;
                                if (!_parts.Exists(x => (x is E3Cable) && ((x as E3Cable).ATR_E3_ENTRY == wiregrouptype) && ((x as E3Cable).ATR_E3_WIRETYPE == wiretype)))
                                {
                                    cable = new E3Cable(pin);
                                    _parts.Add(cable);
                                    Console.WriteLine("ИНФО: Провод добавлен");
                                }
                                else
                                {
                                    cable = (E3Cable)_parts.Find(x => (x is E3Cable) && ((x as E3Cable).ATR_E3_ENTRY == wiregrouptype) && ((x as E3Cable).ATR_E3_WIRETYPE == wiretype));
                                    Console.WriteLine("ИНФО: Провод был добавлен ранее");
                                }
                                assembly.AddUsage(pin, cable);
                                ///========================================================================
                                Console.WriteLine("0: " + "Провод " + pin.GetName());
                                Console.WriteLine("2: " + "Провод " + wiregrouptype + " - " + wiretype);
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("WCH_id"));
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("WCH_number"));
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("WCH_name"));
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("WCH_barcode"));
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue(AttrsName.getAttrsName("atrBomRs")));
                                Console.WriteLine("2: " + "Провод " + pin.GetLength());
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("CuttingLength"));

                            }
                        }
                        else
                        {
                            MessageBox.Show("ОШИБКА: Для Компонента ДСЕ Жгута нe задано поле \"Устройство\".");
                        }


                        Console.WriteLine("ПРЕДУПРЕЖДЕНИЕ: Компонент ДСЕ Жгута");
                    }
                    else if (dev.IsCable() == 0)
                    {
                        Console.WriteLine("0: " + dev.GetName());
                        Console.WriteLine("1: " + dev.GetAssignment());
                        Console.WriteLine("2: " + dev.GetLocation());
                        Console.WriteLine("WireGroup: " + dev.IsWireGroup());
                        Console.WriteLine("TerminalBlock: " + dev.IsTerminalBlock());
                        if (dev.IsWireGroup() == 1)
                        {
                            e3Pin pin = job.CreatePinObject();
                            dynamic sAllPinIds = null;
                            int nAllPin = dev.GetAllPinIds(ref sAllPinIds);

                            for (int j = 1; j <= nAllPin; j++)
                            {
                                pin.SetId(sAllPinIds[j]);
                                dynamic wiregrouptype = null, wiretype = null;
                                pin.GetWireType(ref wiregrouptype, ref wiretype);
                                Console.WriteLine("0: " + "Провод " + pin.GetName());
                                Console.WriteLine("2: " + "Провод " + wiregrouptype + " - " + wiretype);
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("WCH_id"));
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("WCH_number"));
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("WCH_name"));
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("WCH_barcode"));
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue(AttrsName.getAttrsName("atrBomRs")));
                                Console.WriteLine("2: " + "Провод " + pin.GetLength());
                                Console.WriteLine("2: " + "Провод " + pin.GetAttributeValue("CuttingLength"));

                                if (!_parts.Exists(x => (x is E3GeneralCable) && ((x as E3GeneralCable).ATR_E3_ENTRY == wiregrouptype)))
                                {
                                    e3Component generalCableComp = job.CreateComponentObject();
                                    dynamic sAllCompIds = null;
                                    int nAllComp = job.GetAllComponentIds(ref sAllCompIds);
                                    for (int k = 0; k <= nAllComp; k++)
                                    {
                                        generalCableComp.SetId(sAllCompIds[k]);
                                        Console.WriteLine(generalCableComp.GetId() + " " + generalCableComp.GetName());
                                        if (generalCableComp.GetName().Equals(wiregrouptype))
                                        {
                                            Console.WriteLine("E3GeneralCable: " + generalCableComp.GetName() + " - " + generalCableComp.GetAttributeValue("Class"));

                                            E3GeneralCable generalCable = new E3GeneralCable(generalCableComp);
                                            _parts.Add(generalCable);
                                            Console.WriteLine("ИНФО: Общий провод " + generalCable.ATR_E3_ENTRY + " добавлен");
                                        }
                                    }


                                    if (!_parts.Exists(x => (x is E3GeneralCable) && ((x as E3GeneralCable).ATR_E3_ENTRY == wiregrouptype)))
                                    {
                                        E3GeneralCable generalCable = new E3GeneralCable(wiregrouptype);
                                        _parts.Add(generalCable);
                                        Console.WriteLine("ИНФО: Общий провод " + generalCable.ATR_E3_ENTRY + " добавлен");
                                    }
                                }

                                E3Cable cable = null;
                                if (!_parts.Exists(x => (x is E3Cable) && ((x as E3Cable).ATR_E3_ENTRY == wiregrouptype) && ((x as E3Cable).ATR_E3_WIRETYPE == wiretype)))
                                {
                                    cable = new E3Cable(pin);
                                    _parts.Add(cable);
                                    Console.WriteLine("ИНФО: Провод добавлен");
                                }
                                else
                                {
                                    cable = (E3Cable)_parts.Find(x => (x is E3Cable) && ((x as E3Cable).ATR_E3_ENTRY == wiregrouptype) && ((x as E3Cable).ATR_E3_WIRETYPE == wiretype));
                                    Console.WriteLine("ИНФО: Провод был добавлен ранее");
                                }
                                AddUsage(pin, cable);
                            }
                        }

                        if (dev.IsTerminalBlock() == 1)
                        {
                            e3Device localDev = job.CreateDeviceObject();
                            e3Component localComp = job.CreateComponentObject();
                            dynamic sAllLocalDevIds = null;
                            int nAllLocalDev = dev.GetDeviceIds(ref sAllLocalDevIds);
                            // for (int j = 1; j <= nAllLocalDev; j++)
                            if (nAllLocalDev >= 1)
                            {
                                localDev.SetId(sAllLocalDevIds[1]);
                                Console.WriteLine(localDev.GetName());
                                Console.WriteLine(localDev.GetComponentName());
                                localComp.SetId(localDev.GetId());
                                Console.WriteLine(localComp.GetName());
                                if (localComp.GetId() != 0)
                                {
                                    //==============================================================
                                    string assignment = localDev.GetAssignment().Trim();
                                    E3Part part = new E3Part(localComp);
                                    //==============================================================
                                    if (!_parts.Contains(part))
                                    {
                                        _parts.Add(part);
                                        Console.WriteLine("ИНФО: Компонент добавлен");
                                    }
                                    else
                                    {
                                        part = (E3Part)_parts.Find(x => x.oidMaster == part.oidMaster);
                                        Console.WriteLine("ИНФО: Компонент был добавлен ранее");
                                    }
                                    //==============================================================
                                    E3PartUsage usage = null;
                                    if (!String.IsNullOrEmpty(assignment) && !String.Equals(number, assignment))
                                    {
                                        E3Assembly assembly = null;
                                        if (!_parts.Exists(x => x.number == assignment))
                                        {
                                            assembly = new E3Assembly(assignment, null);
                                            _parts.Add(assembly);
                                            AddUsage(assembly);
                                            Console.WriteLine("ИНФО: Сборка " + assignment + " добавлена");
                                        }
                                        else
                                        {
                                            assembly = (E3Assembly)_parts.Find(x => x.number == assignment);
                                        }

                                        usage = assembly.AddUsage(dev, part);
                                    }
                                    else
                                    {
                                        usage = AddUsage(dev, part);
                                    }


                                    for (int k = 1; k <= nAllLocalDev; k++)
                                    {
                                        usage.addID(sAllLocalDevIds[k]);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("ОШИБКА: Компонент IsTerminalBlock() (" + localDev.GetName() + ") отсутсвует в библиотеке");
                                }
                            }

                        }
                    }
                    else
                    {
                        Console.WriteLine("ОШИБКА: Компонент отсутсвует в библиотеке");
                        MessageBox.Show("ОШИБКА: Компонент отсутсвует в библиотеке");
                    }

                    /* if (dev.IsWireGroup() == 1)
                     {
                         Console.WriteLine("ПРЕДУПРЕЖДЕНИЕ: Компонент <ПРОВОДА>");
                     }
                     else
                             if (dev.IsTerminalBlock() == 1)
                     {
                         Console.WriteLine("ПРЕДУПРЕЖДЕНИЕ: Компонент TerminalBlock");
                     }*/

                }
                else
                {
                    if (dev.IsTerminal() == 1)
                    {
                        E3Part part = new E3Part(comp);
                        if (!_parts.Contains(part))
                        {
                            Console.WriteLine("ПРЕДУПРЕЖДЕНИЕ: Компонент Terminal");
                            MessageBox.Show("ПРЕДУПРЕЖДЕНИЕ: Компонент Terminal");
                        }

                    }
                    else
                    /*if (dev.IsCable() == 1)
                {
                    Console.WriteLine("ПРЕДУПРЕЖДЕНИЕ: Компонент Покупной Жгут");
                    //MessageBox.Show("ПРЕДУПРЕЖДЕНИЕ: Компонент Покупной Жгут");
                }
                else*/
                    {
                        //==============================================================
                        string assignment = dev.GetAssignment().Trim();
                        E3Part part = new E3Part(comp);

                        if (part.ATR_BOM_RS.Equals(BomRSValues.getBomRSValue((int)BomRSEnum.NO)))
                        {
                            continue;
                        }
                        //==============================================================
                        if (!_parts.Contains(part))
                        {
                            _parts.Add(part);
                            Console.WriteLine("ИНФО: Компонент добавлен");
                        }
                        else
                        {
                            part = (E3Part)_parts.Find(x => (x is E3Part) && (x as E3Part).ID == part.ID);
                            Console.WriteLine("ИНФО: Компонент был добавлен ранее");
                        }
                        //==============================================================
                        if (!String.IsNullOrEmpty(assignment) && !String.Equals(number, assignment))
                        {
                            E3Assembly assembly = null;
                            if (!_parts.Exists(x => x.number == assignment))
                            {
                                assembly = new E3Assembly(assignment, null);
                                _parts.Add(assembly);
                                AddUsage(assembly);
                                Console.WriteLine("ИНФО: Сборка " + assignment + " добавлена");
                            }
                            else
                            {
                                assembly = (E3Assembly)_parts.Find(x => x.number == assignment);
                            }

                            assembly.AddUsage(dev, part);

//                            addAdditionalParts(dev, assembly);
                        }
                        else
                        {
                            AddUsage(dev, part);

//                          addAdditionalParts(dev, this);
                        }
                    }
                }
            }
        }


/*
        /// <summary>
        /// Взял с GitHub с ветки OKBTSP
        /// </summary>
        /// <param name="dev"></param>
        /// <param name="assembly"></param>
        private void addAdditionalParts(e3Device dev, E3Assembly assembly)
        {
            if (!E3WGMForm.wchHTTPClient.isAuthorization())
            {
                //return;
                WindchillLoginForm wchLogin = new WindchillLoginForm(E3WGMForm.wchHTTPClient);
                wchLogin.ShowDialog();
                if (wchLogin.DialogResult.Equals(DialogResult.Cancel))
                {
                    return;
                }
            }

            for (int ii = 1; ii <= AdditionalPart.additionPartsMaxCount; ii++)
            {
                String valueId = dev.GetAttributeValue("WCH_AdditionalPart0" + ii + "_id");
                if (valueId != null && valueId != "")
                {

                    AdditionalPart additionalPart = new AdditionalPart();
                    string additionalPartLength = "0";
                    string additionalLineNumber = "0";

                    for (int jj = 0; jj < AdditionalPart.additionPartsSuffix.Length; jj++)
                    {
                        String value = dev.GetAttributeValue("WCH_AdditionalPart0" + ii + "_" + AdditionalPart.additionPartsSuffix[jj]);
                        switch (AdditionalPart.additionPartsSuffix[jj])
                        {
                            case "id":
                                additionalPart.oidMaster = value;
                                break;
                            case "number":
                                additionalPart.number = value;
                                break;
                            case "name":
                                additionalPart.name = value;
                                break;
                            case "length":
                                additionalPartLength = value;
                                break;
                            case "lineNumber":
                                additionalLineNumber = value;
                                break;
                        }
                    }


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
                        MessageBox.Show("ОШИБКА: В Windchill отсутствует СЧ " + additionalPart.number);
                        return;
                    }

                    additionalPart.merge(tempPart);

                    if (!_parts.Contains(additionalPart))
                    {
                        _parts.Add(additionalPart);
                        Console.WriteLine("ИНФО: Компонент добавлен - " + additionalPart.name);
                    }
                    else
                    {
                        additionalPart = (AdditionalPart)_parts.Find(x => x.oidMaster == additionalPart.oidMaster);
                        Console.WriteLine("ИНФО: Компонент был добавлен ранее - " + additionalPart.name);
                    }


                    E3PartUsage usage = null;
                    double amount = 0;

                    if (!additionalPart.ATR_BOM_RS.Equals("Материалы"))
                    {
                        usage = assembly.AddUsage(additionalPart, "ea");
                        amount = 1; //TODO ввод количества если у одного Е3 компонента сразу несколько одинаковой AdditionalPart (завести такой атрибут ?)
                    }
                    else
                    {
                        usage = assembly.AddUsage(additionalPart, "m");

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
                    }

                    usage.AddAmount(amount); // наращивает количество если Компонент или AdditionalPart был добавлен ранее
                    usage.addParentID(dev.GetId());
                }

            }
        }
*/


        /*
        private void addAdditionalParts(e3Device dev, E3Assembly assembly)
        {
            if (!E3WGMForm.wchHTTPClient.isAuthorization())
            {
                //return;
                WindchillLoginForm wchLogin = new WindchillLoginForm(E3WGMForm.wchHTTPClient);
                wchLogin.ShowDialog();
                if (wchLogin.DialogResult.Equals(DialogResult.Cancel))
                {
                    return;
                }
            }
            String test1 = dev.GetAttributeValue(AttrsName.getAttrsName("lineNumber"));
            //String test2 = dev.GetAttributeValue(AttrsName.getAttrsName("additionalPart"));
            dynamic sAllAttributeIds = null;
            int nAllAttribute = dev.GetAttributeIds(ref sAllAttributeIds, AttrsName.getAttrsName("additionalPart"));
            e3Attribute attrobj = job.CreateAttributeObject();
            for (int i = 1; i <= nAllAttribute; i++)
            {
                attrobj.SetId(sAllAttributeIds[i]);
                String entryAdditionalPart = attrobj.GetValue();

                if (String.IsNullOrEmpty(entryAdditionalPart))
                {
                    attrobj.Delete();
                    continue;
                }

                addAdditionalParts(entryAdditionalPart, dev, assembly);

            }

            //Fitting (Наконечники)
            dynamic sAllPinIds = null;
            int nAllPin = dev.GetAllPinIds(ref sAllPinIds); //Находит все пины устройства в E3.series
            e3Pin pin = job.CreatePinObject();
            for (int i = 1; i <= nAllPin; i++)
            {
                pin.SetId(sAllPinIds[i]);
                string entryAdditionalPart = pin.GetFitting();

                if (String.IsNullOrEmpty(entryAdditionalPart)) // Для каждого пина проверяет, есть ли у него наконечник(Fitting).
                {
                    attrobj.Delete(); //  выглядит подозрительно, потому что attrobj не связан напрямую с пинами. Возможно, это ошибка, и вместо этого должен удаляться сам пин или его атрибут.
                    continue;
                }

                addAdditionalParts(entryAdditionalPart, dev, assembly);
            }
        }

        private void addAdditionalParts(String entryAdditionalPart, e3Device dev, E3Assembly assembly)
        {
            E3Part part = null;
            if (!_parts.Exists(x => (x is E3Part) && (x as E3Part).ATR_E3_ENTRY == entryAdditionalPart))
            {
                part = new E3Part();
                part.ATR_E3_ENTRY = entryAdditionalPart;

                MemoryStream stream = new MemoryStream();
                DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                settings.UseSimpleDictionaryFormat = true;
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(E3Part), settings);
                ser.WriteObject(stream, part);
                stream.Position = 0;
                StreamReader sr = new StreamReader(stream);
                string jsonE3Part = sr.ReadToEnd();
                jsonE3Part = "{\"__type\":\"E3Part:#E3WGM\"," + jsonE3Part.Substring(1);
                string jsonE3PartFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonE3Part, "netmarkets/jsp/by/iba/e3/http/findE3PartByEntry.jsp");
                MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonE3PartFromWindchill));
                DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Part), settings);
                E3Part tempE3Part = (E3Part)ser2.ReadObject(stream2);
                if (!String.IsNullOrEmpty(tempE3Part.oidMaster))
                {
                    part.merge(tempE3Part);
                    _parts.Add(part);

                    Console.WriteLine("ИНФО: Компонент добавлен");
                }
                else
                {
                    MessageBox.Show("ОШИБКА: В Windchill отсутствует СЧ в которой ATR_E3_ENTRY = " + entryAdditionalPart);
                    return;
                }

            }
            else
            {
                part = (E3Part)_parts.Find(x => (x is E3Part) && (x as E3Part).ATR_E3_ENTRY == entryAdditionalPart);
                Console.WriteLine("ИНФО: Компонент был добавлен ранее");
            }

            assembly.AddUsage(part, dev.GetId());
        }
        */


/*
        private void ReadDoc()
        {
            Dictionary<string, string> allDocType = new Dictionary<string, string>();
            if (File.Exists("doctype.json"))
            {
                String jsonDocType = "";
                using (StreamReader streamReader = new StreamReader("doctype.json"))
                {
                    jsonDocType = streamReader.ReadToEnd();
                }

                using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonDocType)))
                {
                    DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                    settings.UseSimpleDictionaryFormat = true;
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Dictionary<string, string>), settings);
                    allDocType = ser.ReadObject(stream) as Dictionary<string, string>;
                }
            }

            List<string> allDocFormat = new List<string>();
            if (File.Exists("docformat.json"))
            {
                String jsonDocFormat = "";
                using (StreamReader streamReader = new StreamReader("docformat.json"))
                {
                    jsonDocFormat = streamReader.ReadToEnd();
                }

                using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonDocFormat)))
                {
                    DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                    settings.UseSimpleDictionaryFormat = true;
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(List<string>), settings);
                    allDocFormat = ser.ReadObject(stream) as List<string>;
                }
            }

            e3Sheet sheet = job.CreateSheetObject();

            dynamic sSheetIds = null;
            int nSheetIds = job.GetSheetIds(ref sSheetIds);
            String message = "" + nSheetIds;
            Debug.WriteLine(message);
            for (int i = 1; i <= nSheetIds; i++)
            {
                sheet.SetId(sSheetIds[i]);

                string wchID = sheet.GetAttributeValue("WCH_id");

                string partNumber = sheet.GetAssignment();
                if (String.IsNullOrEmpty(partNumber))
                {
                    partNumber = sheet.GetAttributeValue(AttrsName.getAttrsName("partNumber"));
                }
                string docType = sheet.GetAttributeValue(AttrsName.getAttrsName("docType"));
                docType = docType.Replace("\r\n", "");
                string docFormat = sheet.GetAttributeValue(AttrsName.getAttrsName("docFormat"));
                if (String.IsNullOrEmpty(docFormat))
                {
                    docFormat.Replace("А", "A");
                    docFormat.Replace("х", "x");
                }

                string sheetName = sheet.GetName();
                int sheetN = -1;

                string docName = "";
                dynamic sTextIds = null;
                int nTextIds = sheet.GetTextIds(ref sTextIds);
                e3Text text = job.CreateTextObject();
                for (int j = 1; j <= nTextIds; j++)
                {
                    text.SetId(sTextIds[j]);
                    switch (text.GetType())
                    {
                        case 507:
                            docName = text.GetText();
                            break;
                    }
                }



                if (Int32.TryParse(sheetName, out sheetN)
                    && !String.IsNullOrEmpty(partNumber)
                    && !String.IsNullOrEmpty(docType)
                    && !String.IsNullOrEmpty(docFormat)
                    && allDocType.ContainsKey(docType)
                    && allDocFormat.Contains(docFormat))
                {
                    E3Assembly assembly;
                    E3ProjectDocument e3PrjDoc = this.getE3ProjectDocument();
                    if (partNumber.Equals(number) || (e3PrjDoc.differentNumbers && e3PrjDoc.number.StartsWith(partNumber)))
                    {
                        assembly = this;
                    }
                    else
                    {
                        if (_parts.Exists(x => x.number.Equals(partNumber)))
                        {
                            assembly = (E3Assembly)_parts.Find(x => x.number.Equals(partNumber));
                        }
                        else
                        {
                            assembly = new E3Assembly(partNumber, "");
                            _parts.Add(assembly);
                            AddUsage(assembly);
                        }
                    }

                    E3Documentation doc;
                    String docTypeSuff = "";
                    allDocType.TryGetValue(docType, out docTypeSuff);
                    String docNumber = partNumber + docTypeSuff;

                    if (_docs.Exists(x => x.number.Equals(docNumber)))
                    {
                        doc = (E3Documentation)_docs.Find(x => x.number.Equals(docNumber));
                    }
                    else
                    {
                        doc = new E3Documentation(docNumber, "", job.GetPath(), docNumber, docType, sheetN);
                        _docs.Add(doc);

                        if (!E3WGMForm.useRefLinkForDocumentation)
                        {
                            if (String.IsNullOrEmpty(wchID))
                            {
                                assembly.AddDescribe(doc.number, E3PartDescribe.TypeNumber);
                            }
                            else
                            {
                                doc.oidMaster = wchID;
                                assembly.AddDescribe(doc.oidMaster, E3PartDescribe.TypeOIDMaster);
                            }

                        }
                        else
                        {
                            if (String.IsNullOrEmpty(wchID))
                            {
                                assembly.AddReference(doc.number, E3PartReference.TypeNumber);
                            }
                            else
                            {
                                doc.oidMaster = wchID;
                                assembly.AddReference(doc.oidMaster, E3PartReference.TypeOIDMaster);
                            }
                        }
                    }
                    doc.AddSheet(sheet, sheetN, docFormat);

                    if (!String.IsNullOrEmpty(docName))
                    {
                        doc.name = docName;
                    }
                    if (String.IsNullOrEmpty(doc.oidMaster) && !String.IsNullOrEmpty(wchID))
                    {
                        doc.oidMaster = wchID;
                        E3PartDescribe describe = this.Describes.Find(x => x.value == doc.number);
                        describe.updateDescribe(doc);
                    }
                    else if (!String.IsNullOrEmpty(doc.oidMaster) && !String.IsNullOrEmpty(wchID) && !String.Equals(doc.oidMaster, wchID))
                    {
                        throw new Exception("ОШИБКА: oidMaster документации не совпадает с Windchill");
                    }
                    sheet.SetAssignment(partNumber);
                    sheet.SetAttributeValue(AttrsName.getAttrsName("partNumber"), doc.number);
                    sheet.SetAttributeValue(AttrsName.getAttrsName("docTypeSuff"), "");
                }
                else
                {
                    sheet.Display();
                    EditSheetAttrForm editSheetAttrForm = new EditSheetAttrForm(partNumber, docType, docFormat, sheetName);
                    editSheetAttrForm.ShowDialog();
                    if (editSheetAttrForm.DialogResult.Equals(DialogResult.OK))
                    {
                        sheet.SetAssignment(editSheetAttrForm.getPartNumber());
                        sheet.SetAttributeValue(AttrsName.getAttrsName("partNumber"), editSheetAttrForm.getNumber());
                        sheet.SetAttributeValue(AttrsName.getAttrsName("docTypeSuff"), "");
                        sheet.SetAttributeValue(AttrsName.getAttrsName("docType"), editSheetAttrForm.getDocType());
                        sheet.SetAttributeValue(AttrsName.getAttrsName("docFormat"), editSheetAttrForm.getDocFormat());
                        sheet.SetName(editSheetAttrForm.getSheetN());
                        // i--;
                        continue;
                    }
                    else
                    {
                        throw new NotImplementedException();
                    }
                }
            }
        }
*/

//        internal void updateDescribes(Document tempE3Doc)
//        {
        /*
            E3PartDescribe e3Desc = null;
            if (e3Desc == null && tempE3Doc.oidMaster != null && tempE3Doc.oidMaster != "")
            {
                e3Desc = Describes.Find(x => x.oidMaster == tempE3Doc.oidMaster);
            }

            if (e3Desc == null && tempE3Doc.number != null && tempE3Doc.number != "")
            {
                e3Desc = Describes.Find(x => x.number == tempE3Doc.number);
            }

            if (e3Desc == null)
            {
                foreach (Part part in Parts)
                {
                    if (part is E3Assembly)
                    {
                        if (e3Desc == null && tempE3Doc.oidMaster != null && tempE3Doc.oidMaster != "")
                        {
                            e3Desc = (part as E3Assembly).Describes.Find(x => x.oidMaster == tempE3Doc.oidMaster);
                        }

                        if (e3Desc == null && tempE3Doc.number != null && tempE3Doc.number != "")
                        {
                            e3Desc = (part as E3Assembly).Describes.Find(x => x.number == tempE3Doc.number);
                        }

                        if (e3Desc != null)
                        {
                            break;
                        }
                    }
                }
            }

            if(e3Desc == null)
            {
                throw new Exception("ОШИБКА: Отсутствует связь для документа "+tempE3Doc.number + "("+tempE3Doc.oidMaster+")");
            }

            e3Desc.updateDescribe(tempE3Doc);*/
 //       }

        internal void updateUsages(Part tempPart)
        {
            E3PartUsage e3Usage = null;
            if (e3Usage == null && tempPart.oidMaster != null && tempPart.oidMaster != "")
            {
                e3Usage = Usages.Find(x => x.oidMaster == tempPart.oidMaster);
            }

            if (e3Usage == null && tempPart is E3Part && (tempPart as E3Part).ATR_E3_ENTRY != null && (tempPart as E3Part).ATR_E3_ENTRY != "")
            {
                e3Usage = Usages.Find(x => x.ATR_E3_ENTRY == (tempPart as E3Part).ATR_E3_ENTRY);
            }

            if (e3Usage == null && tempPart is E3Cable && (tempPart as E3Cable).ATR_E3_ENTRY != null && (tempPart as E3Cable).ATR_E3_ENTRY != "" && (tempPart as E3Cable).ATR_E3_WIRETYPE != null && (tempPart as E3Cable).ATR_E3_WIRETYPE != "")
            {
                e3Usage = Usages.Find(x => x.ATR_E3_ENTRY == (tempPart as E3Cable).ATR_E3_ENTRY && x.ATR_E3_WIRETYPE == (tempPart as E3Cable).ATR_E3_WIRETYPE);
            }

            if (e3Usage == null && tempPart.number != null && tempPart.number != "")
            {
                e3Usage = Usages.Find(x => x.number == tempPart.number);
            }

            if (e3Usage != null)
            {
                e3Usage.updateUsage(tempPart);
            }

            foreach (Part part in Parts)
            {
                if (part is E3Assembly)
                {
                    if (e3Usage == null && tempPart.oidMaster != null && tempPart.oidMaster != "")
                    {
                        e3Usage = (part as E3Assembly).Usages.Find(x => x.oidMaster == tempPart.oidMaster);
                    }

                    if (e3Usage == null && tempPart is E3Part && (tempPart as E3Part).ATR_E3_ENTRY != null && (tempPart as E3Part).ATR_E3_ENTRY != "")
                    {
                        e3Usage = (part as E3Assembly).Usages.Find(x => x.ATR_E3_ENTRY == (tempPart as E3Part).ATR_E3_ENTRY);
                    }

                    if (e3Usage == null && tempPart is E3Cable && (tempPart as E3Cable).ATR_E3_ENTRY != null && (tempPart as E3Cable).ATR_E3_ENTRY != "" && (tempPart as E3Cable).ATR_E3_WIRETYPE != null && (tempPart as E3Cable).ATR_E3_WIRETYPE != "")
                    {
                        e3Usage = (part as E3Assembly).Usages.Find(x => x.ATR_E3_ENTRY == (tempPart as E3Cable).ATR_E3_ENTRY && x.ATR_E3_WIRETYPE == (tempPart as E3Cable).ATR_E3_WIRETYPE);
                    }

                    if (e3Usage == null && tempPart.number != null && tempPart.number != "")
                    {
                        e3Usage = (part as E3Assembly).Usages.Find(x => x.number == tempPart.number);
                    }

                    if (e3Usage != null)
                    {
                        e3Usage.updateUsage(tempPart);
                    }
                }
            }

            if (e3Usage == null)
            {
                throw new Exception("ОШИБКА: Отсутствует связь для компонента " + tempPart.number + "(" + tempPart.oidMaster + ")");
            }
        }


/*
        internal E3ProjectDocument getE3ProjectDocument()
        {
            List<Document> listE3ProjectDocument = Docs.FindAll(x => x is E3ProjectDocument);

            if (listE3ProjectDocument.Count != 1)
            {
                throw new Exception("ОШИБКА: Количество документов проекта E3.series равно " + listE3ProjectDocument.Count);
            }

            return (E3ProjectDocument)listE3ProjectDocument[0];
        }
*/

        internal E3PartDescribe getE3ProjectDocumentDescribe(string type, string value)
        {
            List<E3PartDescribe> listE3PartDescribe = Describes.FindAll(x => String.Equals(x.type, type) && String.Equals(x.value, value));

            if (listE3PartDescribe.Count != 1)
            {
                throw new Exception("ОШИБКА: Количество связей документов проекта E3.series равно " + listE3PartDescribe.Count);
            }

            return listE3PartDescribe[0];
        }

    }
}
