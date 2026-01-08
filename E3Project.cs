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
        */

        /*
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
