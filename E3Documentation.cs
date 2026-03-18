using e3;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    [DataContract]
    public class E3Documentation : Document
    {
        [DataMember]
        private string _bomRS = BomRSValues.getBomRSValue((int)BomRSEnum.DOCUMENTATION);
        internal string ATR_BOM_RS
        {
            get { return _bomRS; }
            set { _bomRS = value; }
        }
        [DataMember]
        private string _docType = "Сборочный чертеж";
        internal string ATR_DOC_TYPE
        {
            get { return _docType; }
            set { _docType = value; }
        }

        [DataMember]
        private List<string> _docFormat = new List<string>();
        internal List<string> ATR_DOC_FORMAT
        {
            get { return _docFormat; }
            set { }
        }

        [DataMember]
        private string _numberDescribedPart;
        internal string numberDescribedPart
        {
            get { return _numberDescribedPart; }
            set { }
        }

        [DataMember]
        private List<string> _numberPartsForE3PrjDoc = new List<string>();
        internal List<string> numberPartsForE3ProjectDocument
        {
            get { return _numberPartsForE3PrjDoc; }
            set { _numberPartsForE3PrjDoc = value; }
        }

        private E3Assembly assembly;

        // ? string docFormatValues = "A0|A0x2|A0x3|A1|A1x3|A1x4|A2|A2x3|A2x4|A2x5|A3|A3x3|A3x4|A3x5|A3x6|A3x7|A4|A4x3|A4x4|A4x5|A4x6|A4x7|A4x8|A4x9";

        private SortedDictionary<int, object> dSheetIds = new SortedDictionary<int, object>();
        private int listov = 0;

        public E3Documentation(E3Assembly assembly, string docNumber, string docName, e3Job job, string fileName, string docType, string tempPathForDoc) : base(docNumber, docName, job.GetPath(), fileName)
        {
            this.assembly = assembly;
            this._numberDescribedPart = this.assembly.number;

            this.filePath = tempPathForDoc; // job.GetPath() + job.GetName() + "_PDF\\";

            
            //TODO проверить, почему сразу не взять переданный сюда fileName ?
            if (fileName.EndsWith("e3s"))
                this.fileName = fileName;
            else
                this.fileName = this.number + ".pdf"; //TODO проверить, почему сразу не взять переданный сюда fileName ?



            if (docType == null || docType == "")
            {
                throw new Exception("ОШИБКА: Документ " + docNumber + ": Параметр листа \"Тип документа\" должен быть заполнен.");
            }

            this._docType = docType;

            this.dSheetIds.Add(0, null); // ????
        }


        /// <summary>
        /// Корректирует число листов у документа.
        /// Наращивает перечень используемых форматов
        /// Наращивает перечень id листов документа
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="docFormat"></param>
        internal void AddSheet(e3Sheet sheet, string docFormat)
        {
            String sheetName = sheet.GetName(); // это номер листа

            if (sheet.IsEmbedded() == 1) // т.е. это не реальный лист, а лист области formboard. Для публикации они не нужны
                return;

            
            int sheetN = -1;
            if (int.TryParse(sheetName, out sheetN)) // здесь у листов номера будут типа 1, 2
            {
                if (this.listov < sheetN)
                {
                   this.listov = sheetN;
                }
            }
            else
            {
                // тут надо ругаться, что пользователь изменил некорректно имя листа
            }            

            if (dSheetIds.ContainsKey( sheetN))
            {
                throw new Exception("ОШИБКА: Лист " + sheet.GetName() + ": Параметр листа \"Лист\" должен быть уникальным.");
            }

            dSheetIds.Add( sheetN, sheet.GetId());

            if (!String.IsNullOrEmpty(docFormat))
            {
                if (!ATR_DOC_FORMAT.Exists(x => x == docFormat))
                {
                    ATR_DOC_FORMAT.Add(docFormat);
                }
            }

        }



        /*
        private List<string> getDocFormatFromRow(string value)
        {
            List<string> localDocFormat = new List<string>();

            if (value != null && value != "")
            {
                String[] split = value.Split('|');
                for (int i = 0; i < split.Length; i++)
                {
                    if (!docFormatValues.Contains(split[i]))
                    {
                        throw new Exception("ОШИБКА: Формат документации (" + split[i] + ") заполнен не верно. Допускаются только цифры и латинские буквы (A x) ");
                    }
                    localDocFormat.Add(split[i]);
                }
            }

            return localDocFormat;
        }
        */

        internal void merge(e3Job job, E3Documentation wchDoc)
        {
            /* Мне непонятно зачем это нужно
            E3PartDescribe describe = null;
            if (String.IsNullOrEmpty( this.oidMaster) && !String.IsNullOrEmpty(wchDoc.oidMaster))
            {
                if ( assembly.Describes.Exists(x => x.value == number))
                {
                    describe = assembly.Describes.Find(x => x.value == number);
                }
                else
                {
                    foreach (Part part in assembly.Parts)
                    {
                        if ((part is E3Assembly) && (part as E3Assembly).Describes.Exists(x => x.value == number))
                        {
                            describe = (part as E3Assembly).Describes.Find(x => x.value == number);
                        }
                    }
                }

                if (describe != null)
                {
                    describe.updateDescribe(wchDoc);
                }
            }
            else if (!String.IsNullOrEmpty( this.oidMaster)
                && !String.IsNullOrEmpty(wchDoc.oidMaster)
                && !String.Equals(oidMaster, wchDoc.oidMaster))
            {
                throw new Exception("ОШИБКА: oidMaster документации не совпадает с Windchill");
            }

            E3PartReference reference = null;
            if (String.IsNullOrEmpty(oidMaster) && !String.IsNullOrEmpty(wchDoc.oidMaster))
            {
                if (assembly.References.Exists(x => x.value == number))
                {
                    reference = assembly.References.Find(x => x.value == number);
                }
                else
                {
                    foreach (Part part in assembly.Parts)
                    {
                        if ((part is E3Assembly) && (part as E3Assembly).References.Exists(x => x.value == number))
                        {
                            reference = (part as E3Assembly).References.Find(x => x.value == number);
                        }
                    }
                }

                if (reference != null)
                {
                    reference.updateReference(wchDoc);
                }
            }
            else if (!String.IsNullOrEmpty(oidMaster)
                && !String.IsNullOrEmpty(wchDoc.oidMaster)
                && !String.Equals(oidMaster, wchDoc.oidMaster))
            {
                throw new Exception("ОШИБКА: oidMaster документации не совпадает с Windchill");
            }
            
            updateDoc(job, wchDoc);
            */
        }

        internal void updateDoc( E3Documentation wchDoc)
        {
            if (String.IsNullOrEmpty(this.oidMaster))
            {
                this.oidMaster = wchDoc.oidMaster;
            }
            else if (!String.IsNullOrEmpty(oidMaster)
                && !String.IsNullOrEmpty(wchDoc.oidMaster)
                && !String.Equals(oidMaster, wchDoc.oidMaster))
            {
                throw new Exception("ОШИБКА: oidMaster документации не совпадает с Windchill");
            }

            this.oid = wchDoc.oid;
            this.number = wchDoc.number;
            this.name = wchDoc.name;
            this.folder = wchDoc.folder;
            //this.filePath = E3WGMForm.project.getJob().GetPath() + E3WGMForm.project.number + "_PDF\\";

            if (fileName.EndsWith("e3s"))
                this.fileName = wchDoc.fileName;
            else
                this.fileName = wchDoc.number + ".pdf"; //TODO проверить, почему сразу не взять переданный сюда fileName ?

            this.contextOid = wchDoc.contextOid;
            this.ATR_BOM_RS = wchDoc.ATR_BOM_RS;
            this.ATR_DOC_TYPE = wchDoc.ATR_DOC_TYPE;
            // Я это убрал   updateAttribute( job);
        }

        private void updateAttribute( e3Job job)
        {
            e3Sheet sheet = job.CreateSheetObject();
            foreach (KeyValuePair<int, object> sheetId in dSheetIds)
            {
                if (sheetId.Key == 0)
                {
                    continue;
                }
                sheet.SetId((int)sheetId.Value);
                sheet.SetAttributeValue("WCH_id", oidMaster);
                sheet.SetAttributeValue("docname", this.number);
                sheet.SetAttributeValue("doccode", "");
                sheet.SetAttributeValue(".DOCUMENT_TYPE", this.ATR_DOC_TYPE);

                dynamic sTextIds = null;
                int nTextIds = sheet.GetTextIds(ref sTextIds);
                e3Text text = job.CreateTextObject();
                for (int j = 1; j <= nTextIds; j++)
                {
                    text.SetId(sTextIds[j]);
                    switch (text.GetType())
                    {
                        case 507:
                            text.SetText(this.name);
                            break;
                    }
                }
            }
        }

        internal object[] getDataForRow()
        {
            return new Object[] {false,
                                    oidMaster,
                                    oid,
                                    number,
                                    name,
                                    ATR_BOM_RS,
                                    ATR_DOC_TYPE,
                                    convertDocFormatForRow(),
                                    listov
            };
        }

        private string convertDocFormatForRow()
        {
            string result = "";
            foreach (string format in _docFormat)
            {
                if (result != "")
                {
                    result = result + "|";
                }
                result = result + format;
            }
            return result;
        }


        internal void ExportToPDF(e3Job job)
        {
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            String pdfPath = filePath + fileName;
            if ( job.ExportPDF(pdfPath, dSheetIds.Values.ToArray(), 0 + 16 + 4096, null) == 0)
            {
                throw new Exception("ОШИБКА: Не удалось выполнить экспорт в PDF документа " + number);
            }
        }
    }
}
