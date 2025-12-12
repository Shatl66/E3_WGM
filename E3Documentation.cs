using e3;
using System;
using System.Collections.Generic;
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

        string docFormatValues = "A0|A0x2|A0x3|A1|A1x3|A1x4|A2|A2x3|A2x4|A2x5|A3|A3x3|A3x4|A3x5|A3x6|A3x7|A4|A4x3|A4x4|A4x5|A4x6|A4x7|A4x8|A4x9";

        private SortedDictionary<int, object> dSheetIds = new SortedDictionary<int, object>();
        private int listov = 0;

        public E3Documentation(string docNumber, string docName, string filePath, string fileName, string docType, int listN) : base(docNumber, docName, filePath, fileName)
        {

            this.filePath = E3WGMForm.job.GetPath() + E3WGMForm.job.GetName() + "_PDF\\";
            this.fileName = this.number + ".pdf";
            if (docType == null || docType == "")
            {
                throw new Exception("ОШИБКА: Документ " + docNumber + ": Параметр листа \"Тип документа\" должен быть заполнен.");
            }
            if (docType == null || docType == "")
            {
                throw new Exception("ОШИБКА: Документ " + docNumber + ": Параметр листа \"Листов\" должен быть больше 0.");
            }
            this._docType = docType;
            if (this.listov > listN)
            {
                this.listov = listN;
            }
            this.dSheetIds.Add(0, null);
        }

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

        internal void merge(E3Documentation tempE3Doc)
        {
            E3PartDescribe describe = null;
            if (String.IsNullOrEmpty(oidMaster) && !String.IsNullOrEmpty(tempE3Doc.oidMaster))
            {
                if (E3WGMForm.project.Describes.Exists(x => x.value == number))
                {
                    describe = E3WGMForm.project.Describes.Find(x => x.value == number);
                }
                else
                {
                    foreach (Part part in E3WGMForm.project.Parts)
                    {
                        if ((part is E3Assembly) && (part as E3Assembly).Describes.Exists(x => x.value == number))
                        {
                            describe = (part as E3Assembly).Describes.Find(x => x.value == number);
                        }
                    }
                }

                if (describe != null)
                {
                    describe.updateDescribe(tempE3Doc);
                }
            }
            else if (!String.IsNullOrEmpty(oidMaster)
                && !String.IsNullOrEmpty(tempE3Doc.oidMaster)
                && !String.Equals(oidMaster, tempE3Doc.oidMaster))
            {
                throw new Exception("ОШИБКА: oidMaster документации не совпадает с Windchill");
            }

            E3PartReference reference = null;
            if (String.IsNullOrEmpty(oidMaster) && !String.IsNullOrEmpty(tempE3Doc.oidMaster))
            {
                if (E3WGMForm.project.References.Exists(x => x.value == number))
                {
                    reference = E3WGMForm.project.References.Find(x => x.value == number);
                }
                else
                {
                    foreach (Part part in E3WGMForm.project.Parts)
                    {
                        if ((part is E3Assembly) && (part as E3Assembly).References.Exists(x => x.value == number))
                        {
                            reference = (part as E3Assembly).References.Find(x => x.value == number);
                        }
                    }
                }

                if (reference != null)
                {
                    reference.updateReference(tempE3Doc);
                }
            }
            else if (!String.IsNullOrEmpty(oidMaster)
                && !String.IsNullOrEmpty(tempE3Doc.oidMaster)
                && !String.Equals(oidMaster, tempE3Doc.oidMaster))
            {
                throw new Exception("ОШИБКА: oidMaster документации не совпадает с Windchill");
            }

            updateDoc(tempE3Doc);
        }

        private void updateDoc(E3Documentation tempE3Doc)
        {
            if (String.IsNullOrEmpty(this.oidMaster))
            {
                this.oidMaster = tempE3Doc.oidMaster;
            }
            else if (!String.IsNullOrEmpty(oidMaster)
                && !String.IsNullOrEmpty(tempE3Doc.oidMaster)
                && !String.Equals(oidMaster, tempE3Doc.oidMaster))
            {
                throw new Exception("ОШИБКА: oidMaster документации не совпадает с Windchill");
            }

            this.oid = tempE3Doc.oid;
            this.number = tempE3Doc.number;
            this.name = tempE3Doc.name;
            this.folder = tempE3Doc.folder;
            //this.filePath = E3WGMForm.project.getJob().GetPath() + E3WGMForm.project.number + "_PDF\\";
            this.fileName = this.number + ".pdf";
            this.contextOid = tempE3Doc.contextOid;
            this.ATR_BOM_RS = tempE3Doc.ATR_BOM_RS;
            this.ATR_DOC_TYPE = tempE3Doc.ATR_DOC_TYPE;
            updateAttribute();
        }

        private void updateAttribute()
        {
            e3Sheet sheet = E3WGMForm.project.getJob().CreateSheetObject();
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
                e3Text text = E3WGMForm.project.getJob().CreateTextObject();
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
            return new Object[] {oidMaster,
                                    oid,
                                    number,
                                    name,
                                    ATR_BOM_RS,
                                    ATR_DOC_TYPE,
                                    convertDocFormatForRow()};
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

        internal void AddSheet(e3Sheet sheet, int list, string docFormat)
        {
            if (list <= 0)
            {
                throw new Exception("ОШИБКА: Лист " + sheet.GetName() + ": Параметр листа \"Лист\" должен быть больше 0.");
            }
            if (dSheetIds.ContainsKey(list))
            {
                throw new Exception("ОШИБКА: Лист " + sheet.GetName() + ": Параметр листа \"Лист\" должен быть уникальным.");
            }

            dSheetIds.Add(list, sheet.GetId());

            if (!String.IsNullOrEmpty(docFormat))
            {
                if (!ATR_DOC_FORMAT.Exists(x => x == docFormat))
                {
                    ATR_DOC_FORMAT.Add(docFormat);
                }
            }

        }

        internal void ExportToPDF()
        {
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            String pdfPath = filePath + fileName;
            if (E3WGMForm.project.getJob().ExportPDF(pdfPath, dSheetIds.Values.ToArray(), 0 + 16 + 4096, null) == 0)
            {
                throw new Exception("ОШИБКА: Не удалось выполнить экспорт в PDF документа " + number);
            }
        }
    }
}
