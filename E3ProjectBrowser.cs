using e3;
using E3SetAdditionalPart;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E3_WGM
{
    public partial class E3ProjectBrowser : UserControl
    {
        public E3ProjectBrowser()
        {
            InitializeComponent();
        }


        public void RefreshData(E3Assembly umens_e3projec)
        {
            textBoxNumberProject.Text = umens_e3projec.number;
            textBoxNameProject.Text = umens_e3projec.name;
            textBoxRestrict.Text = E3WGMForm.UtilsInstance.restrictProject;

            if (umens_e3projec.name != "Temp_Name")
            {
                textBoxNumberProject.ReadOnly = true;
                textBoxNameProject.ReadOnly = true;
            }
        }

        /// <summary>
        /// Из ТЗ
        /// Кнопка "Выгрузки" во вкладке "Проект" должна всегда выгружать проект Е3 целиком, не зависимо от того,
        /// что выделено в данный момент в дереве листов проекта, и создавать ссылочную связь со всеми СЧ, для которых создана схема или сборочный.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonUploadProjectDoc_Click(object sender, EventArgs e)
        {
            e3Job job = E3WGMForm.UtilsInstance.ConnectToE3Series(); // нужен только job
            String tempPath = E3WGMForm.UtilsInstance.tempPathForDoc;
            String prjFileName = job.GetName() + ".e3s"; // E3WGMForm.UtilsInstance.umens_e3project.number + ".e3s";

            E3Documentation prjDocument = new E3Documentation(E3WGMForm.public_umens_e3project, // любую Assembly, ее не используем для выгрузки
                                                                    job.GetName(), // тут обозначение СЧ
                                                                    "пока не знаю", // в Windchill узнаем наименование
                                                                    job,
                                                                    prjFileName,
                                                                    "E3ProjectDoc", // признак, что создаем E3ProjectDoc
                                                                    E3WGMForm.UtilsInstance.tempPathForDoc);


            prjDocument.numberPartsForE3ProjectDocument = E3WGMForm.UtilsInstance.numberPartsForE3ProjectDocument; // тоже признак, что создаем E3ProjectDoc

            //E3WGMForm.UtilsInstance.SyncE3Document(prjDocument);

            // 1. создаем WTDocument проекта Е3 в Windchill если его там еще нет
                string jsonDocumentationFromWindchill = "";

                MemoryStream stream = new MemoryStream();
                DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                settings.UseSimpleDictionaryFormat = true;
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(E3Documentation), settings);
                ser.WriteObject(stream, prjDocument);
                stream.Position = 0;
                StreamReader sr = new StreamReader(stream);
                string jsonDocumentation = sr.ReadToEnd();
                //В классах Windchill Андреем прописано пространство имен E3WGM. Я пока использую эти же классы, поэтому нужно сопоставлять E3WGM и мое E3_WGM
                jsonDocumentation = "{\"__type\":\"E3Documentation:#E3WGM\"," + jsonDocumentation.Substring(1);

                try
                {
                    jsonDocumentationFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonDocumentation, "netmarkets/jsp/by/iba/e3/http/createE3Documentation.jsp");
                }
                catch (Exception ex)
                {
                    UmensLogger.Log($"Создание документа Е3 проект в Windchill. Сообщение Windchill: {ex.Message}");
                    return;
                }

                // Обратная замена при десериализации. Правильнее было бы прописать везде - [DataContract(Namespace = "E3WGM")]
                jsonDocumentationFromWindchill = jsonDocumentationFromWindchill.Replace("E3Documentation:#E3WGM", "E3Documentation:#E3_WGM");

                MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonDocumentationFromWindchill));
                DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Documentation), settings);
                E3Documentation wchDoc = (E3Documentation)ser2.ReadObject(stream2);
                prjDocument.updateDoc(wchDoc);


            // 2.выгружаем файл проекта Е3 на локальный диск пользователя в папку заданную в конфиг.файле windchillserver.json
            job.SaveAs( Path.Combine(tempPath, prjFileName));

            // 3. передаем файл в Windchill где он будет привязан к документу как содержимое
            MemoryStream stream3 = new MemoryStream();
            DataContractJsonSerializerSettings settings3 = new DataContractJsonSerializerSettings();
            settings3.UseSimpleDictionaryFormat = true;
            DataContractJsonSerializer ser3 = new DataContractJsonSerializer(typeof(E3Documentation), settings3);
            ser3.WriteObject(stream3, prjDocument);
            stream3.Position = 0;
            StreamReader sr3 = new StreamReader(stream3);
            string jsonE3Doc = sr3.ReadToEnd();
            jsonE3Doc = "{\"__type\":\"E3Documentation:#E3WGM\"," + jsonE3Doc.Substring(1);
            E3WGMForm.wchHTTPClient.updateE3DocumentationContent(jsonE3Doc, "netmarkets/jsp/by/iba/e3/http/updateE3DocumentationContent.jsp", prjDocument.filePath, prjDocument.fileName);

            UmensLogger.Log($"Выгрузка проекта Е3 {prjDocument.number} в Windchill завершена");

            E3WGMForm.UtilsInstance.DisconnectFromE3Series();
        }
    }
}
