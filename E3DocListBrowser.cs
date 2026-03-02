using e3;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

namespace E3_WGM
{
    public partial class E3DocListBrowser : UserControl
    {
        public E3DocListBrowser()
        {
            InitializeComponent();

            buttonUploadDoc.Enabled = false;

            // Подписываемся только на события, связанные с чекбоксами
            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
            dataGridView1.CurrentCellDirtyStateChanged += dataGridView1_CurrentCellDirtyStateChanged;

        }


        public void Refresh (E3Assembly umens_e3project)
        {
            dataGridView1.Rows.Clear();

            foreach (E3Documentation doc in umens_e3project.Docs)                
            {
                int row = dataGridView1.Rows.Add( doc.getDataForRow());
                dataGridView1.Rows[row].Tag = doc;

                if (!String.IsNullOrEmpty(doc.oidMaster)) // зеленым цветом отмечаем уже имеющиеся в Wch документы
                {
                    dataGridView1.Rows[row].DefaultCellStyle.BackColor = Color.GreenYellow;
                }

            }

            }

        private void buttonUploadDoc_Click(object sender, EventArgs e)
        {
            e3Job job = E3WGMForm.UtilsInstance.ConnectToE3Series(); // нужен только job
            //UmensLogger.Log($"Начинаю выгрузку документов в Windchill");

            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Проверяем, отмечен ли чекбокс в первом столбце (isRowSelected)
                    DataGridViewCheckBoxCell checkBoxCell = row.Cells["isRowSelected"] as DataGridViewCheckBoxCell;

                    if (checkBoxCell != null && checkBoxCell.Value != null)
                    {
                        bool isSelected = Convert.ToBoolean(checkBoxCell.Value);

                        if (isSelected)
                        {
                            if (row.Tag is E3Documentation doc) // проверили и сразу инициализировали doc
                            {
                                // 1. создаем WTDocument в Windchill если его там еще нет
                                if (String.IsNullOrEmpty(doc.oidMaster))
                                {
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
                                        jsonDocumentationFromWindchill = E3WGMForm.wchHTTPClient.getJSON(jsonDocumentation, "netmarkets/jsp/by/iba/e3/http/createE3Documentation.jsp");
                                    }
                                    catch (Exception ex)
                                    {
                                        UmensLogger.Log($"Создание документа {doc.number} в Windchill. Сообщение Windchill: {ex.Message}");
                                    }

                                    // Обратная замена при десериализации. Правильнее было бы прописать везде - [DataContract(Namespace = "E3WGM")]
                                    jsonDocumentationFromWindchill = jsonDocumentationFromWindchill.Replace("E3Documentation:#E3WGM", "E3Documentation:#E3_WGM");

                                    MemoryStream stream2 = new MemoryStream(Encoding.UTF8.GetBytes(jsonDocumentationFromWindchill));
                                    DataContractJsonSerializer ser2 = new DataContractJsonSerializer(typeof(E3Documentation), settings);
                                    E3Documentation wchDoc = (E3Documentation)ser2.ReadObject(stream2);
                                    doc.updateDoc(wchDoc);

                                    row.DefaultCellStyle.BackColor = Color.GreenYellow; // отметили, что документ в Wch уже есть
                                }

                                // 2.выгружаем документ из Е3 на локальный диск пользователя в папку заданную в конфиг.файле windchillserver.json
                                doc.ExportToPDF(job);

                                // 3. передаем файл в Windchill где он будет привязан к документу как содержимое
                                MemoryStream stream3 = new MemoryStream();
                                DataContractJsonSerializerSettings settings3 = new DataContractJsonSerializerSettings();
                                settings3.UseSimpleDictionaryFormat = true;
                                DataContractJsonSerializer ser3 = new DataContractJsonSerializer(typeof(E3Documentation), settings3);
                                ser3.WriteObject(stream3, doc);
                                stream3.Position = 0;
                                StreamReader sr3 = new StreamReader(stream3);
                                string jsonE3Doc = sr3.ReadToEnd();
                                jsonE3Doc = "{\"__type\":\"E3Documentation:#E3WGM\"," + jsonE3Doc.Substring(1);
                                E3WGMForm.wchHTTPClient.updateE3DocumentationContent(jsonE3Doc, "netmarkets/jsp/by/iba/e3/http/updateE3DocumentationContent.jsp", doc.filePath, doc.fileName);

                                UmensLogger.Log($"Выгрузка документа {doc.number} в Windchill завершена");
                                // чтобы не дать повторно отправить тот же документ, пока не выполнят синхронизацию и таким образом считают новые данные из Е3
                                checkBoxCell.Value = false;
                                row.ReadOnly = true;

                            }
                        }
                    }
                }

                E3WGMForm.UtilsInstance.DisconnectFromE3Series();  //  E3WGMForm.UtilsInstance.getE3Job() вновь создал app и job            
            }
            catch (Exception ex)
            {
                UmensLogger.Log($"Выгрузка документа. Сообщение Windchill: {ex.Message}");
                E3WGMForm.UtilsInstance.DisconnectFromE3Series();           
            }            
        }

        // Возникает при щелчке на содержимом ячейки
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["isRowSelected"].Index && e.RowIndex >= 0)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                CheckCheckboxes();
            }

        }

        // Возникает при изменении состояния ячейки
        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.IsCurrentCellDirty && dataGridView1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                CheckCheckboxes();
            }

        }

        private void CheckCheckboxes()
        {
            bool hasChecked = false;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell checkBoxCell = row.Cells["isRowSelected"] as DataGridViewCheckBoxCell;
                if (checkBoxCell != null && checkBoxCell.Value != null)
                {
                    if (Convert.ToBoolean(checkBoxCell.Value))
                    {
                        hasChecked = true;
                        break;
                    }
                }
            }

            buttonUploadDoc.Enabled = hasChecked;
        }

    }
}
