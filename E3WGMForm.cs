using e3;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E3_WGM
{
    public partial class E3WGMForm : Form
    {

        //public static E3Project public_umens_e3project;
        public static E3Assembly public_umens_e3project;
        public static WindchillHTTPClient wchHTTPClient;
        public static String tempPathForDoc = ""; // папка на ПК пользователя Е3 для временной выгрузки документов перед их передачей в Windchill

        private static Utils _Utils = new Utils();
        public static Utils UtilsInstance //TODO  Utils в static класс ( подобно ESKDHelperReport). см. здесь UmensLogger.cs
        {
            get
            {
                return _Utils;
            }
        }

        public E3WGMForm()
        {
            InitializeComponent();        
        }


        private void E3WGMForm_Load(object sender, EventArgs e)
        {

            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string configFile = Path.Combine(appDir, "windchillserver.json");

            Dictionary<string, string> allWCHServer = new Dictionary<string, string>();
            if (File.Exists( configFile))
            {
                String jsonWCHServer = "";
                using (StreamReader streamReader = new StreamReader( configFile))
                {
                    jsonWCHServer = streamReader.ReadToEnd();
                }

                using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonWCHServer)))
                {
                    DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                    settings.UseSimpleDictionaryFormat = true;
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Dictionary<string, string>), settings);
                    allWCHServer = ser.ReadObject(stream) as Dictionary<string, string>;
                }
            }
            
            string serverProtocol = "";
            string serverName = "";
            string ignoreSSLPolicyErrors = "";
            allWCHServer.TryGetValue("serverProtocol", out serverProtocol);
            allWCHServer.TryGetValue("serverName", out serverName);
            allWCHServer.TryGetValue("ignoreSSLPolicyErrors", out ignoreSSLPolicyErrors);

            string subfolderPath = Path.Combine(appDir, "TEMP");
            Directory.CreateDirectory(subfolderPath); // создаст, если нет, иначе ничего не сделает
            _Utils.tempPathForDoc = subfolderPath;

            wchHTTPClient = new WindchillHTTPClient(serverProtocol, serverName, Boolean.Parse(ignoreSSLPolicyErrors));

            configFile = Path.Combine(appDir, "doctype.json");
            Dictionary<string, string> typeDocuments = new Dictionary<string, string>();
            if (File.Exists(configFile))
            {
                String jsonDocType = "";
                using (StreamReader streamReader = new StreamReader(configFile))
                {
                    jsonDocType = streamReader.ReadToEnd();
                }

                using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonDocType)))
                {
                    DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings();
                    settings.UseSimpleDictionaryFormat = true;
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Dictionary<string, string>), settings);
                    typeDocuments = ser.ReadObject(stream) as Dictionary<string, string>;
                }
            }
            _Utils.typeDocuments = typeDocuments;




            // Подписываем E3WGMForm на событие синхронизации генерируемое кнопкой "Синхронизация"
            e3CommonControl1.SynchronizeClicked += E3CommonControl1_SynchronizeClicked;

            UmensLogger.LogControl = E3Log;

            public_umens_e3project = new E3Assembly("Temp_Number", "Temp_Name");
            _Utils.umens_e3project = public_umens_e3project; 

            _Utils.ConnectToE3Series();
            // _Utils.getRestrictivProject(); без выхода из приложения тут не считает новое значение

            public_umens_e3project = CreateAndFillUmensE3Project();
            _Utils.DisconnectFromE3Series(); // После вычисления всех данных "отключается" от Е3 чтобы пользователь Е3 мог в нем полноценно работать

            e3ProjectBrowser1.RefreshData(public_umens_e3project);

        }


        public E3Assembly CreateAndFillUmensE3Project() // вызывается при загрузке утилиты и при каждом нажатии кнопки "Синхронизация"
        {
            // перед каждой синхронизацией очищаем:
            _Utils.errorMessages = new List<string>();
            _Utils.numberPartsForE3ProjectDocument = new List<string>();
            _Utils.getRestrictivProject();


            if (_Utils.FindAllWTPartsFromSelectedFolder())
            {
                _Utils.SyncronizeE3ProjectDataWithWindchill();
                _Utils.FindAllWTDocsFromSelectedFolder(); // выбранная пользователем папка в дереве листов уже проверена и определена (selectedRootNodeId) в методе выше            
                _Utils.SyncronizeE3DocsWithWindchill();
            }
            else
            { // Если пользователь выбрал неправильно папку в ДеревеЛистов в Е3, то показываем пустые данные в утилите
                public_umens_e3project = new E3Assembly("Temp_Number", "Temp_Name");
                _Utils.umens_e3project = public_umens_e3project;                
            }

            List<string> errorMessages = _Utils.errorMessages;
            // Показываем все ошибки одним MessageBox
            if (errorMessages.Count > 0)
            {
                string allErrors = string.Join("\n\n", errorMessages);
                MessageBox.Show(allErrors, "Ошибки при чтении устройств и синхронизации с Windchill",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }           

            return _Utils.getFilled_Umens_e3project();
        }




        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        ///  Обновляем все данные в зависимости от выбранной папки в дереве листов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void E3CommonControl1_SynchronizeClicked(object sender, EventArgs e)
        {
            public_umens_e3project = new E3Assembly("Temp_Number", "Temp_Name");
            _Utils.umens_e3project = public_umens_e3project; // подготовка утилит к новым расчетам

            _Utils.ConnectToE3Series();
            public_umens_e3project = CreateAndFillUmensE3Project(); // Пользователь выбрал новую папку в ДеревеЛистов в Е3 и нажал кнопку синхронизации. Рассчитали новые данные
            _Utils.DisconnectFromE3Series(); // После вычисления всех данных "отключается" от Е3 чтобы пользователь Е3 мог в нем полноценно работать

            // Обновляем активную вкладку (таблицу)
            RefreshActiveTab();

            UmensLogger.Log($"{public_umens_e3project.number} Проект синхронизирован\r\n");
        }

        private void RefreshActiveTab()
        {
            if (tabControl1.SelectedTab == tabPageStructureBrowser)
            {
                e3StructureBrowser1.Refresh();
            }
            else if (tabControl1.SelectedTab == tabPageProject)
            {
                e3ProjectBrowser1.RefreshData(public_umens_e3project);
            }
            else if (tabControl1.SelectedTab == tabPageDocListBrowser)
            {
                e3DocListBrowser1.Refresh(public_umens_e3project);
            }
        }

        /// <summary>
        /// Метод начинает впервые наполнять таблицы данными расчитанными при загрузке утилиты и после кнопки "Синхронизация"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e) // обновить данные при переключении между вкладками
        {
            // Если выбрана вкладка "Состав"
            if (tabControl1.SelectedTab == tabPageStructureBrowser)
            {
                e3StructureBrowser1.Refresh();
            }
            else if (tabControl1.SelectedTab == tabPageProject)
            {
                e3ProjectBrowser1.RefreshData(public_umens_e3project);
            }
            else if (tabControl1.SelectedTab == tabPageDocListBrowser)
            {
                e3DocListBrowser1.Refresh(public_umens_e3project);
            }

        }
    }
}
