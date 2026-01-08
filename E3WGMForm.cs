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

        public E3WGMForm()
        {
            InitializeComponent();        
        }


        private void E3WGMForm_Load(object sender, EventArgs e)
        {
            Dictionary<string, string> allWCHServer = new Dictionary<string, string>();
            if (File.Exists("windchillserver.json"))
            {
                String jsonWCHServer = "";
                using (StreamReader streamReader = new StreamReader("windchillserver.json"))
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
            wchHTTPClient = new WindchillHTTPClient(serverProtocol, serverName, Boolean.Parse(ignoreSSLPolicyErrors));

            /*
            String tempString = "false";
            allWCHServer.TryGetValue("useRefLinkForDocumentation", out tempString);
            useRefLinkForDocumentation = Boolean.Parse(tempString);

            tempString = "true";
            allWCHServer.TryGetValue("allowUnloadStructure", out tempString);
            allowUploadStructure = Boolean.Parse(tempString);
            */

            public_umens_e3project = CreateAndFillUmensE3Project();

            // Подписываем E3WGMForm на событие синхронизации генерируемое кнопкой "Синхронизация"
            e3CommonControl1.SynchronizeClicked += E3CommonControl1_SynchronizeClicked;

        }


        public E3Assembly CreateAndFillUmensE3Project()
        {
            //E3Project umens_e3project = new E3Project("Temp_Number", "Temp_Name");
            E3Assembly umens_e3project = new E3Assembly("Temp_Number", "Temp_Name");
            E3StructureTreeTraverser traverser = new E3StructureTreeTraverser(umens_e3project); // подключается к объектам Е3 (app, job, tree и др.)
            traverser.FindAllWTPartsFromSelectedFolder();
            traverser.SyncronizeE3ProjectDataWithWindchill();
            
            List<string> errorMessages = traverser.errorMessages;
            // Показываем все ошибки одним MessageBox
            if (errorMessages.Count > 0)
            {
                string allErrors = string.Join("\n\n", errorMessages);
                MessageBox.Show(allErrors, "Ошибки при чтении устройств и синхронизации с Windchill",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            traverser.DisconnectFromE3Series(); // "отключается" от объектов Е3

            return umens_e3project;
        }




        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        ///  Обновляем все данные в зависимости от выбранной папки в дереве листов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void E3CommonControl1_SynchronizeClicked(object sender, EventArgs e)
        {
            public_umens_e3project = CreateAndFillUmensE3Project(); // рассчитали новые данные

            // Обновляем активную вкладку (таблицу)
            RefreshActiveTab();

            ///E3Log.AppendText($"[{DateTime.Now}] Проект синхронизирован\r\n");
        }

        private void RefreshActiveTab()
        {
            if (tabControl1.SelectedTab == tabPageStructureBrowser)
            {
                e3StructureBrowser1.Refresh();
            }
            else if (tabControl1.SelectedTab == tabPageProject)
            {
                // e3ProjectBrowser1.Refresh(); // раскомментировать, если есть метод
            }
            else if (tabControl1.SelectedTab == tabPageDocListBrowser)
            {
                // e3DocListBrowser1.Refresh(); // раскомментировать, если есть метод
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Если выбрана вкладка "Состав"
            if (tabControl1.SelectedTab == tabPageStructureBrowser)
            {
                e3StructureBrowser1.Refresh();
            }
            else if (tabControl1.SelectedTab == tabPageProject)
            {
               // e3ProjectBrowser1.Refresh(); // если есть такой метод
            }
            else if (tabControl1.SelectedTab == tabPageDocListBrowser)
            {
               // e3DocListBrowser1.Refresh(); // если есть такой метод
            }

        }
    }
}
