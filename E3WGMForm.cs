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

        //public static e3Application app;
        //public static e3Job job;
        public static E3Project public_umens_e3project;
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

        }

        private E3Project CreateAndFillUmensE3Project()
        {
            E3Project umens_e3project = new E3Project("Temp_Number", "Temp_Name");
            E3StructureTreeTraverser traverser = new E3StructureTreeTraverser( umens_e3project);
            traverser.FindAllWTPartsFromSelectedFolder();
            traverser.DisconnectFromE3Series();            

            return umens_e3project;
        }

        private void ReadParts()
        {
           // app.PutInfo(0, "старт ReadParts()");
        }








        private void E3WGMForm_FormClosed(object sender, FormClosedEventArgs e)
        {
           // DisconnectFromE3Series();
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
