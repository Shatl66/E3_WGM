using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E3_WGM
{
    public partial class WindchillLoginForm : Form
    {
        WindchillHTTPClient wchHTTPClient;
        public WindchillLoginForm(WindchillHTTPClient wchHTTPClient)
        {
            this.wchHTTPClient = wchHTTPClient;
            InitializeComponent();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void LoginButton_Click(object sender, EventArgs e)
        {
            try
            {
                wchHTTPClient.checkLogin(UserNameTextBox.Text, PasswordTextBox.Text);
                using (StreamWriter streamWriter = new StreamWriter("userlogin.txt"))
                {
                    streamWriter.WriteLine(UserNameTextBox.Text);
                    streamWriter.WriteLine(Encrypt(PasswordTextBox.Text));
                }
                Dispose();
            }
            catch (HttpRequestException ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void WindchillLoginForm_Load(object sender, EventArgs e)
        {
            if (File.Exists("userlogin.txt"))
            {
                using (StreamReader streamReader = new StreamReader("userlogin.txt"))
                {
                    UserNameTextBox.Text = streamReader.ReadLine();
                    PasswordTextBox.Text = Decrypt(streamReader.ReadLine());
                }
            }
        }

        string hash = "j345iuy#56fyh";

        private string Encrypt(string value)
        {
            UTF8Encoding utf8 = new UTF8Encoding();
            byte[] data = utf8.GetBytes(value);
            using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
            {
                byte[] keys = md5.ComputeHash(utf8.GetBytes(hash));
                using (TripleDESCryptoServiceProvider tripDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                {
                    ICryptoTransform transform = tripDes.CreateEncryptor();
                    byte[] result = transform.TransformFinalBlock(data, 0, data.Length);
                    return Convert.ToBase64String(result);
                }
            }
        }

        private string Decrypt(string value)
        {
            UTF8Encoding utf8 = new UTF8Encoding();
            byte[] data = Convert.FromBase64String(value);
            using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
            {
                byte[] keys = md5.ComputeHash(utf8.GetBytes(hash));
                using (TripleDESCryptoServiceProvider tripDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                {
                    ICryptoTransform transform = tripDes.CreateDecryptor();
                    byte[] result = transform.TransformFinalBlock(data, 0, data.Length);
                    return utf8.GetString(result);
                }
            }
        }
    }
}
