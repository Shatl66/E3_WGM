using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace E3_WGM
{
    public class WindchillHTTPClient
    {
        String serverProtocol = "";
        String serverName = "";
        String location = "";
        private HttpClient client = null;
        public WindchillHTTPClient(String serverProtocol, String serverName, bool ignoreSSLPolicyErrors)
        {
            this.serverProtocol = serverProtocol;
            this.serverName = serverName;
            this.location = serverProtocol + "://" + serverName + "/Windchill";
            if (ignoreSSLPolicyErrors)
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
            }
            client = new HttpClient();
        }

        internal bool checkLogin(string userName, string password)
        {
            var authToken = Encoding.ASCII.GetBytes($"{userName}:{password}");
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(authToken));

            String jspFile = "netmarkets/jsp/by/iba/e3/http/checkLogin.jsp";
            String urlpath = location + "/" + jspFile;
            Task<HttpResponseMessage> taskResponse = client.GetAsync(urlpath);
            taskResponse.Wait();
            taskResponse.Result.EnsureSuccessStatusCode();
            return true;
        }

        /*
        internal string getJSON(string json, string jspFile) Код Андрея без обработки сообщения об ошике из Wondchill, всегда выводится "Ошибка код 500"
        {
            String urlpath = location + "/" + jspFile;
            var data = new StringContent(json, Encoding.UTF8, "application/json");
            Task<HttpResponseMessage> taskResponse = client.PostAsync(urlpath, data);
            taskResponse.Wait();
            taskResponse.Result.EnsureSuccessStatusCode();
            Task<byte[]> taskByteArray = taskResponse.Result.Content.ReadAsByteArrayAsync();
            taskByteArray.Wait();
            var responseString = Encoding.UTF8.GetString(taskByteArray.Result, 0, taskByteArray.Result.Length);
            return responseString;
        }
        */

        internal string getJSON(string json, string jspFile)
        {
            try
            {
                String urlpath = location + "/" + jspFile;
                var data = new StringContent(json, Encoding.UTF8, "application/json");

                Task<HttpResponseMessage> taskResponse = client.PostAsync(urlpath, data);
                taskResponse.Wait();

                HttpResponseMessage response = taskResponse.Result;
                response.EnsureSuccessStatusCode();

                //   Task<string> contentTask = response.Content.ReadAsStringAsync(); 
                //   contentTask.Wait();
                //   string responseContent = contentTask.Result; // тут надо разобраться с кракозябрами в имени 

                Task<byte[]> taskByteArray = response.Content.ReadAsByteArrayAsync();
                taskByteArray.Wait();
                String responseContent = Encoding.UTF8.GetString(taskByteArray.Result, 0, taskByteArray.Result.Length);

                // Проверяем различные форматы ошибок: или сообщение об ошике из jsp или непрелвиденная ошибка в Windchill или при http обмене
                CheckForErrorInResponse(responseContent);

                return responseContent;
            }
            catch (Exception ex)
            {
                throw new Exception($"Request failed: {ex.Message}", ex);
            }
        }

        private void CheckForErrorInResponse(string responseContent)
        {
            if (string.IsNullOrEmpty(responseContent))
                return;

            // Ищем структуру: {"error": {"message": "текст ошибки"}}. Например - {"error": {"message": "ERROR: E3.series project document 3103.09.65.030-E3 link with 2 part", "code": "MULTIPLE_LINKS"}}
            int errorStart = responseContent.IndexOf("\"error\"");
            if (errorStart < 0)
                return;

            // Ищем message внутри error (с учетом пробелов после двоеточия)
            int messageStart = responseContent.IndexOf("\"message\"", errorStart);
            if (messageStart >= 0)
            {
                // Ищем двоеточие после message
                int colonPos = responseContent.IndexOf(':', messageStart);
                if (colonPos >= 0)
                {
                    // Ищем начало строки значения (после двоеточия и возможных пробелов)
                    int valueStart = responseContent.IndexOf('\"', colonPos);
                    if (valueStart >= 0)
                    {
                        valueStart += 1; // Пропускаем открывающую кавычку
                        int valueEnd = responseContent.IndexOf('\"', valueStart);

                        if (valueEnd > valueStart)
                        {
                            string errorMessage = responseContent.Substring(valueStart, valueEnd - valueStart);
                            errorMessage = DecodeJsonString(errorMessage);
                            throw new Exception($"Server error: {errorMessage}"); // прерывает работу программы с осмысленным сообщением полученным из jsp
                        }
                    }
                }
            }

            // Ищем простую структуру: {"error": "текст ошибки"}
            int simpleErrorStart = responseContent.IndexOf("\"error\"", errorStart);
            if (simpleErrorStart >= 0)
            {
                int colonPos = responseContent.IndexOf(':', simpleErrorStart);
                if (colonPos >= 0)
                {
                    int valueStart = responseContent.IndexOf('\"', colonPos);
                    if (valueStart >= 0)
                    {
                        valueStart += 1;
                        int valueEnd = responseContent.IndexOf('\"', valueStart);

                        if (valueEnd > valueStart)
                        {
                            string errorMessage = responseContent.Substring(valueStart, valueEnd - valueStart);
                            errorMessage = DecodeJsonString(errorMessage);
                            throw new Exception($"Server error: {errorMessage}"); // прерывает работу программы с сообщением о непредвиденной ошибке
                        }
                    }
                }
            }
        }

        private string DecodeJsonString(string jsonString)
        {
            return jsonString.Replace("\\\"", "\"")
                            .Replace("\\\\", "\\")
                            .Replace("\\n", "\n")
                            .Replace("\\r", "\r")
                            .Replace("\\t", "\t")
                            .Replace("\\/", "/");
        }


        internal string updateE3DocumentationContent(string json, string jspFile, string filePath, string fileName)
        {
            String urlpath = location + "/" + jspFile;

            var content = new MultipartFormDataContent();

            var jsonContent = new StringContent(json, Encoding.UTF8, "application/json");
            content.Add(jsonContent, "attributes");
            byte[] data;
            using (FileStream fstream = File.OpenRead(filePath + "/" + fileName))
            {
                data = new byte[fstream.Length];
                fstream.Read(data, 0, data.Length);
            }
            var bytes = new ByteArrayContent(data);
            content.Add(bytes, "file", fileName);

            Task<HttpResponseMessage> taskResponse = client.PostAsync(urlpath, content);
            taskResponse.Wait();
            taskResponse.Result.EnsureSuccessStatusCode();
            Task<byte[]> taskByteArray = taskResponse.Result.Content.ReadAsByteArrayAsync();
            taskByteArray.Wait();
            var responseString = Encoding.UTF8.GetString(taskByteArray.Result, 0, taskByteArray.Result.Length);
            return responseString;
        }

        internal string updateE3PrjDocContent(string json, string jspFile, string filePath, string fileName)
        {
            return updateE3DocumentationContent(json, jspFile, filePath, fileName);
        }

        internal String downloadE3PrjDocContent(string json, string jspFile, string filePath, string fileName)
        {
            string localPath = "";
            String urlpath = location + "/" + jspFile;
            var data = new StringContent(json, Encoding.UTF8, "application/json");
            Task<HttpResponseMessage> taskResponse = client.PostAsync(urlpath, data);
            taskResponse.Wait();
            taskResponse.Result.EnsureSuccessStatusCode();

            // Анализ типа контента
            var contentType = taskResponse.Result.Content.Headers.ContentType?.MediaType;

            if (contentType == "application/json")
            {
                // Обработка JSON ответа с описанием ошибки в Windchill
                Task<byte[]> jsonResponse = taskResponse.Result.Content.ReadAsByteArrayAsync();
                jsonResponse.Wait();
                String responseContent = Encoding.UTF8.GetString(jsonResponse.Result, 0, jsonResponse.Result.Length);

                CheckForErrorInResponse(responseContent);
            }
            else if (contentType == "application/octet-stream")
            {
                Task<byte[]> taskByteArray = taskResponse.Result.Content.ReadAsByteArrayAsync();
                taskByteArray.Wait();
                localPath = Path.Combine(filePath, fileName);
                File.WriteAllBytes(localPath, taskByteArray.Result);
            }

            return localPath;
        }

        internal bool isAuthorization()
        {
            return client.DefaultRequestHeaders.Authorization != null;
        }

    }
}
