using System;
using System.Net;
using System.Text;

namespace EODAddIn.Utils
{
    /// <summary>
    /// Класс для отправки и получения запросов от сторонних сервисов
    /// </summary>
    public static class Response
    {
        /// <summary>
        /// Отправка POST запроса
        /// </summary>
        /// <param name="Url">Адрес</param>
        /// <param name="Data">Параметры запроса</param>
        /// <returns></returns>
        public static string POST(string Url, string Data)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            WebRequest req = WebRequest.Create(Url);
            req.Method = "POST";
            req.Timeout = 100000;
            req.ContentType = "application/x-www-form-urlencoded";

            byte[] sentData = Encoding.GetEncoding("UTF-8").GetBytes(Data);
            req.ContentLength = sentData.Length;
            System.IO.Stream sendStream = req.GetRequestStream();
            sendStream.Write(sentData, 0, sentData.Length);
            sendStream.Close();
            WebResponse res = req.GetResponse();
            System.IO.Stream ReceiveStream = res.GetResponseStream();
            System.IO.StreamReader sr = new System.IO.StreamReader(ReceiveStream, Encoding.UTF8);
            //Кодировка указывается в зависимости от кодировки ответа сервера
            char[] read = new char[256];
            int count = sr.Read(read, 0, 256);
            string Out = string.Empty;
            while (count > 0)
            {
                string str = new string(read, 0, count);
                Out += str;
                count = sr.Read(read, 0, 256);
            }
            return Out;
        }

        /// <summary>
        /// GET запрос к ресурсу
        /// </summary>
        /// <param name="Url">Адрес ресурса</param>
        /// <param name="Data">Данные</param>
        /// <returns></returns>
        /// <exception cref="APIException">Код ошибки</exception>
        public static string GET(string Url, string Data = "")
        {
            byte[] qwe;
            if (Data == "")
            {
                qwe = Encoding.Unicode.GetBytes(Url);
            }
            else
            {
                qwe = Encoding.Unicode.GetBytes(Url + "?" + Data);
            }

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            WebRequest req = WebRequest.Create(Encoding.Unicode.GetString(qwe));

            try
            {
                WebResponse resp = req.GetResponse();
                System.IO.Stream stream = resp.GetResponseStream();
                System.IO.StreamReader sr = new System.IO.StreamReader(stream);
                string Out = sr.ReadToEnd();
                sr.Close();

                return Out;
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse httpResponse = (HttpWebResponse)ex.Response;
                    throw new APIException((int)httpResponse.StatusCode, ex.Message);
                }
                else
                {
                    throw new APIException(500, ex.Message);
                }

            }
            catch (Exception ex)
            {
                Program.ErrorReport error = new Program.ErrorReport(ex);
                error.Send();
                throw new APIException(0, ex.Message);
            }

        }
    }
}
