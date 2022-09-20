using System;
using System.Net;
using System.Text;

namespace EODAddIn.Utils
{
    /// <summary>
    /// Class for sending and receiving requests from third-party services
    /// </summary>
    public static class Response
    {
        /// <summary>
        /// Sending a POST request
        /// </summary>
        /// <param name="Url">Address</param>
        /// <param name="Data">Settings of request</param>
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
            //The encoding is specified depending on the encoding of the server response
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
        /// GET request to source
        /// </summary>
        /// <param name="Url">Source address</param>
        /// <param name="Data">Data</param>
        /// <returns></returns>
        /// <exception cref="APIException">Error code</exception>
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
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(Encoding.Unicode.GetString(qwe));

            try
            {
                req.UserAgent = Program.Program.ProgramName;
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
                if (ex.Message == "Unauthenticated" || ex.Message == "Forbidden")
                {
                    error.MessageToUser("Your Api key is incorrect or does not give access to the requested information.");
                }
                else
                {
                    error.Send();
                }
                throw new APIException(0, ex.Message);
            }

        }
    }
}
