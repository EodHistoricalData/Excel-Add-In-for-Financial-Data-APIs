using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EODAddIn.Program
{
    public static class Program
    {
        /// <summary>
        /// Название программы
        /// </summary>
        internal const string ProgramName = "EOD Excel Plagin";
        internal const string CompanyName = "EODHistoricalData";   
        internal const string UrlCompany = "https://eodhistoricaldata.com";

        /// <summary>
        /// Папка пользователя
        /// </summary>
        public static string UserFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), CompanyName, ProgramName);

        /// <summary>
        /// Версия программы
        /// </summary>
        internal static Version Version { get; private set; }

        /// <summary>
        /// Ключ активации программы
        /// </summary>
        internal static string APIKey
        {
            get => Settings.SettingsFields.APIKey;
            private set
            {
                Settings.SettingsFields.APIKey = value;
                Settings.Save();
            }
        }

        /// <summary>
        /// Идентификатор компьютера пользователя
        /// </summary>
        internal static string UserHash
        {
            get => string.IsNullOrEmpty(_UserHash) ? GetHash() : _UserHash;
            private set => _UserHash = value;
        }
        private static string _UserHash = string.Empty;




        /// <summary>
        /// Получение хэшкода компьютера пользователя
        /// </summary>
        /// <returns></returns>
        private static string GetHash()
        {
            SHA1Managed sha1 = new SHA1Managed();
            string s = GetProcessorId() + GetMotherBoardID();
            var hash = sha1.ComputeHash(Encoding.UTF8.GetBytes(s));
            var sb = new StringBuilder(hash.Length * 2);

            foreach (byte b in hash)
            {
                sb.Append(b.ToString("x2"));
            }
            return sb.ToString();

            string GetProcessorId()
            {
                SelectQuery query = new SelectQuery("Win32_processor");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);

                string result = string.Empty;
                foreach (ManagementObject info in searcher.Get())
                {
                    result = info["processorId"].ToString().Trim();
                }
                return result;
            }

            string GetMotherBoardID()
            {
                SelectQuery query = new SelectQuery("Win32_BaseBoard");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);

                string result = string.Empty;
                foreach (ManagementObject info in searcher.Get())
                {
                    result = info["SerialNumber"].ToString().Trim();
                }
                return result;
            }
        }
    }
}
