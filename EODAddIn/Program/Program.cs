﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace EODAddIn.Program
{
    public static class Program
    {
        /// <summary>
        /// Program name
        /// </summary>
        internal const string ProgramName = "EOD Excel Add-In";
        internal const string CompanyName = "EODHistoricalData";
        
        internal const string UrlCompany = "https://eodhd.com?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin";
        internal const string UrlRegister = "https://eodhd.com/register?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin";
        internal const string UrlKey = "https://eodhd.com/cp/settings?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin";
        internal const string UrlPrice = "https://eodhd.com/pricing?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin"; 
        internal const string UrlUpdate = "https://eodhd.com/excel-plugin-updates.xml";

        /// <summary>
        /// User folder
        /// </summary>
        public static string UserFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), CompanyName, ProgramName);

        /// <summary>
        /// Program version
        /// </summary>
        internal static Version Version { get; private set; }


        /// <summary>
        /// Program Activation Key
        /// </summary>
        internal static string APIKey
        {
            get => Settings.Data.APIKey;
            private set
            {
                Settings.Data.APIKey = value;
                Settings.Save();
            }
        }

        /// <summary>
        /// User computer ID
        /// </summary>
        internal static string UserHash
        {
            get => string.IsNullOrEmpty(_UserHash) ? GetHash() : _UserHash;
            private set => _UserHash = value;
        }
        private static string _UserHash = string.Empty;

        static Program()
        {
            var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            Version = new Version() { Major = ver.Major, Minor = ver.Minor, Build = ver.Build, Revision = ver.Revision };
        }

        public static void SaveAPI(string api)
        {
            APIKey = api;
        }

        /// <summary>
        /// Getting the hashcode of the user's PC
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


        /// <summary>
        /// Check for updates
        /// </summary>
        /// <returns></returns>
        /// <exception cref="System.Net.WebException">Connection error</exception>
        private static bool CheckUpdate()
        {
            try
            {
                if (GetVersionNews()?.Count > 0) return true;
            }
            catch (System.Net.WebException ex)
            {
                throw ex;
            }

            return false;
        }

        /// <summary>
        /// Get update history
        /// </summary>
        /// <returns></returns>
        /// <exception cref="System.Net.WebException">Connection error</exception>
        private static List<Version> GetVersions()
        {
            List<Version> versions = new List<Version>();

            string response;
            try
            {
                response = Utils.Response.GET(UrlUpdate, "");
            }
            catch (System.Net.WebException ex)
            {
                throw ex;
            }

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(response);
            XmlElement xRoot = xmlDocument.DocumentElement;
            XmlNodeList versionsList = xRoot.SelectNodes("Version");

            foreach (XmlNode versionNode in versionsList)
            {
                Version version = new Version()
                {
                    Name = versionNode.Attributes["name"].Value,
                    Description = versionNode.SelectSingleNode("Description").InnerText,
                    Link = versionNode.SelectSingleNode("Link").InnerText
                };

                string[] versplit = versionNode.SelectSingleNode("Number").InnerText.Split('.');
                version.Major = int.Parse(versplit[0]);
                version.Minor = int.Parse(versplit[1]);
                version.Build = int.Parse(versplit[2]);
                version.Revision = int.Parse(versplit[3]);
                DateTime.TryParse(versionNode.SelectSingleNode("Date").InnerText, out version.Date);
                versions.Add(version);
            }
            return versions;
        }

        /// <summary>
        /// Check for updates and offer to update
        /// </summary>
        private static void DoYouWantUpdate()
        {
            try
            {
                if (CheckUpdate())
                {
                    if (MessageBox.Show($"We've found updates.\nDo you want to update the Add-In?",
                                        ProgramName,
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        new FormUpdateList().ShowDialog();
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Check for updates in a separate thread
        /// </summary>
        public static void Run()
        {
            System.Threading.Tasks.Task.Factory.StartNew(() =>
            {
                try
                {
                    System.Threading.Thread.Sleep(10000);
                    DoYouWantUpdate();
                }
                catch { }
            });
        }

        /// <summary>
        ///Get a list of recent changes
        /// </summary>
        /// <returns></returns>
        /// <exception cref="System.Net.WebException">Connection error</exception>
        internal static List<Version> GetVersionNews()
        {
            List<Version> versions;
            try
            {
                versions = GetVersions();
            }
            catch (System.Net.WebException ex)
            {
                throw ex;
            }

            var ver = Version;

            List<Version> versionsNew = (from i in versions
                                         where (i.Major > ver.Major) ||
                                               (i.Major == ver.Major && i.Minor > ver.Minor) ||
                                               (i.Major == ver.Major && i.Minor == ver.Minor && i.Build > ver.Build) ||
                                               (i.Major == ver.Major && i.Minor == ver.Minor && i.Build == ver.Build && i.Revision > ver.Revision)
                                         orderby i.Major descending, i.Minor descending, i.Build descending, i.Revision descending
                                         select i).ToList();
            return versionsNew;
        }

        /// <summary>
        /// Update check and open change form with option to trigger update
        /// </summary>
        public static void CheckUpdates()
        {
            try
            {
                if (CheckUpdate())
                {
                    if (MessageBox.Show("There is a newer version of the program.\nDo you want to update?", "Updates", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        new FormUpdateList().ShowDialog();
                    }
                }
                else
                {
                    MessageBox.Show("You are using the latest version of the program", "Updates", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Net.WebException ex)
            {
                MessageBox.Show($"Couldn't check for updates.\nStatus - {ex.Status}",
                    "Updates",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
