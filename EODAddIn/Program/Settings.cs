using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EODAddIn.Program
{
    internal static class Settings
    {
        private static readonly string xmlFilename = "settings.xml";
        private static readonly string path;
        internal static SettingsFields SettingsFields;

        static Settings()
        {
            SettingsFields = new SettingsFields();
            path = Path.Combine(Program.UserFolder, xmlFilename);
            Read();
        }

        /// <summary>
        /// Чтение настроек
        /// </summary>
        internal static void Read()
        {
            if (!File.Exists(path)) Save();

            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(SettingsFields));
                using (FileStream fs = File.OpenRead(path))
                {
                    SettingsFields = (SettingsFields)formatter.Deserialize(fs);
                }
            }
            catch (Exception)
            {
                Save();
            }
        }

        /// <summary>
        /// Сохранение настроек
        /// </summary>
        internal static void Save()
        {
            try
            {
                if (!Directory.Exists(Program.UserFolder)) Directory.CreateDirectory(Program.UserFolder);
                XmlSerializer formatter = new XmlSerializer(typeof(SettingsFields));
                using (FileStream fs = new FileStream(path, FileMode.Create))
                {
                    formatter.Serialize(fs, SettingsFields);
                }
            }
            catch (Exception ex)
            {
                new ErrorReport(ex).Send();
            }
        }
    }
}
