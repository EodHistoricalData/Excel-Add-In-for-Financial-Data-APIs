using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace EODAddIn.BL.Live
{
    public static class LiveDownloaderManager
    {
        public static IReadOnlyList<LiveDownloader> LiveDownloaders => _liveDownloaders.AsReadOnly();
        private static readonly List<LiveDownloader> _liveDownloaders  = new List<LiveDownloader>();

        public static void Add(LiveDownloader liveDownloader)
        {
            _liveDownloaders.Add(liveDownloader);
        }

        public static void Delete(LiveDownloader liveDownloader)
        {
            liveDownloader.Delete();
            _liveDownloaders.Remove(liveDownloader);
        }

        /// <summary>
        /// Load workbook downloaders
        /// </summary>
        /// <param name="Wb"></param>
        public static void LoadWorkbook(Workbook Wb)
        {
            var xml = Wb.CustomXMLParts;
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(LiveDownloader));

            if (xml == null) return;
            foreach (CustomXMLPart item in xml)
            {
                try
                {
                    using (TextReader reader = new StringReader(item.XML))
                    {
                        var liveDownloader = xmlSerializer.Deserialize(reader) as LiveDownloader;
                        liveDownloader.Set(Wb, item);
                        _liveDownloaders.Add(liveDownloader);
                    }
                }
                catch { }

            }
        }

        public static void CloseWorkbook(Workbook Wb)
        {
            var toDel = LiveDownloaders.Where(x => x.Workbook == Wb).ToList();

            for (int i = 0; i < toDel.Count(); i++)
            {
                _liveDownloaders.Remove(toDel[i]);
            }
        }
    }
}
