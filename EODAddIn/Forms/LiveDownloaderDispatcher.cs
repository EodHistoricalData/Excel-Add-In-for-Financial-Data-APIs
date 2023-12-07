using EODAddIn.BL.Live;
using EODAddIn.Program;
using EODAddIn.Utils;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using static EODAddIn.Utils.ExcelUtils;

namespace EODAddIn.Forms
{
    public partial class LiveDownloaderDispatcher : Form
    {
        private Dictionary<LiveDownloader, CustomXMLPart> LiveDownloaders = new Dictionary<LiveDownloader, CustomXMLPart>();
        private Dictionary<LiveDownloader, CancellationTokenSource> CancellationTokenSources = new Dictionary<LiveDownloader, CancellationTokenSource>();
        delegate void Download();
        Download _download;

        readonly Bitmap Greenbmp = new Bitmap(Properties.Resources.greenStatus);
        readonly Bitmap Redbmp = new Bitmap(Properties.Resources.redStatus);
        readonly Bitmap Yellowbmp = new Bitmap(Properties.Resources.yellowStatus);

        delegate void StatusHandler(bool? status);
        event StatusHandler StatusChanged;

        public LiveDownloaderDispatcher()
        {
            InitializeComponent();

            UpdateDownloaders();
        }

        public LiveDownloaderDispatcher(Dictionary<LiveDownloader, CustomXMLPart> downloaders, Dictionary<LiveDownloader, CancellationTokenSource> tokens)
        {
            InitializeComponent();

            LoadDownloaders(downloaders, tokens);
        }

        private void LoadDownloaders(Dictionary<LiveDownloader, CustomXMLPart> downloaders, Dictionary<LiveDownloader, CancellationTokenSource> tokens)
        {
            LiveDownloaders = downloaders;
            CancellationTokenSources = tokens;
            UpdateGrid();
        }

        public Dictionary<LiveDownloader, CustomXMLPart> GetDownloaders()
        {
            return LiveDownloaders;
        }

        public Dictionary<LiveDownloader, CancellationTokenSource> GetTokens()
        {
            return CancellationTokenSources;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            FrmGetLive frm = new FrmGetLive();
            frm.Show(new WinHwnd());
            frm.FormClosing += Frm_FormClosing;
        }

        private void Frm_FormClosing(object sender, FormClosingEventArgs e)
        {
            var frm = sender as FrmGetLive;
            var subReq = frm.LiveDownloader;
            if (subReq != null)
            {
                XmlSerializer xsSubmit = new XmlSerializer(typeof(LiveDownloader));
                var xml = "";

                using (var sww = new StringWriter())
                {
                    using (XmlWriter writer = XmlWriter.Create(sww))
                    {
                        xsSubmit.Serialize(writer, subReq);
                        xml = sww.ToString();
                    }
                }
                AddXmlPart(xml);
                SaveWorkbook();
                UpdateDownloaders();
            }
        }

        private void UpdateDownloaders()
        {
            LiveDownloaders.Clear();
            var xml = GetXmlPart();
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(LiveDownloader));
            foreach (CustomXMLPart item in xml)
            {
                LiveDownloader liveDownloader = null;
                try
                {
                    using (TextReader reader = new StringReader(item.XML))
                    {
                        liveDownloader = xmlSerializer.Deserialize(reader) as LiveDownloader;
                    }
                }
                catch
                {

                }
                finally
                {
                    if (liveDownloader != null)
                    {
                        LiveDownloaders.Add(liveDownloader, item);
                    }
                }
            }
            UpdateGrid();
        }

        private void UpdateGrid()
        {
            gridDownloaders.Rows.Clear();
            foreach (var downloader in LiveDownloaders)
            {
                Bitmap bmp;
                if (downloader.Key.IsActive == null)
                {
                    bmp = Yellowbmp;
                }
                else
                {
                    if (downloader.Key.IsActive == true)
                    {
                        bmp = Greenbmp;
                    }
                    else
                    {
                        bmp = Redbmp;
                    }
                }
                int i = gridDownloaders.Rows.Add(downloader.Key.Name, downloader.Key.GetTickers(), bmp);
                //downloader.Key.OnStatusChanged += Downloader_OnStatusChanged;
            }
        }

        private void Downloader_OnStatusChanged(LiveDownloader sender)
        {
            foreach (DataGridViewRow row in gridDownloaders.Rows)
            {
                if (row.Cells[0].Value.ToString() == sender.Name)
                {
                    if (sender.IsActive == null)
                    {
                        row.Cells[2].Value = Yellowbmp;
                    }
                    else
                    {
                        if ((bool)sender.IsActive)
                        {
                            row.Cells[2].Value = Greenbmp;
                        }
                        else
                        {
                            row.Cells[2].Value = Redbmp;
                        }
                    }

                }
            }
        }

        private void GridDownloaders_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            // delete downloader click
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                e.RowIndex >= 0 && e.ColumnIndex == 5 && senderGrid.Rows[e.RowIndex].Cells[0].Value != null)
            {
                string downloaderName = senderGrid.Rows[e.RowIndex].Cells[0].Value.ToString();
                if (!string.IsNullOrEmpty(downloaderName))
                {
                    var downloader = LiveDownloaders.Keys.FirstOrDefault(x => x.Name == downloaderName);

                    if (CancellationTokenSources.ContainsKey(downloader))
                    {
                        CancellationTokenSources[downloader].Cancel();
                        CancellationTokenSources.Remove(downloader);
                    }
                    if (downloader != null)
                    {
                        LiveDownloaders[downloader].Delete();
                        LiveDownloaders.Remove(downloader);
                        senderGrid.Rows.RemoveAt(e.RowIndex);
                    }
                }
            }
            // start
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                e.RowIndex >= 0 && e.ColumnIndex == 3 && senderGrid.Rows[e.RowIndex].Cells[0].Value != null)
            {
                string downloaderName = senderGrid.Rows[e.RowIndex].Cells[0].Value.ToString();
                var downloader = LiveDownloaders.Keys.First(x => x.Name == downloaderName);
                if (downloader != null)
                {
                    CancellationTokenSource src = new CancellationTokenSource();
                    LiveDownloaders.First(x => x.Key == downloader).Key.IsActive = true;
                    senderGrid.Rows[e.RowIndex].Cells[2].Value = Greenbmp;
                    CancellationTokenSources.Add(downloader, src);

                    async void Live() => await downloader.RequestAndPrint(src.Token);
                    _download = Live;
                    _download.Invoke();
                }
            }
            // stop
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                e.RowIndex >= 0 && e.ColumnIndex == 4 && senderGrid.Rows[e.RowIndex].Cells[0].Value != null)
            {
                string downloaderName = senderGrid.Rows[e.RowIndex].Cells[0].Value.ToString();
                var downloader = LiveDownloaders.Keys.First(x => x.Name == downloaderName);
                if (downloader != null && CancellationTokenSources.ContainsKey(downloader))
                {
                    LiveDownloaders.First(x => x.Key == downloader).Key.IsActive = false;
                    senderGrid.Rows[e.RowIndex].Cells[2].Value = Redbmp;
                    CancellationTokenSources[downloader].Cancel();
                    CancellationTokenSources.Remove(downloader);
                }
            }
        }

        private void LiveDownloaderDispatcher_FormClosing(object sender, FormClosingEventArgs e)
        {
            List<string> downloaderNames = new List<string>();
            foreach (DataGridViewRow row in gridDownloaders.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    string downloaderName = row.Cells[0].Value.ToString();
                    downloaderNames.Add(downloaderName);
                }
            }
            Settings.SettingsFields.LiveDownloaderNames = downloaderNames;
            Settings.Save();
        }
    }
}
