using EODAddIn.BL.Live;
using EODAddIn.Utils;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace EODAddIn.Forms
{
    public partial class LiveDownloaderDispatcher : Form
    {

        private readonly Bitmap _greenBmp = new Bitmap(Properties.Resources.greenStatus);
        private readonly Bitmap _redBmp = new Bitmap(Properties.Resources.redStatus);
        private readonly Bitmap _yellowBmp = new Bitmap(Properties.Resources.yellowStatus);
        private static readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);

        public LiveDownloaderDispatcher()
        {
            InitializeComponent();

            foreach (LiveDownloader downloader in LiveDownloaderManager.LiveDownloaders)
            {
                downloader.ActiveChanged += Downloader_ActiveChanged;
            }

            UpdateGrid();
        }

        private void Downloader_ActiveChanged(object sender, EventArgs e)
        {
            UpdateGrid();
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
                subReq.ActiveChanged += Downloader_ActiveChanged;
                LiveDownloaderManager.Add(subReq);
            }
            UpdateGrid();
        }

        private void UpdateGrid()
        {
            _semaphore.Wait();
            try
            {
                gridDownloaders.Rows.Clear();
                foreach (var downloader in LiveDownloaderManager.LiveDownloaders)
                {
                    Bitmap bmp;
                    if (downloader.IsActive == null)
                    {
                        bmp = _yellowBmp;
                    }
                    else
                    {
                        if (downloader.IsActive == true)
                        {
                            bmp = _greenBmp;
                        }
                        else
                        {
                            bmp = _redBmp;
                        }
                    }
                    gridDownloaders.Rows.Add(downloader.Name, downloader.GetTickers(), downloader.Interval, bmp);
                }
            }
            catch
            {


            }
            finally { _semaphore.Release(); }
        }

        private void Downloader_OnStatusChanged(object sender, EventArgs e)
        {
            var downloader = (LiveDownloader)sender;
            foreach (DataGridViewRow row in gridDownloaders.Rows)
            {
                if (row.Cells[0].Value?.ToString() == downloader.Name)
                {
                    if (downloader.IsActive == null)
                    {
                        row.Cells[3].Value = _yellowBmp;
                    }
                    else
                    {
                        if ((bool)downloader.IsActive)
                        {
                            row.Cells[3].Value = _greenBmp;
                        }
                        else
                        {
                            row.Cells[3].Value = _redBmp;
                        }
                    }
                }
            }
        }

        private void LiveDownloaderDispatcher_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (var item in LiveDownloaderManager.LiveDownloaders)
            {
                item.Save();
            }
        }

        private void StartAll_Click(object sender, EventArgs e)
        {
            var selectedLoaders = GetSelectedDownLoaders();
            foreach (var item in selectedLoaders)
            {
                item.Start();
            }
        }

        private void StopAll_Click(object sender, EventArgs e)
        {
            var selectedLoaders = GetSelectedDownLoaders();
            foreach (var item in selectedLoaders)
            {
                item.Stop();
            }
        }

        private void DeleteAll_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to delete selected Live Downloaders",
                "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (result != DialogResult.OK) return;

            var selectedLoaders = GetSelectedDownLoaders();    
            foreach (var item in selectedLoaders)
            {
                LiveDownloaderManager.Delete(item);
            }

            var inxToDelete = new List<int>();
            for (int i = gridDownloaders.Rows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = gridDownloaders.Rows[i];
                if (!row.Selected) continue;
                inxToDelete.Add(i);    
            }

            foreach (var item in inxToDelete)
            {
                gridDownloaders.Rows.RemoveAt(item);
            }
            
        }

        private List<LiveDownloader> GetSelectedDownLoaders()
        {
            List<LiveDownloader> downloaderList = new List<LiveDownloader>();
            foreach (DataGridViewRow row in gridDownloaders.Rows)
            {
                if (!row.Selected) continue;
                string downloaderName = row.Cells[0].Value.ToString();

                var down = LiveDownloaderManager.LiveDownloaders.FirstOrDefault(x => x.Name == downloaderName);
                if (down == null) continue;

                downloaderList.Add(down);
            }
            return downloaderList;
        }

        private void AddToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmGetLive frm = new FrmGetLive();
            frm.Show(new WinHwnd());
            frm.FormClosing += Frm_FormClosing;
        }
    }
}
