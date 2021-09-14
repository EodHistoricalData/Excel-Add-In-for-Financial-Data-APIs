using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;

namespace EODAddIn.Program
{
    /// <summary>
    /// Форма обновления программы
    /// </summary>
    internal partial class FormUpdateList : Form
    {
        private readonly List<Version> Versions;
        private string fileName;

        public FormUpdateList()
        {
            InitializeComponent();

            Versions = Program.GetVersionNews();

            foreach (var version in Versions)
            {
                int start = rtxtUpdatesList.TextLength;
                string header = $"Version: {version.Text}";
                rtxtUpdatesList.AppendText(header);
                rtxtUpdatesList.AppendText(version.Description);
                rtxtUpdatesList.AppendText("\r\n");
                rtxtUpdatesList.Select(start, header.Length);
                rtxtUpdatesList.SelectionFont = new Font(rtxtUpdatesList.SelectionFont, FontStyle.Bold);
            }
            rtxtUpdatesList.Select(0, 0);
        }

        /// <summary>
        /// Кнопка отмены
        /// </summary>
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Кнопка обновления программы
        /// </summary>
        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                lblDownload.Visible = true;
                progbarDownloading.Visible = true;
                btnCancel.Enabled = false;
                btnUpdate.Enabled = false;
                DownLoadNewVersion();
            }
            catch (Exception ex)
            {
                new ErrorReport(ex).ShowAndSend();
            }

        }

        /// <summary>
        /// Скачивание новой версии программы
        /// </summary>
        private void DownLoadNewVersion()
        {
            string[] linksplit = Versions[0].Link.Split('/');

            fileName = Path.GetTempPath() + linksplit[linksplit.Length - 1];
            if (File.Exists(fileName)) File.Delete(fileName);

            WebClient client = new WebClient();
            client.DownloadProgressChanged += new DownloadProgressChangedEventHandler(DownloadProgressChanged);
            client.DownloadFileCompleted += new AsyncCompletedEventHandler(DownloadFileCompleted);
            client.DownloadFileAsync(new Uri(Versions[0].Link), fileName);
        }

        /// <summary>
        /// Завершение скачивания
        /// </summary>
        private void DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            Process.Start(fileName);
            Close();
            MessageBox.Show("Before starting the installation process, close all Excel files", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Процесс скачивания файла
        /// </summary>
        private void DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            progbarDownloading.Maximum = (int)e.TotalBytesToReceive / 100;
            progbarDownloading.Value = (int)e.BytesReceived / 100;
            Application.DoEvents();
        }
    }
}
