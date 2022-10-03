using Microsoft.Office.Tools;
using System.Diagnostics;
using System;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace EODAddIn.Panels
{
    /// <summary>
    /// Базовый класс панелей Excel
    /// </summary>
    public partial class PanelInfo : UserControl
    {
        public Excel.Workbook Workbook;

        /// <summary>
        /// Видимость панели
        /// </summary>
        public bool VisiblePanel
        {
            get
            {
                if (CustomTaskPanel != null)
                {
                    return CustomTaskPanel.Visible;
                }
                return false;
            }
            set
            {
                CustomTaskPanel.Visible = value;
            }
        }

        /// <summary>
        /// Заголовок
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Панель word 
        /// </summary>
        public CustomTaskPane CustomTaskPanel;

        public PanelInfo()
        {
            InitializeComponent();
            this.Dock = DockStyle.Fill;
            Title = "Info";
            CustomTaskPanel = Globals.ThisAddIn.CustomTaskPanes.Add(this, Title);
            
        }

        /// <summary>
        /// Отобразить панель
        /// </summary>
        public void ShowPanel()
        {
            if (CustomTaskPanel != null)
            {
                CustomTaskPanel.Visible = true;
                CustomTaskPanel.Width = 404;
            }
        }

        /// <summary>
        /// Скрыть панель
        /// </summary>
        public void HidePanel()
        {
            CustomTaskPanel.Visible = false;
        }

        //System.Diagnostics.Process.Start("https://eodhistoricaldata.com/financial-apis/list-supported-exchanges/?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");
        private void btnRegister_Click(object sender, System.EventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhistoricaldata.com/register?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");
        }

        private void btnLogin_Click(object sender, System.EventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhistoricaldata.com/login?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");
        }

        private void btnAPIkey_Click(object sender, System.EventArgs e)
        {
            Program.FrmAPIKey frm = new Program.FrmAPIKey();
            frm.ShowDialog();
        }

        private void btnUpgragePackages_Click(object sender, System.EventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhistoricaldata.com/pricing?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");
        }

        private void btnDocumentation_Click(object sender, System.EventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhistoricaldata.com/financial-academy/ready-to-go-solution/excel-add-in/");
        }

        private void btnPrivacyPolicy_Click(object sender, System.EventArgs e)
        {
            System.Diagnostics.Process.Start("https://eodhistoricaldata.com/financial-apis/privacy-policy?utm_source=p_c&utm_medium=excel&utm_campaign=exceladdin");
        }

        private void btnSendAnIdea_Click(object sender, System.EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("mailto:support@eodhistoricaldata.com" +
                          "?subject=Proposal for Excel AddIn ver. " + Program.Program.Version.Text);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void btnErrorReport_Click(object sender, System.EventArgs e)
        {
            try
            {
                Process.Start("mailto:support@eodhistoricaldata.com" +
                                          "?subject=Error in Excel AddIn ver. " + Program.Program.Version.Text);
            }
            catch (Exception ex)
            {
                Program.ErrorReport errorReport = new Program.ErrorReport(ex);
                errorReport.ShowAndSend();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            CustomTaskPanel.Visible = false;
        }

    }
}
