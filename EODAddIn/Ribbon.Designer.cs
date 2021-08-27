
namespace EODAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.tabMain = this.Factory.CreateRibbonTab();
            this.grpMain = this.Factory.CreateRibbonGroup();
            this.btnGetHistorical = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.grpAbout = this.Factory.CreateRibbonGroup();
            this.btnSendIdea = this.Factory.CreateRibbonButton();
            this.btnCheckUpdate = this.Factory.CreateRibbonButton();
            this.btnErrorMessage = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.btnGetFundamentalData = this.Factory.CreateRibbonButton();
            this.tabMain.SuspendLayout();
            this.grpMain.SuspendLayout();
            this.grpAbout.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.Groups.Add(this.grpMain);
            this.tabMain.Groups.Add(this.grpAbout);
            this.tabMain.Label = "EOD";
            this.tabMain.Name = "tabMain";
            // 
            // grpMain
            // 
            this.grpMain.Items.Add(this.btnGetHistorical);
            this.grpMain.Items.Add(this.btnGetFundamentalData);
            this.grpMain.Items.Add(this.btnSettings);
            this.grpMain.Label = "Commands";
            this.grpMain.Name = "grpMain";
            // 
            // btnGetHistorical
            // 
            this.btnGetHistorical.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGetHistorical.Image = ((System.Drawing.Image)(resources.GetObject("btnGetHistorical.Image")));
            this.btnGetHistorical.Label = "Get historical data";
            this.btnGetHistorical.Name = "btnGetHistorical";
            this.btnGetHistorical.ShowImage = true;
            this.btnGetHistorical.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetHistorical_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSettings.Image = ((System.Drawing.Image)(resources.GetObject("btnSettings.Image")));
            this.btnSettings.Label = "Settings";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.ShowImage = true;
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSettings_Click);
            // 
            // grpAbout
            // 
            this.grpAbout.Items.Add(this.btnSendIdea);
            this.grpAbout.Items.Add(this.btnCheckUpdate);
            this.grpAbout.Items.Add(this.btnErrorMessage);
            this.grpAbout.Items.Add(this.btnAbout);
            this.grpAbout.Label = "About";
            this.grpAbout.Name = "grpAbout";
            // 
            // btnSendIdea
            // 
            this.btnSendIdea.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSendIdea.Image = ((System.Drawing.Image)(resources.GetObject("btnSendIdea.Image")));
            this.btnSendIdea.Label = "Send an idea";
            this.btnSendIdea.Name = "btnSendIdea";
            this.btnSendIdea.ShowImage = true;
            // 
            // btnCheckUpdate
            // 
            this.btnCheckUpdate.Image = ((System.Drawing.Image)(resources.GetObject("btnCheckUpdate.Image")));
            this.btnCheckUpdate.Label = "Сheck for updates";
            this.btnCheckUpdate.Name = "btnCheckUpdate";
            this.btnCheckUpdate.ShowImage = true;
            // 
            // btnErrorMessage
            // 
            this.btnErrorMessage.Image = ((System.Drawing.Image)(resources.GetObject("btnErrorMessage.Image")));
            this.btnErrorMessage.Label = "Error message";
            this.btnErrorMessage.Name = "btnErrorMessage";
            this.btnErrorMessage.ShowImage = true;
            // 
            // btnAbout
            // 
            this.btnAbout.Image = ((System.Drawing.Image)(resources.GetObject("btnAbout.Image")));
            this.btnAbout.Label = "About";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.ShowImage = true;
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAbout_Click);
            // 
            // btnGetFundamentalData
            // 
            this.btnGetFundamentalData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGetFundamentalData.Image = ((System.Drawing.Image)(resources.GetObject("btnGetFundamentalData.Image")));
            this.btnGetFundamentalData.Label = "Get fundamental data";
            this.btnGetFundamentalData.Name = "btnGetFundamentalData";
            this.btnGetFundamentalData.ShowImage = true;
            this.btnGetFundamentalData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetFundamentalData_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabMain);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabMain.ResumeLayout(false);
            this.tabMain.PerformLayout();
            this.grpMain.ResumeLayout(false);
            this.grpMain.PerformLayout();
            this.grpAbout.ResumeLayout(false);
            this.grpAbout.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSendIdea;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnErrorMessage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetHistorical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetFundamentalData;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
