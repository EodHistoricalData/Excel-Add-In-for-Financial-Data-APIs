
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
            this.tabMain = this.Factory.CreateRibbonTab();
            this.grpMain = this.Factory.CreateRibbonGroup();
            this.grpAbout = this.Factory.CreateRibbonGroup();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.btnCheckUpdate = this.Factory.CreateRibbonButton();
            this.btnErrorMessage = this.Factory.CreateRibbonButton();
            this.btnSendIdea = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
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
            this.grpMain.Items.Add(this.btnSettings);
            this.grpMain.Label = "Commands";
            this.grpMain.Name = "grpMain";
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
            // btnAbout
            // 
            this.btnAbout.Label = "About";
            this.btnAbout.Name = "btnAbout";
            // 
            // btnCheckUpdate
            // 
            this.btnCheckUpdate.Label = "Сheck for updates";
            this.btnCheckUpdate.Name = "btnCheckUpdate";
            // 
            // btnErrorMessage
            // 
            this.btnErrorMessage.Label = "Error message";
            this.btnErrorMessage.Name = "btnErrorMessage";
            // 
            // btnSendIdea
            // 
            this.btnSendIdea.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSendIdea.Label = "Send an idea";
            this.btnSendIdea.Name = "btnSendIdea";
            this.btnSendIdea.ShowImage = true;
            // 
            // btnSettings
            // 
            this.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSettings.Label = "Settings";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.ShowImage = true;
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
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
