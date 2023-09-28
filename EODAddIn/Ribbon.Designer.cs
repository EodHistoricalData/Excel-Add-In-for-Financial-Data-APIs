
namespace EODAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required constructor variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Release all used resources
        /// </summary>
        /// <param name="disposing">true if the managed resource is to be deleted; else false.</param>
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
            this.btnGetIntradayHistoricalData = this.Factory.CreateRibbonButton();
            this.splitbtnFundamental = this.Factory.CreateRibbonSplitButton();
            this.btnFundamentalAllData = this.Factory.CreateRibbonButton();
            this.btnGetGeneral = this.Factory.CreateRibbonButton();
            this.btnGetHighlights = this.Factory.CreateRibbonButton();
            this.btnGetBalanceSheet = this.Factory.CreateRibbonButton();
            this.btnGetIncomeStatement = this.Factory.CreateRibbonButton();
            this.btnGetFlowCash = this.Factory.CreateRibbonButton();
            this.btnGetEarnings = this.Factory.CreateRibbonButton();
            this.BtnGetBulk = this.Factory.CreateRibbonButton();
            this.BtnGetEtf = this.Factory.CreateRibbonButton();
            this.BtnOptions = this.Factory.CreateRibbonButton();
            this.BtnBulkEod = this.Factory.CreateRibbonButton();
            this.btnTechnicals = this.Factory.CreateRibbonButton();
            this.menuScreener = this.Factory.CreateRibbonMenu();
            this.btnCreateScreener = this.Factory.CreateRibbonButton();
            this.btnGetScreenerFundamenat = this.Factory.CreateRibbonButton();
            this.btnGetSreenerHistorical = this.Factory.CreateRibbonButton();
            this.btnGetScreenerIntraday = this.Factory.CreateRibbonButton();
            this.menuAvailableDataFeeds = this.Factory.CreateRibbonMenu();
            this.btnListOfExchanges = this.Factory.CreateRibbonButton();
            this.btnListOfCRYPTOCurrencies = this.Factory.CreateRibbonButton();
            this.btnListOfFutures = this.Factory.CreateRibbonButton();
            this.btnListOfForexCurrencies = this.Factory.CreateRibbonButton();
            this.btnListOfIndices = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.lblRequest = this.Factory.CreateRibbonLabel();
            this.lblRequestLeft = this.Factory.CreateRibbonLabel();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.grpAbout = this.Factory.CreateRibbonGroup();
            this.btnSendIdea = this.Factory.CreateRibbonButton();
            this.btnCheckUpdate = this.Factory.CreateRibbonButton();
            this.btnErrorMessage = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabMain.SuspendLayout();
            this.grpMain.SuspendLayout();
            this.group1.SuspendLayout();
            this.grpAbout.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.Groups.Add(this.grpMain);
            this.tabMain.Groups.Add(this.group1);
            this.tabMain.Groups.Add(this.grpAbout);
            this.tabMain.Label = "EODHD APIs";
            this.tabMain.Name = "tabMain";
            // 
            // grpMain
            // 
            this.grpMain.Items.Add(this.btnGetHistorical);
            this.grpMain.Items.Add(this.btnGetIntradayHistoricalData);
            this.grpMain.Items.Add(this.splitbtnFundamental);
            this.grpMain.Items.Add(this.BtnGetEtf);
            this.grpMain.Items.Add(this.BtnOptions);
            this.grpMain.Items.Add(this.BtnBulkEod);
            this.grpMain.Items.Add(this.btnTechnicals);
            this.grpMain.Items.Add(this.menuScreener);
            this.grpMain.Items.Add(this.menuAvailableDataFeeds);
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
            // btnGetIntradayHistoricalData
            // 
            this.btnGetIntradayHistoricalData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGetIntradayHistoricalData.Image = ((System.Drawing.Image)(resources.GetObject("btnGetIntradayHistoricalData.Image")));
            this.btnGetIntradayHistoricalData.Label = "Get intraday historical data";
            this.btnGetIntradayHistoricalData.Name = "btnGetIntradayHistoricalData";
            this.btnGetIntradayHistoricalData.ShowImage = true;
            this.btnGetIntradayHistoricalData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetIntradayHistoricalData_Click);
            // 
            // splitbtnFundamental
            // 
            this.splitbtnFundamental.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitbtnFundamental.Image = ((System.Drawing.Image)(resources.GetObject("splitbtnFundamental.Image")));
            this.splitbtnFundamental.Items.Add(this.btnFundamentalAllData);
            this.splitbtnFundamental.Items.Add(this.btnGetGeneral);
            this.splitbtnFundamental.Items.Add(this.btnGetHighlights);
            this.splitbtnFundamental.Items.Add(this.btnGetBalanceSheet);
            this.splitbtnFundamental.Items.Add(this.btnGetIncomeStatement);
            this.splitbtnFundamental.Items.Add(this.btnGetFlowCash);
            this.splitbtnFundamental.Items.Add(this.btnGetEarnings);
            this.splitbtnFundamental.Items.Add(this.BtnGetBulk);
            this.splitbtnFundamental.Label = "Get fundamental data";
            this.splitbtnFundamental.Name = "splitbtnFundamental";
            this.splitbtnFundamental.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SplitbtnFundamental_Click);
            // 
            // btnFundamentalAllData
            // 
            this.btnFundamentalAllData.Label = "All Data";
            this.btnFundamentalAllData.Name = "btnFundamentalAllData";
            this.btnFundamentalAllData.ShowImage = true;
            this.btnFundamentalAllData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFundamentalAllData_Click);
            // 
            // btnGetGeneral
            // 
            this.btnGetGeneral.Label = "General";
            this.btnGetGeneral.Name = "btnGetGeneral";
            this.btnGetGeneral.ShowImage = true;
            this.btnGetGeneral.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetGeneral_Click);
            // 
            // btnGetHighlights
            // 
            this.btnGetHighlights.Label = "Highlights";
            this.btnGetHighlights.Name = "btnGetHighlights";
            this.btnGetHighlights.ShowImage = true;
            this.btnGetHighlights.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetHighlights_Click);
            // 
            // btnGetBalanceSheet
            // 
            this.btnGetBalanceSheet.Label = "Balance Sheet";
            this.btnGetBalanceSheet.Name = "btnGetBalanceSheet";
            this.btnGetBalanceSheet.ShowImage = true;
            this.btnGetBalanceSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetBalanceSheet_Click);
            // 
            // btnGetIncomeStatement
            // 
            this.btnGetIncomeStatement.Label = "Income Statement";
            this.btnGetIncomeStatement.Name = "btnGetIncomeStatement";
            this.btnGetIncomeStatement.ShowImage = true;
            this.btnGetIncomeStatement.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetIncomeStatement_Click);
            // 
            // btnGetFlowCash
            // 
            this.btnGetFlowCash.Label = "FlowCash";
            this.btnGetFlowCash.Name = "btnGetFlowCash";
            this.btnGetFlowCash.ShowImage = true;
            this.btnGetFlowCash.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetCashFlow_Click);
            // 
            // btnGetEarnings
            // 
            this.btnGetEarnings.Label = "Earnings";
            this.btnGetEarnings.Name = "btnGetEarnings";
            this.btnGetEarnings.ShowImage = true;
            this.btnGetEarnings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetEarnings_Click);
            // 
            // BtnGetBulk
            // 
            this.BtnGetBulk.Image = ((System.Drawing.Image)(resources.GetObject("BtnGetBulk.Image")));
            this.BtnGetBulk.Label = "Bulk Fundamentals";
            this.BtnGetBulk.Name = "BtnGetBulk";
            this.BtnGetBulk.ShowImage = true;
            this.BtnGetBulk.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetBulk_Click);
            // 
            // BtnGetEtf
            // 
            this.BtnGetEtf.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnGetEtf.Image = ((System.Drawing.Image)(resources.GetObject("BtnGetEtf.Image")));
            this.BtnGetEtf.Label = "Get ETF";
            this.BtnGetEtf.Name = "BtnGetEtf";
            this.BtnGetEtf.ShowImage = true;
            this.BtnGetEtf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetEtf_Click);
            // 
            // BtnOptions
            // 
            this.BtnOptions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnOptions.Image = ((System.Drawing.Image)(resources.GetObject("BtnOptions.Image")));
            this.BtnOptions.Label = "Get Options";
            this.BtnOptions.Name = "BtnOptions";
            this.BtnOptions.ShowImage = true;
            this.BtnOptions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnOptions_Click);
            // 
            // BtnBulkEod
            // 
            this.BtnBulkEod.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnBulkEod.Image = global::EODAddIn.Properties.Resources.button_80;
            this.BtnBulkEod.Label = "Get Bulk EOD data";
            this.BtnBulkEod.Name = "BtnBulkEod";
            this.BtnBulkEod.ShowImage = true;
            this.BtnBulkEod.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnBulkEod_Click);
            // 
            // btnTechnicals
            // 
            this.btnTechnicals.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTechnicals.Image = ((System.Drawing.Image)(resources.GetObject("btnTechnicals.Image")));
            this.btnTechnicals.Label = "Get Technicals";
            this.btnTechnicals.Name = "btnTechnicals";
            this.btnTechnicals.ShowImage = true;
            this.btnTechnicals.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTechnicals_Click);
            // 
            // menuScreener
            // 
            this.menuScreener.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuScreener.Image = ((System.Drawing.Image)(resources.GetObject("menuScreener.Image")));
            this.menuScreener.Items.Add(this.btnCreateScreener);
            this.menuScreener.Items.Add(this.btnGetScreenerFundamenat);
            this.menuScreener.Items.Add(this.btnGetSreenerHistorical);
            this.menuScreener.Items.Add(this.btnGetScreenerIntraday);
            this.menuScreener.Label = "Stock Screener";
            this.menuScreener.Name = "menuScreener";
            this.menuScreener.ShowImage = true;
            // 
            // btnCreateScreener
            // 
            this.btnCreateScreener.Label = "Create";
            this.btnCreateScreener.Name = "btnCreateScreener";
            this.btnCreateScreener.ShowImage = true;
            this.btnCreateScreener.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateScreener_Click);
            // 
            // btnGetScreenerFundamenat
            // 
            this.btnGetScreenerFundamenat.Label = "Get fundamental";
            this.btnGetScreenerFundamenat.Name = "btnGetScreenerFundamenat";
            this.btnGetScreenerFundamenat.ShowImage = true;
            this.btnGetScreenerFundamenat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetScreenerFundamental_Click);
            // 
            // btnGetSreenerHistorical
            // 
            this.btnGetSreenerHistorical.Label = "Get hictorical";
            this.btnGetSreenerHistorical.Name = "btnGetSreenerHistorical";
            this.btnGetSreenerHistorical.ShowImage = true;
            this.btnGetSreenerHistorical.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetSreenerHistorical_Click);
            // 
            // btnGetScreenerIntraday
            // 
            this.btnGetScreenerIntraday.Label = "Get intraday ";
            this.btnGetScreenerIntraday.Name = "btnGetScreenerIntraday";
            this.btnGetScreenerIntraday.ShowImage = true;
            this.btnGetScreenerIntraday.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // menuAvailableDataFeeds
            // 
            this.menuAvailableDataFeeds.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuAvailableDataFeeds.Image = ((System.Drawing.Image)(resources.GetObject("menuAvailableDataFeeds.Image")));
            this.menuAvailableDataFeeds.Items.Add(this.btnListOfExchanges);
            this.menuAvailableDataFeeds.Items.Add(this.btnListOfCRYPTOCurrencies);
            this.menuAvailableDataFeeds.Items.Add(this.btnListOfFutures);
            this.menuAvailableDataFeeds.Items.Add(this.btnListOfForexCurrencies);
            this.menuAvailableDataFeeds.Items.Add(this.btnListOfIndices);
            this.menuAvailableDataFeeds.Label = "Available Data Feeds";
            this.menuAvailableDataFeeds.Name = "menuAvailableDataFeeds";
            this.menuAvailableDataFeeds.ShowImage = true;
            // 
            // btnListOfExchanges
            // 
            this.btnListOfExchanges.Label = "List of Exchanges";
            this.btnListOfExchanges.Name = "btnListOfExchanges";
            this.btnListOfExchanges.ShowImage = true;
            this.btnListOfExchanges.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnListOfExchanges_Click);
            // 
            // btnListOfCRYPTOCurrencies
            // 
            this.btnListOfCRYPTOCurrencies.Label = "List of CRYPTO Currencies";
            this.btnListOfCRYPTOCurrencies.Name = "btnListOfCRYPTOCurrencies";
            this.btnListOfCRYPTOCurrencies.ShowImage = true;
            this.btnListOfCRYPTOCurrencies.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnListOfCRYPTOCurrencies_Click);
            // 
            // btnListOfFutures
            // 
            this.btnListOfFutures.Label = "List of Futures/Commodities";
            this.btnListOfFutures.Name = "btnListOfFutures";
            this.btnListOfFutures.ShowImage = true;
            this.btnListOfFutures.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnListOfFutures_Click);
            // 
            // btnListOfForexCurrencies
            // 
            this.btnListOfForexCurrencies.Label = "List of Forex Currencies";
            this.btnListOfForexCurrencies.Name = "btnListOfForexCurrencies";
            this.btnListOfForexCurrencies.ShowImage = true;
            this.btnListOfForexCurrencies.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnListOfForexCurrencies_Click);
            // 
            // btnListOfIndices
            // 
            this.btnListOfIndices.Label = "List of Indices";
            this.btnListOfIndices.Name = "btnListOfIndices";
            this.btnListOfIndices.ShowImage = true;
            this.btnListOfIndices.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnListOfIndices_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.label2);
            this.group1.Items.Add(this.label3);
            this.group1.Items.Add(this.label1);
            this.group1.Items.Add(this.lblRequest);
            this.group1.Items.Add(this.lblRequestLeft);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btnSettings);
            this.group1.Label = "Limits";
            this.group1.Name = "group1";
            // 
            // label2
            // 
            this.label2.Label = "Request   ";
            this.label2.Name = "label2";
            // 
            // label3
            // 
            this.label3.Label = "Left";
            this.label3.Name = "label3";
            // 
            // label1
            // 
            this.label1.Label = " ";
            this.label1.Name = "label1";
            // 
            // lblRequest
            // 
            this.lblRequest.Label = "-";
            this.lblRequest.Name = "lblRequest";
            // 
            // lblRequestLeft
            // 
            this.lblRequestLeft.Label = "-";
            this.lblRequestLeft.Name = "lblRequestLeft";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
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
            this.btnSendIdea.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SendIdea_Click);
            // 
            // btnCheckUpdate
            // 
            this.btnCheckUpdate.Image = ((System.Drawing.Image)(resources.GetObject("btnCheckUpdate.Image")));
            this.btnCheckUpdate.Label = "Сheck for updates";
            this.btnCheckUpdate.Name = "btnCheckUpdate";
            this.btnCheckUpdate.ShowImage = true;
            this.btnCheckUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckUpdate_Click);
            // 
            // btnErrorMessage
            // 
            this.btnErrorMessage.Image = ((System.Drawing.Image)(resources.GetObject("btnErrorMessage.Image")));
            this.btnErrorMessage.Label = "Error message";
            this.btnErrorMessage.Name = "btnErrorMessage";
            this.btnErrorMessage.ShowImage = true;
            this.btnErrorMessage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ErrorMessage_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Image = ((System.Drawing.Image)(resources.GetObject("btnAbout.Image")));
            this.btnAbout.Label = "About";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.ShowImage = true;
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAbout_Click);
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
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblRequestLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitbtnFundamental;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetGeneral;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetHighlights;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetBalanceSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetIncomeStatement;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetEarnings;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFundamentalAllData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetFlowCash;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetIntradayHistoricalData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnGetEtf;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuScreener;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateScreener;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnGetBulk;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetScreenerFundamenat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetSreenerHistorical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetScreenerIntraday;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAvailableDataFeeds;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnListOfExchanges;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnListOfCRYPTOCurrencies;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnListOfFutures;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnListOfForexCurrencies;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnListOfIndices;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnBulkEod;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTechnicals;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
