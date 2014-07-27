namespace Word_WritingTracker
{
    partial class HomeRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public HomeRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabHome = this.Factory.CreateRibbonTab();
            this.groupTracking = this.Factory.CreateRibbonGroup();
            this.checkBoxTrackMetrics = this.Factory.CreateRibbonCheckBox();
            this.buttonChart = this.Factory.CreateRibbonButton();
            this.buttonExport = this.Factory.CreateRibbonButton();
            this.buttonSettings = this.Factory.CreateRibbonButton();
            this.tabHome.SuspendLayout();
            this.groupTracking.SuspendLayout();
            // 
            // tabHome
            // 
            this.tabHome.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabHome.ControlId.OfficeId = "TabHome";
            this.tabHome.Groups.Add(this.groupTracking);
            this.tabHome.Label = "TabHome";
            this.tabHome.Name = "tabHome";
            // 
            // groupTracking
            // 
            this.groupTracking.Items.Add(this.checkBoxTrackMetrics);
            this.groupTracking.Items.Add(this.buttonExport);
            this.groupTracking.Items.Add(this.buttonSettings);
            this.groupTracking.Items.Add(this.buttonChart);
            this.groupTracking.Label = "Tracking";
            this.groupTracking.Name = "groupTracking";
            // 
            // checkBoxTrackMetrics
            // 
            this.checkBoxTrackMetrics.Label = "Track Metrics";
            this.checkBoxTrackMetrics.Name = "checkBoxTrackMetrics";
            this.checkBoxTrackMetrics.ScreenTip = "Track this document";
            this.checkBoxTrackMetrics.SuperTip = "Enables tracking of the current document\'s word count";
            // 
            // buttonChart
            // 
            this.buttonChart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonChart.Label = "Chart";
            this.buttonChart.Name = "buttonChart";
            this.buttonChart.OfficeImageId = "Chart3DColumnChart";
            this.buttonChart.ScreenTip = "Open Charts";
            this.buttonChart.ShowImage = true;
            this.buttonChart.SuperTip = "Opens configurable word count charts";
            // 
            // buttonExport
            // 
            this.buttonExport.Label = "Export";
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.OfficeImageId = "FileSaveAsExcelXlsx";
            this.buttonExport.ScreenTip = "Export Data";
            this.buttonExport.ShowImage = true;
            this.buttonExport.SuperTip = "Export SQL data to a formated Excel document";
            // 
            // buttonSettings
            // 
            this.buttonSettings.Label = "Settings";
            this.buttonSettings.Name = "buttonSettings";
            this.buttonSettings.OfficeImageId = "AdpStoredProcedureEditSql";
            this.buttonSettings.ScreenTip = "Project Settings";
            this.buttonSettings.ShowImage = true;
            this.buttonSettings.SuperTip = "Adjust project settings such as project name and location";
            // 
            // HomeRibbon
            // 
            this.Name = "HomeRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabHome);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.HomeRibbon_Load);
            this.tabHome.ResumeLayout(false);
            this.tabHome.PerformLayout();
            this.groupTracking.ResumeLayout(false);
            this.groupTracking.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabHome;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTracking;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxTrackMetrics;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonChart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSettings;
    }

    partial class ThisRibbonCollection
    {
        internal HomeRibbon HomeRibbon
        {
            get { return this.GetRibbon<HomeRibbon>(); }
        }
    }
}
