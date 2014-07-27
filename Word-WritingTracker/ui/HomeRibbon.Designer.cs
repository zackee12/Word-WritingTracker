﻿namespace Word_WritingTracker
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
            this.button1 = this.Factory.CreateRibbonButton();
            this.buttonExport = this.Factory.CreateRibbonButton();
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
            this.groupTracking.Items.Add(this.button1);
            this.groupTracking.Items.Add(this.buttonExport);
            this.groupTracking.Label = "Tracking";
            this.groupTracking.Name = "groupTracking";
            // 
            // checkBoxTrackMetrics
            // 
            this.checkBoxTrackMetrics.Label = "Track Metrics";
            this.checkBoxTrackMetrics.Name = "checkBoxTrackMetrics";
            // 
            // button1
            // 
            this.button1.Label = "Charts";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "Chart3DColumnChart";
            this.button1.ShowImage = true;
            // 
            // buttonExport
            // 
            this.buttonExport.Label = "Export";
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.OfficeImageId = "FileSaveAsExcelXlsx";
            this.buttonExport.ShowImage = true;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonExport;
    }

    partial class ThisRibbonCollection
    {
        internal HomeRibbon HomeRibbon
        {
            get { return this.GetRibbon<HomeRibbon>(); }
        }
    }
}