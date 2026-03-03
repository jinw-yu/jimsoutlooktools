namespace jimsoutlooktools
{
    partial class RibbonTools : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonTools()
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
            this.tabJimsOutlookTools = this.Factory.CreateRibbonTab();
            this.groupAttachments = this.Factory.CreateRibbonGroup();
            this.groupSync = this.Factory.CreateRibbonGroup();
            this.btnSaveAttachments = this.Factory.CreateRibbonButton();
            this.btnDownloadOnline = this.Factory.CreateRibbonButton();
            this.tabJimsOutlookTools.SuspendLayout();
            this.groupAttachments.SuspendLayout();
            this.groupSync.SuspendLayout();
            // 
            // tabJimsOutlookTools
            // 
            this.tabJimsOutlookTools.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabJimsOutlookTools.Groups.Add(this.groupAttachments);
            this.tabJimsOutlookTools.Groups.Add(this.groupSync);
            this.tabJimsOutlookTools.Label = "JTools";
            this.tabJimsOutlookTools.Name = "tabJimsOutlookTools";
            // 
            // groupAttachments
            // 
            this.groupAttachments.Items.Add(this.btnSaveAttachments);
            this.groupAttachments.Label = "附件管理";
            this.groupAttachments.Name = "groupAttachments";
            // 
            // groupSync
            // 
            this.groupSync.Items.Add(this.btnDownloadOnline);
            this.groupSync.Label = "数据同步";
            this.groupSync.Name = "groupSync";
            // 
            // btnSaveAttachments
            // 
            this.btnSaveAttachments.Label = "保存附件";
            this.btnSaveAttachments.Name = "btnSaveAttachments";
            this.btnSaveAttachments.OfficeImageId = "AttachFile";
            this.btnSaveAttachments.ScreenTip = "保存收件箱附件";
            this.btnSaveAttachments.ShowImage = true;
            this.btnSaveAttachments.SuperTip = "将收件箱中指定日期范围内的邮件附件保存到本地文件夹";
            this.btnSaveAttachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveAttachments_Click);
            // 
            // btnDownloadOnline
            // 
            this.btnDownloadOnline.Label = "下载联机";
            this.btnDownloadOnline.Name = "btnDownloadOnline";
            this.btnDownloadOnline.OfficeImageId = "Download";
            this.btnDownloadOnline.ScreenTip = "从联机存档同步到本地PST";
            this.btnDownloadOnline.ShowImage = true;
            this.btnDownloadOnline.SuperTip = "选择联机存档数据文件和本地PST文件，分析差异后同步邮件";
            this.btnDownloadOnline.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDownloadOnline_Click);
            // 
            // RibbonTools
            // 
            this.Name = "RibbonTools";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabJimsOutlookTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonTools_Load);
            this.tabJimsOutlookTools.ResumeLayout(false);
            this.tabJimsOutlookTools.PerformLayout();
            this.groupAttachments.ResumeLayout(false);
            this.groupAttachments.PerformLayout();
            this.groupSync.ResumeLayout(false);
            this.groupSync.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabJimsOutlookTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAttachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSync;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveAttachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDownloadOnline;
    }
}
