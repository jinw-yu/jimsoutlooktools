using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace jtools_outlook
{
    public partial class ThisAddIn
    {
        private const string AppVersion = "v1.1.0";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 备注: Outlook不会再触发这个事件。如果具有
            //    在 Outlook 关闭时必须运行，详请参阅 https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    public class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Button btnCancel;

        public bool IsCancelled { get; private set; }

        public ProgressForm()
        {
            IsCancelled = false;
            this.Text = "保存进度";
            this.Width = 550;
            this.Height = 180;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(15)
            };

            // 状态标签
            lblStatus = new Label
            {
                Text = "准备保存...",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10),
                Height = 25
            };

            // 进度条
            progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = 100,
                Height = 25
            };

            // 取消按钮
            btnCancel = new Button
            {
                Text = "停止保存",
                Width = 100,
                Height = 30,
                Anchor = AnchorStyles.None
            };
            btnCancel.Click += (s, e) =>
            {
                IsCancelled = true;
                lblStatus.Text = "正在停止，请稍候...";
                btnCancel.Enabled = false;
            };

            var buttonPanel = new Panel { Height = 50 };
            buttonPanel.Controls.Add(btnCancel);
            btnCancel.Left = (buttonPanel.Width - btnCancel.Width) / 2;
            btnCancel.Top = 10;

            tableLayout.Controls.Add(lblStatus, 0, 0);
            tableLayout.Controls.Add(progressBar, 0, 1);
            tableLayout.Controls.Add(buttonPanel, 0, 2);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            this.Controls.Add(tableLayout);
        }

        public void SetProgress(int current, int total)
        {
            if (total > 0)
            {
                progressBar.Maximum = total;
                progressBar.Value = Math.Min(current, total);
            }
            else
            {
                progressBar.Maximum = 1;
                progressBar.Value = 0;
            }
        }

        public void IncrementProgress()
        {
            if (progressBar.Value < progressBar.Maximum)
            {
                progressBar.Value++;
            }
        }

        public void UpdateStatus(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<string>(UpdateStatus), message);
                return;
            }
            lblStatus.Text = message;
        }
    }

    public class DateRangePickerForm : Form
    {
        private DateTimePicker startDatePicker;
        private DateTimePicker endDatePicker;
        private TextBox pathTextBox;
        private Button browseButton;
        private Button okButton;
        private Button cancelButton;
        private CheckBox chkInbox;
        private CheckBox chkSentItems;

        public DateTime StartDate { get; private set; }
        public DateTime EndDate { get; private set; }
        public string SavePath { get; private set; }
        public bool SaveInbox { get; private set; }
        public bool SaveSentItems { get; private set; }

        public DateRangePickerForm()
        {
            this.Text = "保存邮件附件";
            this.Width = 480;
            this.Height = 480;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 主容器
            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 11,
                Padding = new Padding(20)
            };



            // 保存路径标签
            var pathLabel = new Label
            {
                Text = "保存路径：",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            // 路径选择面板
            var pathPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 30
            };

            pathTextBox = new TextBox
            {
                Left = 0,
                Top = 2,
                Width = 330,
                Height = 25,
                ReadOnly = true
            };

            browseButton = new Button
            {
                Text = "浏览...",
                Left = 340,
                Top = 0,
                Width = 70,
                Height = 28
            };
            browseButton.Click += BrowseButton_Click;

            pathPanel.Controls.Add(pathTextBox);
            pathPanel.Controls.Add(browseButton);

            // 起始日期标签
            var startLabel = new Label
            {
                Text = "起始日期：",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            // 起始日期选择器
            startDatePicker = new DateTimePicker
            {
                Dock = DockStyle.Fill,
                Format = DateTimePickerFormat.Short,
                Height = 25
            };

            // 分隔
            var spacer = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                Height = 5
            };

            // 结束日期标签
            var endLabel = new Label
            {
                Text = "结束日期：",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            // 结束日期选择器
            endDatePicker = new DateTimePicker
            {
                Dock = DockStyle.Fill,
                Format = DateTimePickerFormat.Short,
                Height = 25
            };

            // 分隔
            var spacer2 = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                Height = 5
            };

            // 文件夹选择标签
            var folderLabel = new Label
            {
                Text = "选择要保存附件的文件夹：",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            // 收件箱复选框
            chkInbox = new CheckBox
            {
                Text = "收件箱",
                Dock = DockStyle.Fill,
                Height = 25,
                Checked = true
            };

            // 已发送邮件复选框
            chkSentItems = new CheckBox
            {
                Text = "已发送邮件",
                Dock = DockStyle.Fill,
                Height = 25,
                Checked = false
            };

            // 按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 40
            };

            okButton = new Button
            {
                Text = "确定",
                Width = 80,
                Height = 30,
                Left = 120,
                Top = 5
            };

            cancelButton = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 30,
                Left = 220,
                Top = 5
            };

            okButton.Click += (sender, e) =>
            {
                if (string.IsNullOrWhiteSpace(pathTextBox.Text))
                {
                    MessageBox.Show("请先选择保存路径！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!chkInbox.Checked && !chkSentItems.Checked)
                {
                    MessageBox.Show("请至少选择一个文件夹！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                SavePath = pathTextBox.Text;
                StartDate = startDatePicker.Value;
                EndDate = endDatePicker.Value;
                SaveInbox = chkInbox.Checked;
                SaveSentItems = chkSentItems.Checked;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            cancelButton.Click += (sender, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };

            buttonPanel.Controls.Add(okButton);
            buttonPanel.Controls.Add(cancelButton);

            // 添加到布局
            tableLayout.Controls.Add(pathLabel, 0, 0);
            tableLayout.Controls.Add(pathPanel, 0, 1);
            tableLayout.Controls.Add(startLabel, 0, 2);
            tableLayout.Controls.Add(startDatePicker, 0, 3);
            tableLayout.Controls.Add(spacer, 0, 4);
            tableLayout.Controls.Add(endLabel, 0, 5);
            tableLayout.Controls.Add(endDatePicker, 0, 6);
            tableLayout.Controls.Add(spacer2, 0, 7);
            tableLayout.Controls.Add(folderLabel, 0, 8);
            tableLayout.Controls.Add(chkInbox, 0, 9);
            tableLayout.Controls.Add(chkSentItems, 0, 10);
            tableLayout.Controls.Add(buttonPanel, 0, 11);

            // 设置行高 - 增加间距避免重叠
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));  // 保存路径标签
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40));  // 路径选择面板
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));  // 起始日期标签
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40));  // 起始日期选择器
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 15));  // 分隔
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));  // 结束日期标签
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40));  // 结束日期选择器
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20));  // 分隔
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));  // 文件夹选择标签
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30));  // 收件箱复选框
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30));  // 已发送邮件复选框
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 55));  // 按钮面板

            this.Controls.Add(tableLayout);
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "请选择附件保存的根文件夹";
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    pathTextBox.Text = folderDialog.SelectedPath;
                }
            }
        }
    }

    public class SaveResultForm : Form
    {
        public SaveResultForm(string resultText)
        {
            this.Text = "保存结果详情";
            this.Width = 700;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimizeBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(10)
            };

            // 标题
            var titleLabel = new Label
            {
                Text = "保存结果详情",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 30
            };

            // 文本框显示结果
            var textBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Text = resultText,
                Font = new System.Drawing.Font("Consolas", 9),
                BackColor = System.Drawing.Color.White
            };

            // 按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };

            var okButton = new Button
            {
                Text = "确定",
                Width = 80,
                Height = 30,
                Left = 300,
                Top = 5,
                DialogResult = DialogResult.OK
            };
            buttonPanel.Controls.Add(okButton);

            // 复制按钮
            var copyButton = new Button
            {
                Text = "复制到剪贴板",
                Width = 100,
                Height = 30,
                Left = 180,
                Top = 5
            };
            copyButton.Click += (s, e) =>
            {
                Clipboard.SetText(resultText);
                MessageBox.Show("已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            buttonPanel.Controls.Add(copyButton);

            tableLayout.Controls.Add(titleLabel, 0, 0);
            tableLayout.Controls.Add(textBox, 0, 1);
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            this.Controls.Add(tableLayout);
            this.Controls.Add(buttonPanel);

            this.AcceptButton = okButton;
        }
    }

    #region 联机存档同步窗体

    /// <summary>
    /// 数据文件信息
    /// </summary>
    public class StoreInfo
    {
        public Outlook.Store Store { get; set; }
        public string DisplayName { get; set; }
        public bool IsArchive { get; set; }
    }

    /// <summary>
    /// 关于对话框
    /// </summary>
    public class AboutForm : Form
    {
        public AboutForm()
        {
            this.Text = "关于 JTools-outlook";
            this.Width = 450;
            this.Height = 380;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(25)
            };

            // 应用名称
            var lblTitle = new Label
            {
                Text = "JTools-outlook",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 20, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 45
            };

            // 版本号
            var lblVersion = new Label
            {
                Text = "版本 v1.1.0",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 11),
                Height = 25
            };

            // 分隔线
            var separator = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                Height = 2,
                BorderStyle = BorderStyle.Fixed3D
            };

            // 描述
            var lblDescription = new Label
            {
                Text = "Outlook功能增强工具",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                Height = 50
            };

            // 版权信息
            var lblCopyright = new Label
            {
                Text = "Copyright © 2025 Jim\n基于 MIT 协议开源",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8),
                ForeColor = System.Drawing.Color.Gray,
                Height = 40
            };

            // 按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 45
            };

            var btnOK = new Button
            {
                Text = "确定",
                Width = 80,
                Height = 30,
                Left = 160,
                Top = 8,
                DialogResult = DialogResult.OK
            };
            buttonPanel.Controls.Add(btnOK);

            tableLayout.Controls.Add(lblTitle, 0, 0);
            tableLayout.Controls.Add(lblVersion, 0, 1);
            tableLayout.Controls.Add(separator, 0, 2);
            tableLayout.Controls.Add(lblDescription, 0, 3);
            tableLayout.Controls.Add(lblCopyright, 0, 4);
            tableLayout.Controls.Add(buttonPanel, 0, 5);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 15));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 55));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            this.Controls.Add(tableLayout);
            this.AcceptButton = btnOK;
        }
    }

    /// <summary>
    /// 年度邮件统计信息
    /// </summary>
    public class YearlyMailStats
    {
        public int Year { get; set; }
        public int InboxCount { get; set; }
        public int SentCount { get; set; }
        public Outlook.MAPIFolder InboxFolder { get; set; }
        public Outlook.MAPIFolder SentFolder { get; set; }
    }

    /// <summary>
    /// 年度邮件统计项（用于显示）
    /// </summary>
    public class YearlyStatsItem
    {
        public int Year { get; set; }
        public int OnlineInboxCount { get; set; }
        public int OnlineSentCount { get; set; }
        public int LocalInboxCount { get; set; }
        public int LocalSentCount { get; set; }
        public YearlyMailStats OnlineStats { get; set; }
        public YearlyMailStats LocalStats { get; set; }
    }

    /// <summary>
    /// 下载联机向导窗口 - 左右分栏布局
    /// </summary>
    public class DownloadOnlineWizardForm : Form
    {
        private const int MAX_LOG_LINES = 1000;

        // 向导步骤
        private enum WizardStep
        {
            SelectDataFiles = 0,    // 选择数据文件
            Analyzing = 1,          // 分析差异
            Syncing = 2             // 同步中
        }

        private WizardStep currentStep = WizardStep.SelectDataFiles;

        // 主布局控件
        private Panel mainPanel;
        private Panel topPanel;
        private Panel bottomPanel;
        private Label lblStepTitle;
        private Button btnNext;
        private Button btnCancel;

        // 步骤1：选择数据文件
        private ComboBox cmbSourceStore;
        private ComboBox cmbTargetStore;
        private ProgressBar progressLoading;
        private Label lblLoading;

        // 步骤2：左右分栏显示
        private TableLayoutPanel statsLayout;
        private ListView lvOnlineStats;      // 左侧：联机数据
        private ListView lvLocalStats;       // 右侧：本地数据
        private Label lblOnlineTitle;
        private Label lblLocalTitle;
        private Button btnSelectAll;
        private Button btnDeselectAll;
        private Label lblStatsSummary;

        // 步骤3：同步进度
        private ProgressBar progressSync;
        private Label lblSyncStatus;

        // 下部分：日志区域
        private TextBox txtLog;
        private Button btnCopyLog;

        // 数据
        private Outlook.Application _application;
        private List<StoreInfo> _sourceStores;
        private List<StoreInfo> _targetStores;
        private Outlook.Store _sourceStore;
        private Outlook.Store _targetStore;
        private List<YearlyStatsItem> _yearlyStats;
        private List<YearlyStatsItem> _selectedYears;

        // BackgroundWorker 用于后台同步
        private System.ComponentModel.BackgroundWorker _syncWorker;
        private volatile bool _isCancelled = false;
        public bool IsCancelled => _isCancelled;

        public void CancelSync()
        {
            _isCancelled = true;
            if (_syncWorker != null && _syncWorker.IsBusy)
            {
                _syncWorker.CancelAsync();
            }
        }

        public DownloadOnlineWizardForm(Outlook.Application application)
        {
            _application = application;
            _sourceStores = new List<StoreInfo>();
            _targetStores = new List<StoreInfo>();
            _yearlyStats = new List<YearlyStatsItem>();
            _selectedYears = new List<YearlyStatsItem>();
            _isCancelled = false;

            InitializeComponent();

            // 使用 Load 事件，确保窗口先显示出来再加载数据
            this.Load += (s, e) =>
            {
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 100;
                timer.Tick += (sender, args) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    LoadStoresAsync();
                };
                timer.Start();
            };
        }

        private void InitializeComponent()
        {
            this.Text = "下载联机存档";
            this.Width = 1000;
            this.Height = 750;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.ShowInTaskbar = true;

            // 主布局：上下分割
            mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            // 上部分：向导区域
            topPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle
            };

            // 步骤标题
            lblStepTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 40,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };

            // 导航按钮面板
            var navPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };

            btnNext = new Button
            {
                Text = "下一步",
                Width = 100,
                Height = 32,
                Left = 10,
                Top = 9,
                Enabled = false
            };
            btnNext.Click += BtnNext_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Width = 90,
                Height = 32,
                Left = 780,
                Top = 9
            };
            btnCancel.Click += (s, e) =>
            {
                if (currentStep == WizardStep.Syncing)
                {
                    var result = MessageBox.Show(
                        "确定要取消同步吗？\n\n已同步的邮件将保留，未同步的邮件将不会继续同步。",
                        "确认取消",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        CancelSync();  // 使用新方法，立即设置取消标志
                        AddLog("用户取消同步操作");
                        btnCancel.Enabled = false;  // 禁用按钮，防止重复点击
                        btnCancel.Text = "正在取消...";
                    }
                }
                else
                {
                    CancelSync();
                    this.Close();
                }
            };

            navPanel.Controls.Add(btnNext);
            navPanel.Controls.Add(btnCancel);

            topPanel.Controls.Add(lblStepTitle);
            topPanel.Controls.Add(navPanel);

            // 下部分：日志区域
            bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 200,
                BorderStyle = BorderStyle.FixedSingle
            };

            var logTitleLabel = new Label
            {
                Text = "日志",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold),
                Padding = new Padding(5, 0, 0, 0)
            };

            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new System.Drawing.Font("Consolas", 9),
                BackColor = System.Drawing.Color.WhiteSmoke,
                ReadOnly = true,
                Text = "=== 操作日志 ===\r\n",
                Margin = new Padding(5),
                Padding = new Padding(3, 5, 3, 5)  // 上下增加内边距，避免文字被遮挡
            };

            var logButtonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 35
            };

            btnCopyLog = new Button
            {
                Text = "复制日志",
                Width = 90,
                Height = 26,
                Left = 10,
                Top = 4
            };
            btnCopyLog.Click += (s, e) =>
            {
                try
                {
                    Clipboard.SetText(txtLog.Text);
                    MessageBox.Show("日志已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch { }
            };

            logButtonPanel.Controls.Add(btnCopyLog);
            bottomPanel.Controls.Add(logButtonPanel);
            bottomPanel.Controls.Add(txtLog);
            bottomPanel.Controls.Add(logTitleLabel);

            mainPanel.Controls.Add(topPanel);
            mainPanel.Controls.Add(bottomPanel);

            this.Controls.Add(mainPanel);

            ShowStep(WizardStep.SelectDataFiles);
        }

        private void ShowStep(WizardStep step)
        {
            currentStep = step;

            // 移除旧的内容面板
            var oldContent = topPanel.Controls.OfType<Panel>().FirstOrDefault(p => p.Name == "contentPanel");
            if (oldContent != null)
            {
                topPanel.Controls.Remove(oldContent);
                oldContent.Dispose();
            }

            // 移除全选/全不选按钮
            var navPanel = topPanel.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Bottom);
            if (navPanel != null)
            {
                var buttonsToRemove = navPanel.Controls.OfType<Button>()
                    .Where(b => b.Text == "全选" || b.Text == "全不选").ToList();
                foreach (var btn in buttonsToRemove)
                    navPanel.Controls.Remove(btn);
            }

            switch (step)
            {
                case WizardStep.SelectDataFiles:
                    ShowSelectDataFilesStep();
                    break;
                case WizardStep.Analyzing:
                    ShowYearlyStatsStep();
                    break;
                case WizardStep.Syncing:
                    ShowSyncingStep();
                    break;
            }

            // 更新按钮状态
            btnCancel.Enabled = true;
            btnCancel.Text = step == WizardStep.Syncing ? "取消同步" : "取消";

            AddLog($"[DEBUG] ShowStep completed");
        }

        private void ShowSelectDataFilesStep()
        {
            lblStepTitle.Text = "步骤 1/2: 选择数据文件";

            var contentPanel = new Panel
            {
                Name = "contentPanel",
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 10, 20, 20)
            };

            // 加载状态
            var loadingPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50
            };

            lblLoading = new Label
            {
                Text = "正在加载数据文件列表...",
                Dock = DockStyle.Top,
                Height = 20,
                ForeColor = System.Drawing.Color.SteelBlue
            };

            progressLoading = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 18,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };

            loadingPanel.Controls.Add(progressLoading);
            loadingPanel.Controls.Add(lblLoading);

            // 源数据文件
            var sourceLabel = new Label
            {
                Text = "源数据文件（联机存档）:",
                Dock = DockStyle.Top,
                Height = 25,
                Margin = new Padding(0, 10, 0, 0)
            };

            cmbSourceStore = new ComboBox
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 25,
                Enabled = false
            };
            cmbSourceStore.SelectedIndexChanged += (s, e) => CheckCanProceed();

            // 目标数据文件
            var targetLabel = new Label
            {
                Text = "目标数据文件（本地PST）:",
                Dock = DockStyle.Top,
                Height = 25,
                Margin = new Padding(0, 10, 0, 0)
            };

            cmbTargetStore = new ComboBox
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 25,
                Enabled = false
            };
            cmbTargetStore.SelectedIndexChanged += (s, e) => CheckCanProceed();

            // 提示
            var hintLabel = new Label
            {
                Text = "提示：源数据文件是Office 365联机存档，目标数据文件是本地PST文件。\n将统计收件箱和已发送邮件的邮件数量，按年份显示。",
                Dock = DockStyle.Bottom,
                Height = 60,
                ForeColor = System.Drawing.Color.Gray,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8),
                Padding = new Padding(5, 8, 5, 8)  // 上下增加内边距
            };

            contentPanel.Controls.Add(hintLabel);
            contentPanel.Controls.Add(cmbTargetStore);
            contentPanel.Controls.Add(targetLabel);
            contentPanel.Controls.Add(cmbSourceStore);
            contentPanel.Controls.Add(sourceLabel);
            contentPanel.Controls.Add(loadingPanel);

            topPanel.Controls.Add(contentPanel);

            btnNext.Enabled = false;
        }

        private void ShowYearlyStatsStep()
        {
            lblStepTitle.Text = "步骤 2/2: 选择要下载的年份";

            var contentPanel = new Panel
            {
                Name = "contentPanel",
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 60, 10, 10)  // 顶部留出60像素空白
            };

            // 创建左右分栏布局容器
            var tableContainer = new Panel
            {
                Dock = DockStyle.Top,
                Height = 380  // 固定高度，约15行数据
            };

            // 创建左右分栏布局
            statsLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(5)
            };
            statsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            statsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            statsLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));  // 标题行高度增加50%
            statsLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // 左侧标题：联机数据
            lblOnlineTitle = new Label
            {
                Text = "【联机存档】",
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 11, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.SteelBlue,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Margin = new Padding(0)
            };
            statsLayout.Controls.Add(lblOnlineTitle, 0, 0);

            // 右侧标题：本地数据
            lblLocalTitle = new Label
            {
                Text = "【本地PST】",
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 11, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.SeaGreen,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Margin = new Padding(0)
            };
            statsLayout.Controls.Add(lblLocalTitle, 1, 0);

            // 左侧：联机数据列表（带复选框）
            lvOnlineStats = new ListView
            {
                View = System.Windows.Forms.View.Details,
                FullRowSelect = true,
                GridLines = true,
                Dock = DockStyle.Fill,
                CheckBoxes = true
            };
            lvOnlineStats.Columns.Add("年份", 60, HorizontalAlignment.Center);
            lvOnlineStats.Columns.Add("收件箱", 80, HorizontalAlignment.Right);
            lvOnlineStats.Columns.Add("已发送", 80, HorizontalAlignment.Right);
            lvOnlineStats.Columns.Add("合计", 80, HorizontalAlignment.Right);
            lvOnlineStats.ItemChecked += (s, e) => UpdateStatsSummary();
            statsLayout.Controls.Add(lvOnlineStats, 0, 1);

            // 右侧：本地数据列表（不带复选框）
            lvLocalStats = new ListView
            {
                View = System.Windows.Forms.View.Details,
                FullRowSelect = true,
                GridLines = true,
                Dock = DockStyle.Fill,
                CheckBoxes = false
            };
            lvLocalStats.Columns.Add("年份", 60, HorizontalAlignment.Center);
            lvLocalStats.Columns.Add("收件箱", 80, HorizontalAlignment.Right);
            lvLocalStats.Columns.Add("已发送", 80, HorizontalAlignment.Right);
            lvLocalStats.Columns.Add("合计", 80, HorizontalAlignment.Right);
            statsLayout.Controls.Add(lvLocalStats, 1, 1);

            tableContainer.Controls.Add(statsLayout);
            contentPanel.Controls.Add(tableContainer);

            // 底部统计信息
            lblStatsSummary = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 30,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };
            contentPanel.Controls.Add(lblStatsSummary);

            topPanel.Controls.Add(contentPanel);

            // 添加全选/全不选按钮
            var navPanel = topPanel.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Bottom);
            if (navPanel != null)
            {
                btnSelectAll = new Button
                {
                    Text = "全选",
                    Width = 80,
                    Height = 32,
                    Left = 120,
                    Top = 9
                };
                btnSelectAll.Click += (s, e) =>
                {
                    foreach (ListViewItem item in lvOnlineStats.Items)
                        item.Checked = true;
                };

                btnDeselectAll = new Button
                {
                    Text = "全不选",
                    Width = 80,
                    Height = 32,
                    Left = 210,
                    Top = 9
                };
                btnDeselectAll.Click += (s, e) =>
                {
                    foreach (ListViewItem item in lvOnlineStats.Items)
                        item.Checked = false;
                };

                navPanel.Controls.Add(btnSelectAll);
                navPanel.Controls.Add(btnDeselectAll);
            }

            // 加载数据
            LoadYearlyStats();
        }

        private void ShowSyncingStep()
        {
            lblStepTitle.Text = "正在下载邮件...";

            var contentPanel = new Panel
            {
                Name = "contentPanel",
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 10, 20, 20)
            };

            lblSyncStatus = new Label
            {
                Text = "准备下载...",
                Dock = DockStyle.Top,
                Height = 30,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10)
            };

            progressSync = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 30,
                Minimum = 0,
                Maximum = 100
            };

            contentPanel.Controls.Add(progressSync);
            contentPanel.Controls.Add(lblSyncStatus);

            topPanel.Controls.Add(contentPanel);

            btnNext.Enabled = false;
        }

        private async void LoadStoresAsync()
        {
            AddLog("正在加载数据文件列表...");

            try
            {
                var storeList = await System.Threading.Tasks.Task.Run(() =>
                {
                    var stores = _application.Session.Stores;
                    var list = new List<StoreInfo>();

                    foreach (Outlook.Store store in stores)
                    {
                        try
                        {
                            string displayName = store.DisplayName;
                            bool isArchive = displayName.ToLower().Contains("archive") ||
                                             displayName.ToLower().Contains("联机") ||
                                             displayName.ToLower().Contains("online");

                            list.Add(new StoreInfo
                            {
                                Store = store,
                                DisplayName = $"{(isArchive ? "[联机] " : "[本地] ")}{displayName}",
                                IsArchive = isArchive
                            });
                        }
                        catch { }
                    }
                    return list;
                });

                _sourceStores = storeList.Where(s => s.IsArchive).ToList();
                _targetStores = storeList.Where(s => !s.IsArchive).ToList();

                cmbSourceStore.DataSource = _sourceStores;
                cmbSourceStore.DisplayMember = "DisplayName";
                cmbSourceStore.ValueMember = "Store";
                cmbSourceStore.Enabled = true;

                cmbTargetStore.DataSource = _targetStores;
                cmbTargetStore.DisplayMember = "DisplayName";
                cmbTargetStore.ValueMember = "Store";
                cmbTargetStore.Enabled = true;

                progressLoading.Visible = false;
                lblLoading.Text = $"已加载 {_sourceStores.Count} 个联机存档, {_targetStores.Count} 个本地PST";
                AddLog($"已加载 {_sourceStores.Count} 个联机存档, {_targetStores.Count} 个本地PST");

                CheckCanProceed();
            }
            catch (System.Exception ex)
            {
                AddLog($"✗ 加载数据文件失败: {ex.Message}");
                MessageBox.Show($"加载数据文件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CheckCanProceed()
        {
            btnNext.Enabled = (cmbSourceStore.SelectedItem != null && cmbTargetStore.SelectedItem != null);
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            switch (currentStep)
            {
                case WizardStep.SelectDataFiles:
                    StartAnalysis();
                    break;
                case WizardStep.Analyzing:
                    StartSync();
                    break;
            }
        }

        private async void StartAnalysis()
        {
            try
            {
                _sourceStore = ((StoreInfo)cmbSourceStore.SelectedItem).Store;
                _targetStore = ((StoreInfo)cmbTargetStore.SelectedItem).Store;

                AddLog($"开始分析邮件数据...");
                AddLog($"联机存档: {_sourceStore.DisplayName}");
                AddLog($"本地PST: {_targetStore.DisplayName}");

                // 切换到统计界面
                ShowStep(WizardStep.Analyzing);
                btnNext.Enabled = false;

                // 在后台线程执行分析
                var stats = await System.Threading.Tasks.Task.Run(() =>
                {
                    return AnalyzeYearlyStats(_sourceStore, _targetStore);
                });

                _yearlyStats = stats;

                // 显示统计结果
                this.Invoke(new System.Action(() =>
                {
                    DisplayYearlyStats();
                    btnNext.Enabled = _yearlyStats.Any(s => s.OnlineInboxCount > 0 || s.OnlineSentCount > 0);
                }));
            }
            catch (System.Exception ex)
            {
                AddLog($"✗ 分析失败: {ex.Message}");
                MessageBox.Show($"分析失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowStep(WizardStep.SelectDataFiles);
            }
        }

        private List<YearlyStatsItem> AnalyzeYearlyStats(Outlook.Store onlineStore, Outlook.Store localStore)
        {
            var result = new List<YearlyStatsItem>();

            try
            {
                // 更新标题显示数据文件名称
                string onlineName = onlineStore?.DisplayName ?? "联机存档";
                string localName = localStore?.DisplayName ?? "本地PST";
                this.Invoke(new System.Action(() =>
                {
                    if (lblOnlineTitle != null)
                        lblOnlineTitle.Text = $"【联机存档】{onlineName}";
                    if (lblLocalTitle != null)
                        lblLocalTitle.Text = $"【本地PST】{localName}";
                }));

                // 获取联机存档的收件箱和已发送文件夹
                this.Invoke(new System.Action(() => AddLog("正在获取联机存档文件夹...")));
                var onlineInbox = FindFolder(onlineStore, "收件箱", "Inbox");
                var onlineSent = FindFolder(onlineStore, "已发送邮件", "Sent Items", "已发送");

                // 获取本地PST的收件箱和已发送文件夹
                this.Invoke(new System.Action(() => AddLog("正在获取本地PST文件夹...")));
                var localInbox = FindFolder(localStore, "收件箱", "Inbox");
                var localSent = FindFolder(localStore, "已发送邮件", "Sent Items", "已发送");

                // 获取所有年份范围
                this.Invoke(new System.Action(() => AddLog("正在确定年份范围...")));
                var allYears = GetYearRange(onlineInbox, onlineSent, localInbox, localSent);
                
                if (allYears.Count == 0)
                {
                    this.Invoke(new System.Action(() => AddLog("未找到任何邮件")));
                    return result;
                }

                this.Invoke(new System.Action(() => AddLog($"发现邮件年份范围: {allYears.Min()} - {allYears.Max()}，共 {allYears.Count} 个年份")));

                // 按年份逐个统计
                foreach (var year in allYears.OrderByDescending(y => y))
                {
                    this.Invoke(new System.Action(() => AddLog($"正在统计 {year} 年...")));

                    int onlineInboxCount = onlineInbox != null ? CountMailsByYear(onlineInbox, year) : 0;
                    int onlineSentCount = onlineSent != null ? CountMailsByYear(onlineSent, year) : 0;
                    int localInboxCount = localInbox != null ? CountMailsByYear(localInbox, year) : 0;
                    int localSentCount = localSent != null ? CountMailsByYear(localSent, year) : 0;

                    var yearStats = new YearlyStatsItem
                    {
                        Year = year,
                        OnlineInboxCount = onlineInboxCount,
                        OnlineSentCount = onlineSentCount,
                        LocalInboxCount = localInboxCount,
                        LocalSentCount = localSentCount,
                        OnlineStats = new YearlyMailStats
                        {
                            Year = year,
                            InboxCount = onlineInboxCount,
                            SentCount = onlineSentCount,
                            InboxFolder = onlineInbox,
                            SentFolder = onlineSent
                        },
                        LocalStats = new YearlyMailStats
                        {
                            Year = year,
                            InboxCount = localInboxCount,
                            SentCount = localSentCount,
                            InboxFolder = localInbox,
                            SentFolder = localSent
                        }
                    };

                    result.Add(yearStats);

                    // 每统计完一年，刷新UI显示
                    this.Invoke(new System.Action(() =>
                    {
                        AddLog($"  {year}年: 联机收件箱 {onlineInboxCount}, 联机已发送 {onlineSentCount}, 本地收件箱 {localInboxCount}, 本地已发送 {localSentCount}");
                        // 实时更新显示
                        UpdateYearlyStatsDisplay(result);
                    }));

                    System.Threading.Thread.Sleep(500); // 暂停0.5秒让UI刷新
                }

                this.Invoke(new System.Action(() => AddLog($"分析完成，共 {result.Count} 个年份的数据")));
            }
            catch (System.Exception ex)
            {
                this.Invoke(new System.Action(() => AddLog($"分析出错: {ex.Message}")));
            }

            return result;
        }

        /// <summary>
        /// 获取所有邮件的年份范围
        /// </summary>
        private HashSet<int> GetYearRange(params Outlook.MAPIFolder[] folders)
        {
            var years = new HashSet<int>();

            foreach (var folder in folders)
            {
                if (folder == null) continue;

                try
                {
                    var items = folder.Items;
                    items.Sort("[ReceivedTime]", true);

                    // 获取第一封邮件的年份（最新的）
                    try
                    {
                        var firstItem = items.GetFirst();
                        if (firstItem != null)
                        {
                            DateTime? firstTime = GetItemTime(firstItem);
                            if (firstTime.HasValue)
                                years.Add(firstTime.Value.Year);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(firstItem);
                        }
                    }
                    catch { }

                    // 获取最后一封邮件的年份（最早的）
                    try
                    {
                        var lastItem = items.GetLast();
                        if (lastItem != null)
                        {
                            DateTime? lastTime = GetItemTime(lastItem);
                            if (lastTime.HasValue)
                                years.Add(lastTime.Value.Year);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(lastItem);
                        }
                    }
                    catch { }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                }
                catch { }
            }

            // 补充中间的年份
            if (years.Count > 0)
            {
                int minYear = years.Min();
                int maxYear = years.Max();
                years.Clear();
                for (int y = minYear; y <= maxYear; y++)
                {
                    years.Add(y);
                }
            }

            return years;
        }

        /// <summary>
        /// 获取邮件项目的时间
        /// </summary>
        private DateTime? GetItemTime(object item)
        {
            try
            {
                if (item is Outlook.MailItem mail)
                {
                    var time = mail.ReceivedTime;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                    return time;
                }
                else if (item is Outlook.MeetingItem meeting)
                {
                    var time = meeting.ReceivedTime;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(meeting);
                    return time;
                }
                else if (item is Outlook.ReportItem report)
                {
                    var time = report.CreationTime;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(report);
                    return time;
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// 统计指定年份的邮件数量
        /// </summary>
        private int CountMailsByYear(Outlook.MAPIFolder folder, int year)
        {
            int count = 0;

            try
            {
                var items = folder.Items;
                items.Sort("[ReceivedTime]", true);

                // 使用筛选获取指定年份的邮件
                string filter = $"[ReceivedTime] >= '{year}/1/1' AND [ReceivedTime] < '{year + 1}/1/1'";
                Outlook.Items filteredItems = null;

                try
                {
                    filteredItems = items.Restrict(filter);
                    count = filteredItems.Count;
                }
                catch
                {
                    // 如果筛选失败，遍历计数
                    foreach (object item in items)
                    {
                        try
                        {
                            DateTime? receivedTime = GetItemTime(item);
                            if (receivedTime.HasValue && receivedTime.Value.Year == year)
                            {
                                count++;
                            }
                        }
                        catch { }
                        finally
                        {
                            if (item != null)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                        }
                    }
                }

                if (filteredItems != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
            }
            catch { }

            return count;
        }

        /// <summary>
        /// 实时更新统计显示
        /// </summary>
        private void UpdateYearlyStatsDisplay(List<YearlyStatsItem> currentStats)
        {
            if (lvOnlineStats == null || lvLocalStats == null) return;

            lvOnlineStats.Items.Clear();
            lvLocalStats.Items.Clear();

            foreach (var stat in currentStats.OrderByDescending(s => s.Year))
            {
                // 左侧：联机数据（带复选框）
                var onlineItem = new ListViewItem(stat.Year.ToString());
                onlineItem.SubItems.Add(stat.OnlineInboxCount.ToString("N0"));
                onlineItem.SubItems.Add(stat.OnlineSentCount.ToString("N0"));
                onlineItem.SubItems.Add((stat.OnlineInboxCount + stat.OnlineSentCount).ToString("N0"));
                onlineItem.Checked = stat.OnlineInboxCount > 0 || stat.OnlineSentCount > 0;
                onlineItem.Tag = stat;
                lvOnlineStats.Items.Add(onlineItem);

                // 右侧：本地数据（不带复选框）
                var localItem = new ListViewItem(stat.Year.ToString());
                localItem.SubItems.Add(stat.LocalInboxCount.ToString("N0"));
                localItem.SubItems.Add(stat.LocalSentCount.ToString("N0"));
                localItem.SubItems.Add((stat.LocalInboxCount + stat.LocalSentCount).ToString("N0"));
                lvLocalStats.Items.Add(localItem);
            }

            UpdateStatsSummary();
        }

        private Outlook.MAPIFolder FindFolder(Outlook.Store store, params string[] possibleNames)
        {
            try
            {
                var root = store.GetRootFolder();
                foreach (Outlook.MAPIFolder folder in root.Folders)
                {
                    foreach (var name in possibleNames)
                    {
                        if (folder.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                        {
                            return folder;
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        private Dictionary<int, int> GetYearlyMailCount(Outlook.MAPIFolder folder)
        {
            var result = new Dictionary<int, int>();

            try
            {
                var items = folder.Items;
                items.Sort("[ReceivedTime]", true);

                foreach (object item in items)
                {
                    try
                    {
                        DateTime? receivedTime = null;

                        if (item is Outlook.MailItem mail)
                        {
                            receivedTime = mail.ReceivedTime;
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                        }
                        else if (item is Outlook.MeetingItem meeting)
                        {
                            receivedTime = meeting.ReceivedTime;
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(meeting);
                        }
                        else if (item is Outlook.ReportItem report)
                        {
                            receivedTime = report.CreationTime;
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(report);
                        }

                        if (receivedTime.HasValue)
                        {
                            int year = receivedTime.Value.Year;
                            if (!result.ContainsKey(year))
                                result[year] = 0;
                            result[year]++;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (item != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
            }
            catch { }

            return result;
        }

        private void LoadYearlyStats()
        {
            // 数据已在 StartAnalysis 中加载
            DisplayYearlyStats();
        }

        private void DisplayYearlyStats()
        {
            lvOnlineStats.Items.Clear();
            lvLocalStats.Items.Clear();

            foreach (var stat in _yearlyStats)
            {
                // 左侧：联机数据（带复选框）
                var onlineItem = new ListViewItem(stat.Year.ToString());
                onlineItem.SubItems.Add(stat.OnlineInboxCount.ToString("N0"));
                onlineItem.SubItems.Add(stat.OnlineSentCount.ToString("N0"));
                onlineItem.SubItems.Add((stat.OnlineInboxCount + stat.OnlineSentCount).ToString("N0"));
                onlineItem.Checked = stat.OnlineInboxCount > 0 || stat.OnlineSentCount > 0;
                onlineItem.Tag = stat;
                lvOnlineStats.Items.Add(onlineItem);

                // 右侧：本地数据（不带复选框）
                var localItem = new ListViewItem(stat.Year.ToString());
                localItem.SubItems.Add(stat.LocalInboxCount.ToString("N0"));
                localItem.SubItems.Add(stat.LocalSentCount.ToString("N0"));
                localItem.SubItems.Add((stat.LocalInboxCount + stat.LocalSentCount).ToString("N0"));
                lvLocalStats.Items.Add(localItem);
            }

            UpdateStatsSummary();
        }

        private void UpdateStatsSummary()
        {
            var selectedItems = lvOnlineStats.Items.Cast<ListViewItem>()
                .Where(i => i.Checked)
                .Select(i => (YearlyStatsItem)i.Tag)
                .ToList();

            int totalInbox = selectedItems.Sum(s => s.OnlineInboxCount);
            int totalSent = selectedItems.Sum(s => s.OnlineSentCount);
            int total = totalInbox + totalSent;

            lblStatsSummary.Text = $"已选择 {selectedItems.Count} 个年份，共 {total:N0} 封邮件待下载（收件箱: {totalInbox:N0}，已发送: {totalSent:N0}）";
            btnNext.Enabled = selectedItems.Count > 0;
        }

        // 同步参数类
        private class SyncArguments
        {
            public string OnlineStoreId { get; set; }
            public string LocalStoreId { get; set; }
            public List<SyncYearItem> Years { get; set; }
        }

        private class SyncYearItem
        {
            public int Year { get; set; }
            public int OnlineInboxCount { get; set; }
            public int OnlineSentCount { get; set; }
        }

        // 同步结果类
        private class SyncResult
        {
            public int TotalEmails { get; set; }
            public int SyncedEmails { get; set; }
            public bool Cancelled { get; set; }
            public string ErrorMessage { get; set; }
        }

        // 进度报告类
        private class SyncProgressReport
        {
            public int SyncedCount { get; set; }
            public int TotalCount { get; set; }
            public string Status { get; set; }
            public string LogMessage { get; set; }
        }

        private void StartSync()
        {
            _selectedYears.Clear();
            foreach (ListViewItem item in lvOnlineStats.Items)
            {
                if (item.Checked)
                    _selectedYears.Add((YearlyStatsItem)item.Tag);
            }

            ShowStep(WizardStep.Syncing);
            AddLog($"开始下载 {_selectedYears.Count} 个年份的邮件");

            // 准备同步参数
            var args = new SyncArguments
            {
                OnlineStoreId = _sourceStore?.StoreID,
                LocalStoreId = _targetStore?.StoreID,
                Years = _selectedYears.Select(y => new SyncYearItem
                {
                    Year = y.Year,
                    OnlineInboxCount = y.OnlineInboxCount,
                    OnlineSentCount = y.OnlineSentCount
                }).ToList()
            };

            // 创建并配置 BackgroundWorker
            _syncWorker = new System.ComponentModel.BackgroundWorker();
            _syncWorker.WorkerReportsProgress = true;
            _syncWorker.WorkerSupportsCancellation = true;

            _syncWorker.DoWork += SyncWorker_DoWork;
            _syncWorker.ProgressChanged += SyncWorker_ProgressChanged;
            _syncWorker.RunWorkerCompleted += SyncWorker_RunWorkerCompleted;

            // 启动后台任务
            _syncWorker.RunWorkerAsync(args);
        }

        // 后台工作线程 - 所有COM操作在这里执行
        private void SyncWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            var args = (SyncArguments)e.Argument;
            var worker = (System.ComponentModel.BackgroundWorker)sender;

            int totalEmails = args.Years.Sum(y => y.OnlineInboxCount + y.OnlineSentCount);
            int syncedEmails = 0;
            int currentYearIndex = 0;

            // 报告开始
            worker.ReportProgress(0, new SyncProgressReport
            {
                LogMessage = $"总共需要下载 {totalEmails} 封邮件"
            });

            try
            {
                // 获取Store对象
                Outlook.Store onlineStore = null;
                Outlook.Store localStore = null;

                foreach (Outlook.Store store in _application.Session.Stores)
                {
                    try
                    {
                        if (store.StoreID == args.OnlineStoreId)
                            onlineStore = store;
                        if (store.StoreID == args.LocalStoreId)
                            localStore = store;
                    }
                    catch { }
                }

                if (onlineStore == null)
                {
                    e.Result = new SyncResult { ErrorMessage = "无法找到联机存档数据文件" };
                    return;
                }

                // 获取文件夹
                Outlook.MAPIFolder onlineInbox = FindFolder(onlineStore, "收件箱", "Inbox");
                Outlook.MAPIFolder onlineSent = FindFolder(onlineStore, "已发送邮件", "Sent Items", "已发送");
                Outlook.MAPIFolder localInbox = localStore != null ? FindFolder(localStore, "收件箱", "Inbox") : null;
                Outlook.MAPIFolder localSent = localStore != null ? FindFolder(localStore, "已发送邮件", "Sent Items", "已发送") : null;

                try
                {
                    foreach (var yearItem in args.Years)
                    {
                        if (worker.CancellationPending)
                        {
                            e.Cancel = true;
                            break;
                        }

                        currentYearIndex++;

                        worker.ReportProgress(
                            totalEmails > 0 ? (int)((double)syncedEmails / totalEmails * 100) : 0,
                            new SyncProgressReport
                            {
                                SyncedCount = syncedEmails,
                                TotalCount = totalEmails,
                                Status = $"正在下载 {yearItem.Year} 年的邮件...",
                                LogMessage = $"[{currentYearIndex}/{args.Years.Count}] 正在下载 {yearItem.Year} 年的邮件..."
                            });

                        // 下载收件箱邮件
                        if (onlineInbox != null && yearItem.OnlineInboxCount > 0)
                        {
                            int inboxSynced = SyncFolderByYear(worker, onlineInbox, localInbox, yearItem.Year, syncedEmails, totalEmails, "收件箱");
                            syncedEmails += inboxSynced;

                            worker.ReportProgress(
                                totalEmails > 0 ? (int)((double)syncedEmails / totalEmails * 100) : 0,
                                new SyncProgressReport
                                {
                                    LogMessage = $"  收件箱: 已下载 {inboxSynced} 封邮件"
                                });
                        }

                        // 下载已发送邮件
                        if (onlineSent != null && yearItem.OnlineSentCount > 0)
                        {
                            int sentSynced = SyncFolderByYear(worker, onlineSent, localSent, yearItem.Year, syncedEmails, totalEmails, "已发送");
                            syncedEmails += sentSynced;

                            worker.ReportProgress(
                                totalEmails > 0 ? (int)((double)syncedEmails / totalEmails * 100) : 0,
                                new SyncProgressReport
                                {
                                    LogMessage = $"  已发送: 已下载 {sentSynced} 封邮件"
                                });
                        }
                    }
                }
                finally
                {
                    // 释放COM对象
                    if (onlineInbox != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(onlineInbox);
                    if (onlineSent != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(onlineSent);
                    if (localInbox != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(localInbox);
                    if (localSent != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(localSent);
                    if (onlineStore != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(onlineStore);
                    if (localStore != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(localStore);
                }

                e.Result = new SyncResult
                {
                    TotalEmails = totalEmails,
                    SyncedEmails = syncedEmails,
                    Cancelled = e.Cancel
                };
            }
            catch (System.Exception ex)
            {
                e.Result = new SyncResult
                {
                    TotalEmails = totalEmails,
                    SyncedEmails = syncedEmails,
                    ErrorMessage = ex.Message
                };
            }
        }

        // 进度更新 - 在UI线程执行
        private void SyncWorker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            var report = (SyncProgressReport)e.UserState;

            if (!string.IsNullOrEmpty(report.LogMessage))
                AddLog(report.LogMessage);

            if (!string.IsNullOrEmpty(report.Status))
                UpdateSyncProgress(report.SyncedCount, report.TotalCount, report.Status);
        }

        // 任务完成 - 在UI线程执行
        private void SyncWorker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            var result = (SyncResult)e.Result;

            if (result.Cancelled || _isCancelled)
            {
                AddLog("");
                AddLog("=== 下载已取消 ===");
                AddLog($"已下载: {result.SyncedEmails} 封邮件");
                AddLog($"未下载: {result.TotalEmails - result.SyncedEmails} 封邮件");

                MessageBox.Show(
                    $"下载已取消\n\n已下载: {result.SyncedEmails} 封邮件\n未下载: {result.TotalEmails - result.SyncedEmails} 封邮件",
                    "下载已取消",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else if (!string.IsNullOrEmpty(result.ErrorMessage))
            {
                AddLog($"✗ 下载失败: {result.ErrorMessage}");
                MessageBox.Show($"下载失败: {result.ErrorMessage}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                AddLog($"下载完成，共下载 {result.SyncedEmails} 封邮件");
                UpdateSyncProgress(result.TotalEmails, result.TotalEmails, "下载完成");
            }

            btnNext.Text = "完成";
            btnNext.Enabled = true;
            btnNext.Click -= BtnNext_Click;
            btnNext.Click += (s, ev) => this.Close();
            btnCancel.Enabled = true;
            btnCancel.Text = "取消";

            // 清理 BackgroundWorker
            if (_syncWorker != null)
            {
                _syncWorker.Dispose();
                _syncWorker = null;
            }
        }

        // 同步文件夹 - 在后台线程执行
        private int SyncFolderByYear(System.ComponentModel.BackgroundWorker worker, Outlook.MAPIFolder sourceFolder, Outlook.MAPIFolder targetFolder, int year, int alreadySynced, int totalEmails, string folderType)
        {
            int syncedCount = 0;
            int processedCount = 0;
            int lastReportedPercent = 0;

            try
            {
                var sourceItems = sourceFolder.Items;
                sourceItems.Sort("[ReceivedTime]", true);

                var filter = $"[ReceivedTime] >= '{year}/1/1' AND [ReceivedTime] < '{year + 1}/1/1'";
                Outlook.Items filteredItems = null;

                try
                {
                    filteredItems = sourceItems.Restrict(filter);
                }
                catch
                {
                    filteredItems = sourceItems;
                }

                int total = filteredItems.Count;
                object sourceItem = null;

                try
                {
                    sourceItem = filteredItems.GetFirst();

                    while (sourceItem != null)
                    {
                        if (worker.CancellationPending)
                            break;

                        processedCount++;

                        // 更新进度（减少频率）
                        int currentProgress = alreadySynced + syncedCount;
                        int percent = totalEmails > 0 ? (int)((double)currentProgress / totalEmails * 100) : 0;

                        if (percent > lastReportedPercent || processedCount % 20 == 0)
                        {
                            lastReportedPercent = percent;
                            worker.ReportProgress(percent, new SyncProgressReport
                            {
                                SyncedCount = currentProgress,
                                TotalCount = totalEmails,
                                Status = $"正在下载 {year} 年{folderType}邮件 ({processedCount}/{total})"
                            });
                        }

                        // 定期让出线程
                        if (processedCount % 10 == 0)
                        {
                            System.Threading.Thread.Sleep(5);
                        }

                        if (sourceItem is Outlook.MailItem sourceMail)
                        {
                            try
                            {
                                if (worker.CancellationPending)
                                    break;

                                DateTime receivedTime = sourceMail.ReceivedTime;

                                if (receivedTime.Year == year)
                                {
                                    bool exists = false;

                                    if (targetFolder != null)
                                    {
                                        try
                                        {
                                            if (worker.CancellationPending)
                                                break;

                                            var targetItems = targetFolder.Items;
                                            string subject = sourceMail.Subject ?? "";
                                            var checkFilter = $"[Subject] = '{subject.Replace("'", "''")}'";
                                            var checkItems = targetItems.Restrict(checkFilter);

                                            if (checkItems.Count > 0)
                                            {
                                                object checkItem = checkItems.GetFirst();
                                                while (checkItem != null)
                                                {
                                                    if (worker.CancellationPending)
                                                    {
                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(checkItem);
                                                        break;
                                                    }

                                                    if (checkItem is Outlook.MailItem targetMail)
                                                    {
                                                        try
                                                        {
                                                            if (Math.Abs((targetMail.ReceivedTime - receivedTime).TotalSeconds) < 2)
                                                            {
                                                                exists = true;
                                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(targetMail);
                                                                break;
                                                            }
                                                        }
                                                        finally
                                                        {
                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(targetMail);
                                                        }
                                                    }

                                                    var nextCheck = checkItems.GetNext();
                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(checkItem);
                                                    checkItem = nextCheck;
                                                }
                                            }

                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(checkItems);
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(targetItems);
                                        }
                                        catch { }
                                    }

                                    if (worker.CancellationPending)
                                        break;

                                    if (!exists)
                                    {
                                        var copiedMail = sourceMail.Copy();
                                        if (targetFolder != null)
                                        {
                                            copiedMail.Move(targetFolder);
                                        }
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedMail);
                                        syncedCount++;
                                    }
                                }
                            }
                            catch { }
                            finally
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMail);
                            }
                        }

                        var nextItem = filteredItems.GetNext();
                        if (sourceItem != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceItem);
                        sourceItem = nextItem;
                    }
                }
                finally
                {
                    if (sourceItem != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceItem);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceItems);
                }
            }
            catch { }

            return syncedCount;
        }

        private void UpdateSyncProgress(int current, int total, string status)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<int, int, string>(UpdateSyncProgress), current, total, status);
                return;
            }

            if (total > 0)
            {
                int percent = (int)((double)current / total * 100);
                progressSync.Value = Math.Min(percent, 100);
            }

            lblSyncStatus.Text = status;
        }

        public void AddLog(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<string>(AddLog), message);
                return;
            }

            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string logLine = $"[{timestamp}] {message}\r\n";

            txtLog.AppendText(logLine);

            var lines = txtLog.Lines;
            if (lines.Length > MAX_LOG_LINES)
            {
                var newLines = new string[MAX_LOG_LINES];
                Array.Copy(lines, lines.Length - MAX_LOG_LINES, newLines, 0, MAX_LOG_LINES);
                txtLog.Lines = newLines;
            }

            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.SelectionLength = 0;
            txtLog.ScrollToCaret();
            txtLog.Refresh();
        }
    }

    #endregion

    #region 阻止域对话框

    // 阻止域对话框（合并确认和结果显示）
    public class BlockDomainDialog : Form
    {
        private TextBox txtLog;
        private Button btnConfirm;
        private Button btnCancel;
        private Button btnCopy;
        private Button btnClose;
        private string domain;
        private string registryPath;
        private string valueName;
        private string domainEntry;
        private string fullRegistryPath;

        public BlockDomainDialog(string domain)
        {
            this.domain = domain;
            this.registryPath = @"Software\Microsoft\Office\16.0\Outlook\Options\Mail";
            this.valueName = "BlockedSenders";
            this.domainEntry = $"@{domain}";
            this.fullRegistryPath = $"HKEY_CURRENT_USER\\{registryPath}";

            this.Text = $"JTools-outlook - 阻止域 *@{domain}";
            this.Width = 650;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            InitializeComponents();
            ShowConfirmLog();
        }

        private void InitializeComponents()
        {
            // 标题标签
            Label lblTitle = new Label
            {
                Text = $"阻止域: *@{domain}",
                Font = new Font("Microsoft YaHei", 12, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(20, 20)
            };
            this.Controls.Add(lblTitle);

            // 说明标签
            Label lblDesc = new Label
            {
                Text = "此操作将把该域添加到 Outlook 的阻止发件人列表中。",
                AutoSize = true,
                Location = new Point(20, 50)
            };
            this.Controls.Add(lblDesc);

            // 日志文本框
            txtLog = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Width = 600,
                Height = 300,
                Location = new Point(20, 80),
                Font = new Font("Consolas", 10)
            };
            this.Controls.Add(txtLog);

            // 确认按钮
            btnConfirm = new Button
            {
                Text = "确认执行",
                Width = 100,
                Height = 30,
                Location = new Point(150, 400)
            };
            btnConfirm.Click += BtnConfirm_Click;
            this.Controls.Add(btnConfirm);

            // 取消按钮
            btnCancel = new Button
            {
                Text = "取消",
                Width = 100,
                Height = 30,
                Location = new Point(270, 400)
            };
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            // 复制日志按钮
            btnCopy = new Button
            {
                Text = "复制日志",
                Width = 100,
                Height = 30,
                Location = new Point(390, 400),
                Visible = false
            };
            btnCopy.Click += BtnCopy_Click;
            this.Controls.Add(btnCopy);

            // 关闭按钮
            btnClose = new Button
            {
                Text = "关闭",
                Width = 100,
                Height = 30,
                Location = new Point(390, 400),
                Visible = false
            };
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);

            this.CancelButton = btnCancel;
        }

        private void ShowConfirmLog()
        {
            var log = new System.Text.StringBuilder();
            log.AppendLine("【操作内容】");
            log.AppendLine($"将域 '*@{domain}' 添加到 Outlook 阻止发件人列表");
            log.AppendLine();
            log.AppendLine("【注册表修改】");
            log.AppendLine($"位置: {fullRegistryPath}");
            log.AppendLine($"值名: {valueName}");
            log.AppendLine($"类型: REG_MULTI_SZ (多字符串值)");
            log.AppendLine($"添加内容: {domainEntry}");
            log.AppendLine();
            log.AppendLine("【效果】");
            log.AppendLine("• 来自该域的所有邮件将被自动移动到垃圾邮件文件夹");
            log.AppendLine("• 当前邮件也会被移动到垃圾邮件文件夹");
            log.AppendLine("• 可能需要重启 Outlook 使设置生效");
            log.AppendLine();
            log.AppendLine("请确认是否继续执行？");

            txtLog.Text = log.ToString();
        }

        private void BtnConfirm_Click(object sender, EventArgs e)
        {
            // 添加调试信息
            System.Diagnostics.Debug.WriteLine("=== 点击了确认执行按钮 ===");

            try
            {
                // 执行注册表操作
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(registryPath, true);
                if (key == null)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(registryPath);
                }

                // 获取现有的阻止发件人列表
                string[] existingValues = (string[])key.GetValue(valueName, new string[0]);

                // 检查是否已包含该域
                bool alreadyExists = false;
                foreach (string value in existingValues)
                {
                    if (value.Equals(domainEntry, StringComparison.OrdinalIgnoreCase))
                    {
                        alreadyExists = true;
                        break;
                    }
                }

                if (!alreadyExists)
                {
                    // 添加新域到列表
                    var newValues = new string[existingValues.Length + 1];
                    existingValues.CopyTo(newValues, 0);
                    newValues[existingValues.Length] = domainEntry;

                    // 保存到注册表
                    key.SetValue(valueName, newValues, Microsoft.Win32.RegistryValueKind.MultiString);
                    key.Close();

                    // 将当前邮件移动到垃圾邮件文件夹
                    var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                    if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                    {
                        var selectedItem = explorer.Selection[1];
                        if (selectedItem is Outlook.MailItem mailItem)
                        {
                            var junkFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);
                            mailItem.Move(junkFolder);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(junkFolder);
                        }
                    }

                    // 显示成功日志
                    ShowSuccessLog();
                }
                else
                {
                    key.Close();
                    ShowAlreadyExistsLog();
                }
            }
            catch (System.Exception ex)
            {
                ShowErrorLog(ex.Message);
            }
        }

        private void ShowSuccessLog()
        {
            var log = new System.Text.StringBuilder();
            log.AppendLine("【操作成功】");
            log.AppendLine($"已将域 '*@{domain}' 添加到阻止发件人列表");
            log.AppendLine();
            log.AppendLine("【注册表修改】");
            log.AppendLine($"位置: {fullRegistryPath}");
            log.AppendLine($"值名: {valueName}");
            log.AppendLine($"类型: REG_MULTI_SZ (多字符串值)");
            log.AppendLine($"添加内容: {domainEntry}");
            log.AppendLine();
            log.AppendLine("【效果】");
            log.AppendLine("来自该域的所有邮件将被自动移动到垃圾邮件文件夹");
            log.AppendLine();
            log.AppendLine("【提示】");
            log.AppendLine("可能需要重启 Outlook 使设置生效");

            txtLog.Text = log.ToString();
            SwitchToResultMode();
        }

        private void ShowAlreadyExistsLog()
        {
            var log = new System.Text.StringBuilder();
            log.AppendLine("【提示】");
            log.AppendLine($"域 '*@{domain}' 已在阻止发件人列表中");
            log.AppendLine();
            log.AppendLine("【注册表位置】");
            log.AppendLine($"位置: {fullRegistryPath}");
            log.AppendLine($"值名: {valueName}");

            txtLog.Text = log.ToString();
            SwitchToResultMode();
        }

        private void ShowErrorLog(string errorMessage)
        {
            var log = new System.Text.StringBuilder();
            log.AppendLine("【操作失败】");
            log.AppendLine($"错误: {errorMessage}");
            log.AppendLine();
            log.AppendLine("【手动添加步骤】");
            log.AppendLine("1. 点击\"开始\"选项卡");
            log.AppendLine("2. 点击\"删除\"组中的\"垃圾邮件\"");
            log.AppendLine("3. 选择\"垃圾邮件选项\"");
            log.AppendLine("4. 在\"阻止发件人\"选项卡中点击\"添加\"");
            log.AppendLine($"5. 输入: *@{domain}");
            log.AppendLine("6. 点击\"确定\"");

            txtLog.Text = log.ToString();
            SwitchToResultMode();
        }

        private void SwitchToResultMode()
        {
            // 更新窗口标题
            this.Text = $"JTools-outlook - 阻止域结果 *@{domain}";

            // 隐藏确认和取消按钮
            btnConfirm.Visible = false;
            btnCancel.Visible = false;

            // 显示复制和关闭按钮
            btnCopy.Visible = true;
            btnClose.Visible = true;
            btnClose.Left = 230; // 调整关闭按钮位置

            // 设置关闭按钮为默认按钮
            this.CancelButton = btnClose;
            this.AcceptButton = btnClose;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            // 添加调试信息
            System.Diagnostics.Debug.WriteLine("=== 点击了取消按钮 ===");

            // 点击取消按钮，直接关闭对话框，不执行任何操作
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void BtnCopy_Click(object sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(txtLog.Text);
                MessageBox.Show("日志已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }

    #endregion
}
