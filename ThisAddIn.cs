using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace jtools_outlook
{
    public partial class ThisAddIn
    {
        private const string AppVersion = "v1.1.1";

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
                Text = "版本 v1.1.1",
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

    #region 下载联机功能

    /// <summary>
    /// 年份统计信息
    /// </summary>
    public class YearStats
    {
        public int Year { get; set; }
        public int InboxCount { get; set; }
        public int SentCount { get; set; }
        public int TotalCount { get { return InboxCount + SentCount; } }

        public override string ToString()
        {
            return $"{Year} 年 (收件箱: {InboxCount}, 已发送: {SentCount}, 共: {TotalCount})";
        }
    }

    /// <summary>
    /// 下载联机窗体
    /// </summary>
    public class DownloadOnlineForm : Form
    {
        private Outlook.Application _application;
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isRunning = false;
        private Dictionary<int, YearStats> _yearStats = new Dictionary<int, YearStats>();
        private HashSet<string> _downloadedEntryIds = new HashSet<string>();

        // UI 控件
        private ComboBox cmbSourceStore;
        private Button btnAnalyze;
        private CheckedListBox chkYears;
        private TextBox txtTargetFolder;
        private Button btnBrowseFolder;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblProgress;
        private TextBox txtLog;
        private Button btnStart;
        private Button btnCancel;
        private Button btnClose;
        private Panel selectPanel;
        private Panel progressPanel;
        private GroupBox grpYears;

        public DownloadOnlineForm(Outlook.Application application)
        {
            _application = application;
            InitializeComponent();
            LoadStores();
        }

        private void InitializeComponent()
        {
            this.Text = "下载联机存档 - JTools-outlook v1.1.1";
            this.Width = 750;
            this.Height = 650;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = false;
            this.MinimumSize = new Size(650, 550);

            // 主面板
            var mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };

            // 选择面板
            selectPanel = new Panel { Dock = DockStyle.Top, Height = 320 };
            var lblTitle = new Label
            {
                Text = "下载联机存档到本地文件夹",
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Microsoft YaHei", 12, FontStyle.Bold),
                ForeColor = Color.SteelBlue
            };

            // 源数据文件选择
            var lblSource = new Label
            {
                Text = "源数据文件（联机存档）:",
                Dock = DockStyle.Top,
                Height = 25,
                Margin = new Padding(0, 10, 0, 0)
            };

            var sourcePanel = new Panel { Dock = DockStyle.Top, Height = 32 };
            cmbSourceStore = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 28
            };
            btnAnalyze = new Button
            {
                Text = "分析",
                Width = 80,
                Height = 28,
                Dock = DockStyle.Right
            };
            btnAnalyze.Click += BtnAnalyze_Click;
            sourcePanel.Controls.Add(cmbSourceStore);
            sourcePanel.Controls.Add(btnAnalyze);

            // 年份选择区域
            grpYears = new GroupBox
            {
                Text = "选择要下载的年份（分析后显示）",
                Dock = DockStyle.Top,
                Height = 150,
                Margin = new Padding(0, 10, 0, 0)
            };

            chkYears = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                Font = new Font("Microsoft YaHei", 9),
                Margin = new Padding(5)
            };
            grpYears.Controls.Add(chkYears);

            // 目标文件夹选择
            var lblTarget = new Label
            {
                Text = "目标文件夹:",
                Dock = DockStyle.Top,
                Height = 25,
                Margin = new Padding(0, 10, 0, 0)
            };

            var folderPanel = new Panel { Dock = DockStyle.Top, Height = 32 };
            txtTargetFolder = new TextBox
            {
                Dock = DockStyle.Fill,
                Height = 28
            };
            btnBrowseFolder = new Button
            {
                Text = "浏览...",
                Width = 80,
                Height = 28,
                Dock = DockStyle.Right
            };
            btnBrowseFolder.Click += BtnBrowseFolder_Click;
            folderPanel.Controls.Add(txtTargetFolder);
            folderPanel.Controls.Add(btnBrowseFolder);

            selectPanel.Controls.AddRange(new Control[] {
                folderPanel, lblTarget,
                grpYears,
                sourcePanel, lblSource,
                lblTitle
            });

            // 进度面板
            progressPanel = new Panel { Dock = DockStyle.Top, Height = 80, Margin = new Padding(0, 10, 0, 0) };

            lblStatus = new Label
            {
                Text = "就绪",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font("Microsoft YaHei", 9)
            };

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 25,
                Minimum = 0,
                Maximum = 100
            };

            lblProgress = new Label
            {
                Text = "0 / 0 (0%)",
                Dock = DockStyle.Top,
                Height = 25,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei", 9, FontStyle.Bold)
            };

            progressPanel.Controls.AddRange(new Control[] { lblProgress, progressBar, lblStatus });

            // 日志面板
            var logPanel = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };

            var lblLog = new Label
            {
                Text = "操作日志",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font("Microsoft YaHei", 9, FontStyle.Bold)
            };

            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new Font("Consolas", 9),
                BackColor = Color.WhiteSmoke
            };

            logPanel.Controls.AddRange(new Control[] { txtLog, lblLog });

            // 按钮面板
            var buttonPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };

            btnStart = new Button
            {
                Text = "开始下载",
                Width = 100,
                Height = 32,
                Left = 15,
                Top = 9
            };
            btnStart.Click += BtnStart_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 32,
                Left = 125,
                Top = 9,
                Enabled = false
            };
            btnCancel.Click += BtnCancel_Click;

            btnClose = new Button
            {
                Text = "关闭",
                Width = 80,
                Height = 32,
                Left = 630,
                Top = 9,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnClose.Click += (s, e) => this.Close();

            buttonPanel.Controls.AddRange(new Control[] { btnStart, btnCancel, btnClose });

            mainPanel.Controls.AddRange(new Control[] { logPanel, progressPanel, selectPanel });
            this.Controls.AddRange(new Control[] { mainPanel, buttonPanel });
        }

        private void BtnBrowseFolder_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择保存邮件的目标文件夹";
                dialog.ShowNewFolderButton = true;
                if (!string.IsNullOrEmpty(txtTargetFolder.Text) && Directory.Exists(txtTargetFolder.Text))
                {
                    dialog.SelectedPath = txtTargetFolder.Text;
                }
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtTargetFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void LoadStores()
        {
            try
            {
                cmbSourceStore.Items.Clear();
                int archiveCount = 0;

                foreach (Outlook.Store store in _application.Session.Stores)
                {
                    try
                    {
                        bool isArchive = store.ExchangeStoreType != Outlook.OlExchangeStoreType.olNotExchange &&
                                        store.ExchangeStoreType != Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox;

                        if (isArchive)
                        {
                            var info = new StoreInfo
                            {
                                Store = store,
                                DisplayName = store.DisplayName,
                                IsArchive = true
                            };
                            cmbSourceStore.Items.Add(info);
                            archiveCount++;
                        }
                    }
                    catch { }
                }

                if (cmbSourceStore.Items.Count > 0)
                {
                    cmbSourceStore.SelectedIndex = 0;
                }

                AddLog($"已加载 {archiveCount} 个联机存档");
                AddLog("请选择源数据文件后点击\"分析\"按钮");
            }
            catch (System.Exception ex)
            {
                AddLog($"加载数据文件失败: {ex.Message}");
            }
        }

        private void AddLog(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new System.Action(() => AddLog(message)));
                return;
            }

            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            txtLog.AppendText($"[{timestamp}] {message}\r\n");
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        private async void BtnAnalyze_Click(object sender, EventArgs e)
        {
            var sourceInfo = cmbSourceStore.SelectedItem as StoreInfo;
            if (sourceInfo?.Store == null)
            {
                MessageBox.Show("请选择源数据文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnAnalyze.Enabled = false;
            chkYears.Items.Clear();
            _yearStats.Clear();

            // 先更新 UI 显示状态
            grpYears.Text = "选择要下载的年份（正在分析...）";
            AddLog($"正在分析 {sourceInfo.DisplayName}...");

            // 强制刷新 UI，确保用户看到状态变化
            this.Refresh();
            System.Windows.Forms.Application.DoEvents();

            try
            {
                // 获取 StoreId（在主线程获取，避免跨线程 COM 问题）
                string storeId = sourceInfo.Store.StoreID;
                string storeName = sourceInfo.DisplayName;

                // 在后台线程执行分析
                await Task.Run(() => AnalyzeStoreInBackground(storeId, storeName));

                // 显示年份统计
                grpYears.Text = $"选择要下载的年份（共 {_yearStats.Count} 个年份）";
                foreach (var stats in _yearStats.Values.OrderByDescending(y => y.Year))
                {
                    chkYears.Items.Add(stats, true);
                }

                AddLog($"分析完成，共发现 {_yearStats.Count} 个年份");
            }
            catch (System.Exception ex)
            {
                AddLog($"分析失败: {ex.Message}");
                MessageBox.Show($"分析失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnAnalyze.Enabled = true;
            }
        }

        /// <summary>
        /// 在后台线程分析邮件（使用 STAThread 处理 COM）
        /// </summary>
        private void AnalyzeStoreInBackground(string storeId, string storeName)
        {
            System.Exception bgException = null;

            // 在 STAThread 后台线程中执行分析
            var thread = new Thread(() =>
            {
                try
                {
                    LogToUi("[后台] 开始创建 Outlook 实例...");

                    // 创建新的 Outlook Application 实例
                    var app = new Outlook.Application();
                    var ns = app.GetNamespace("MAPI");

                    LogToUi("[后台] Outlook 实例创建成功");

                    try
                    {
                        // 重新获取 Store
                        LogToUi("[后台] 正在查找数据文件...");
                        Outlook.Store bgStore = null;
                        int storeCount = 0;
                        foreach (Outlook.Store s in ns.Stores)
                        {
                            storeCount++;
                            try
                            {
                                if (s.StoreID == storeId)
                                {
                                    bgStore = s;
                                    LogToUi($"[后台] 找到目标数据文件 (共扫描 {storeCount} 个)");
                                    break;
                                }
                            }
                            catch { }
                        }

                        if (bgStore == null)
                        {
                            LogToUi("[后台] 未找到目标数据文件！");
                        }

                        if (bgStore != null)
                        {
                            // 获取文件夹 EntryID
                            LogToUi("[后台] 正在查找收件箱...");
                            string inboxEntryId = null;
                            string sentEntryId = null;

                            try
                            {
                                var inbox = FindFolder(bgStore, "收件箱", "Inbox");
                                if (inbox != null)
                                {
                                    inboxEntryId = inbox.EntryID;
                                    LogToUi($"[后台] 找到收件箱: {inbox.Name}");
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);
                                }
                                else
                                {
                                    LogToUi("[后台] 未找到收件箱");
                                }
                            }
                            catch (System.Exception ex)
                            {
                                LogToUi($"[后台] 查找收件箱失败: {ex.Message}");
                            }

                            LogToUi("[后台] 正在查找已发送邮件...");
                            try
                            {
                                var sent = FindFolder(bgStore, "已发送邮件", "Sent Items", "已发送");
                                if (sent != null)
                                {
                                    sentEntryId = sent.EntryID;
                                    LogToUi($"[后台] 找到已发送: {sent.Name}");
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sent);
                                }
                                else
                                {
                                    LogToUi("[后台] 未找到已发送邮件文件夹");
                                }
                            }
                            catch (System.Exception ex)
                            {
                                LogToUi($"[后台] 查找已发送失败: {ex.Message}");
                            }

                            // 分析收件箱
                            if (!string.IsNullOrEmpty(inboxEntryId))
                            {
                                LogToUi("[后台] 开始分析收件箱...");
                                try
                                {
                                    var inbox = ns.GetFolderFromID(inboxEntryId);
                                    if (inbox != null)
                                    {
                                        AnalyzeFolderByYearUsingTable(inbox, "Inbox");
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);
                                        LogToUi("[后台] 收件箱分析完成");
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    LogToUi($"[后台] 分析收件箱失败: {ex.Message}");
                                }
                            }

                            // 分析已发送
                            if (!string.IsNullOrEmpty(sentEntryId))
                            {
                                LogToUi("[后台] 开始分析已发送邮件...");
                                try
                                {
                                    var sent = ns.GetFolderFromID(sentEntryId);
                                    if (sent != null)
                                    {
                                        AnalyzeFolderByYearUsingTable(sent, "Sent");
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sent);
                                        LogToUi("[后台] 已发送邮件分析完成");
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    LogToUi($"[后台] 分析已发送失败: {ex.Message}");
                                }
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(bgStore);
                        }
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ns);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                        LogToUi("[后台] Outlook 资源已释放");
                    }
                }
                catch (System.Exception ex)
                {
                    bgException = ex;
                    LogToUi($"[后台] 分析失败: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"后台分析失败: {ex.Message}");
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            LogToUi("[主线程] 启动后台分析线程...");
            thread.Start();
            thread.Join();  // 等待线程完成
            LogToUi("[主线程] 后台分析线程已完成");

            if (bgException != null)
            {
                throw bgException;
            }
        }

        /// <summary>
        /// 线程安全的日志输出
        /// </summary>
        private void LogToUi(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new System.Action(() => LogToUi(message)));
                return;
            }
            AddLog(message);
        }

        /// <summary>
        /// 使用 GetTable() 高效分析文件夹（比遍历 Items 快得多）
        /// </summary>
        private void AnalyzeFolderByYearUsingTable(Outlook.MAPIFolder folder, string folderType)
        {
            try
            {
                LogToUi($"[分析] 获取 {folderType} 邮件总数...");

                // 先获取总数量
                var items = folder.Items;
                int totalCount = items.Count;
                LogToUi($"[分析] {folderType} 共有 {totalCount} 封邮件");

                if (totalCount == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                    return;
                }

                // 获取年份范围
                LogToUi($"[分析] 获取 {folderType} 年份范围...");
                items.Sort("[ReceivedTime]", true);

                DateTime? minDate = null;
                DateTime? maxDate = null;

                // 获取最早和最晚的邮件日期
                try
                {
                    var firstItem = items[1];
                    if (firstItem is Outlook.MailItem firstMail)
                    {
                        minDate = firstMail.ReceivedTime;
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(firstMail);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(firstItem);
                }
                catch { }

                try
                {
                    var lastItem = items[totalCount];
                    if (lastItem is Outlook.MailItem lastMail)
                    {
                        maxDate = lastMail.ReceivedTime;
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(lastMail);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(lastItem);
                }
                catch { }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(items);

                if (minDate.HasValue && maxDate.HasValue)
                {
                    int minYear = minDate.Value.Year;
                    int maxYear = maxDate.Value.Year;

                    // 确保年份范围正确（minYear 可能大于 maxYear，因为排序是降序）
                    int startYear = Math.Min(minYear, maxYear);
                    int endYear = Math.Max(minYear, maxYear);

                    LogToUi($"[分析] {folderType} 年份范围: {startYear} - {endYear}");

                    // 按年份统计数量（使用 Restrict 过滤）
                    for (int year = startYear; year <= endYear; year++)
                    {
                        try
                        {
                            // 使用 Restrict 按年份过滤，然后取 Count
                            string filter = $"[ReceivedTime] >= '{year}/1/1' AND [ReceivedTime] < '{year + 1}/1/1'";
                            var yearItems = folder.Items.Restrict(filter);
                            int yearCount = yearItems.Count;

                            if (yearCount > 0)
                            {
                                lock (_yearStats)
                                {
                                    if (!_yearStats.ContainsKey(year))
                                    {
                                        _yearStats[year] = new YearStats { Year = year };
                                    }
                                    if (folderType == "Inbox")
                                        _yearStats[year].InboxCount = yearCount;
                                    else
                                        _yearStats[year].SentCount = yearCount;
                                }
                                LogToUi($"[分析] {folderType} {year}年: {yearCount} 封");
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(yearItems);
                        }
                        catch (System.Exception ex)
                        {
                            LogToUi($"[分析] 统计 {year} 年失败: {ex.Message}");
                        }
                    }
                }
                else
                {
                    LogToUi($"[分析] 无法获取 {folderType} 年份范围，使用遍历方式...");
                    AnalyzeFolderByYearFallback(folder, folderType);
                }
            }
            catch (System.Exception ex)
            {
                LogToUi($"[分析] 分析失败: {ex.Message}，尝试回退方法...");
                System.Diagnostics.Debug.WriteLine($"分析失败: {ex.Message}");
                AnalyzeFolderByYearFallback(folder, folderType);
            }
        }

        /// <summary>
        /// 回退方法：当 GetTable 不可用时使用（分批处理避免阻塞）
        /// </summary>
        private void AnalyzeFolderByYearFallback(Outlook.MAPIFolder folder, string folderType)
        {
            try
            {
                var items = folder.Items;
                int count = items.Count;
                int batchSize = 100;  // 每批处理100封

                for (int batch = 0; batch < (count + batchSize - 1) / batchSize; batch++)
                {
                    int start = batch * batchSize + 1;
                    int end = Math.Min(start + batchSize - 1, count);

                    for (int i = start; i <= end; i++)
                    {
                        try
                        {
                            var item = items[i];
                            if (item is Outlook.MailItem mail)
                            {
                                int year = mail.ReceivedTime.Year;

                                lock (_yearStats)
                                {
                                    if (!_yearStats.ContainsKey(year))
                                    {
                                        _yearStats[year] = new YearStats { Year = year };
                                    }
                                    if (folderType == "Inbox")
                                        _yearStats[year].InboxCount++;
                                    else
                                        _yearStats[year].SentCount++;
                                }

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                            }
                            if (item != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                        }
                        catch { }
                    }

                    // 每批后让出线程
                    Thread.Sleep(1);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
            }
            catch { }
        }

        private Outlook.MAPIFolder FindFolder(Outlook.Store store, params string[] possibleNames)
        {
            try
            {
                var root = store.GetRootFolder();
                foreach (Outlook.MAPIFolder folder in root.Folders)
                {
                    try
                    {
                        foreach (var name in possibleNames)
                        {
                            if (folder.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(root);
                                return folder;
                            }
                        }
                    }
                    catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(root);
            }
            catch { }
            return null;
        }

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (_isRunning) return;

            var sourceInfo = cmbSourceStore.SelectedItem as StoreInfo;
            if (sourceInfo?.Store == null)
            {
                MessageBox.Show("请选择源数据文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (chkYears.CheckedItems.Count == 0)
            {
                MessageBox.Show("请选择要下载的年份", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtTargetFolder.Text))
            {
                MessageBox.Show("请选择目标文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 确保目标目录存在
            string targetFolder = txtTargetFolder.Text;
            if (!Directory.Exists(targetFolder))
            {
                try
                {
                    Directory.CreateDirectory(targetFolder);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"无法创建目标目录: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // 获取选中的年份
            var selectedYears = chkYears.CheckedItems.Cast<YearStats>().Select(s => s.Year).OrderByDescending(y => y).ToList();

            _isRunning = true;
            btnStart.Enabled = false;
            btnCancel.Enabled = true;
            btnAnalyze.Enabled = false;
            cmbSourceStore.Enabled = false;
            btnBrowseFolder.Enabled = false;

            _cancellationTokenSource = new CancellationTokenSource();

            try
            {
                await DownloadEmailsAsync(sourceInfo.Store, targetFolder, selectedYears, _cancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                AddLog("下载已取消");
            }
            catch (System.Exception ex)
            {
                AddLog($"下载失败: {ex.Message}");
                MessageBox.Show($"下载失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isRunning = false;
                btnStart.Enabled = true;
                btnCancel.Enabled = false;
                btnAnalyze.Enabled = true;
                cmbSourceStore.Enabled = true;
                btnBrowseFolder.Enabled = true;
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _cancellationTokenSource.Cancel();
                AddLog("正在取消下载...");
                btnCancel.Enabled = false;
            }
        }

        private async Task DownloadEmailsAsync(Outlook.Store sourceStore, string targetFolder, List<int> selectedYears, CancellationToken cancellationToken)
        {
            AddLog($"开始下载邮件...");
            AddLog($"源: {sourceStore.DisplayName}");
            AddLog($"目标: {targetFolder}");
            AddLog($"选中年份: {string.Join(", ", selectedYears)}");

            int totalDownloaded = 0;
            int totalSkipped = 0;

            try
            {
                // 获取源文件夹
                var sourceFolders = new List<Outlook.MAPIFolder>();
                var folderNames = new List<string>();

                var sourceInbox = FindFolder(sourceStore, "收件箱", "Inbox");
                if (sourceInbox != null)
                {
                    sourceFolders.Add(sourceInbox);
                    folderNames.Add("收件箱");
                    AddLog($"找到收件箱: {sourceInbox.Name}");
                }

                var sourceSent = FindFolder(sourceStore, "已发送邮件", "Sent Items", "已发送");
                if (sourceSent != null)
                {
                    sourceFolders.Add(sourceSent);
                    folderNames.Add("已发送邮件");
                    AddLog($"找到已发送: {sourceSent.Name}");
                }

                if (sourceFolders.Count == 0)
                {
                    AddLog("未找到可下载的文件夹");
                    return;
                }

                // 按年份下载
                foreach (var year in selectedYears)
                {
                    if (cancellationToken.IsCancellationRequested)
                        break;

                    AddLog($"");
                    AddLog($"--- 下载 {year} 年邮件 ---");

                    // 创建年份目录
                    string yearFolder = Path.Combine(targetFolder, year.ToString());
                    if (!Directory.Exists(yearFolder))
                    {
                        Directory.CreateDirectory(yearFolder);
                    }

                    // 处理每个源文件夹
                    for (int i = 0; i < sourceFolders.Count; i++)
                    {
                        if (cancellationToken.IsCancellationRequested)
                            break;

                        var sourceFolder = sourceFolders[i];
                        var folderName = folderNames[i];

                        // 创建子文件夹
                        string subFolder = Path.Combine(yearFolder, folderName);
                        if (!Directory.Exists(subFolder))
                        {
                            Directory.CreateDirectory(subFolder);
                        }

                        AddLog($"下载 {folderName} 到 {subFolder}...");

                        int yearDownloaded = 0;
                        int yearSkipped = 0;

                        await Task.Run(() =>
                        {
                            DownloadFolderToFiles(sourceFolder, year, subFolder, cancellationToken,
                                ref yearDownloaded, ref yearSkipped);
                        }, cancellationToken);

                        totalDownloaded += yearDownloaded;
                        totalSkipped += yearSkipped;

                        AddLog($"  {folderName}: 下载 {yearDownloaded}，跳过 {yearSkipped}");

                        if (cancellationToken.IsCancellationRequested)
                            break;
                    }

                    if (cancellationToken.IsCancellationRequested)
                        break;
                }

                // 释放源文件夹
                foreach (var folder in sourceFolders)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
                }
            }
            finally
            {
                AddLog($"");
                AddLog($"========== 下载完成 ==========");
                AddLog($"总计: 下载 {totalDownloaded} 封，跳过 {totalSkipped} 封");
            }
        }

        private void DownloadFolderToFiles(Outlook.MAPIFolder sourceFolder, int year, string targetPath,
            CancellationToken cancellationToken, ref int downloadedCount, ref int skippedCount)
        {
            try
            {
                // 使用 Restrict 过滤该年份的邮件
                string filter = $"[ReceivedTime] >= '{year}/1/1' AND [ReceivedTime] < '{year + 1}/1/1'";
                var items = sourceFolder.Items.Restrict(filter);
                items.Sort("[ReceivedTime]", true);

                int total = items.Count;
                int processed = 0;
                int batchSize = 20;

                for (int i = 1; i <= total; i++)
                {
                    if (cancellationToken.IsCancellationRequested)
                        break;

                    object item = null;
                    try
                    {
                        item = items[i];
                        processed++;

                        if (item is Outlook.MailItem mail)
                        {
                            string entryId = mail.EntryID;

                            // 检查是否已下载（通过检查文件是否存在）
                            string safeSubject = GetSafeFileName(mail.Subject, entryId);
                            string filePath = Path.Combine(targetPath, safeSubject + ".msg");

                            if (File.Exists(filePath))
                            {
                                skippedCount++;
                            }
                            else
                            {
                                // 保存邮件为 .msg 文件
                                mail.SaveAs(filePath);
                                downloadedCount++;
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                        }

                        // 更新进度
                        if (processed % 10 == 0)
                        {
                            UpdateProgress(processed, total, downloadedCount, skippedCount, year, "下载中");
                        }

                        // 批次间让出线程
                        if (processed % batchSize == 0)
                        {
                            Thread.Sleep(10);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        AddLog($"保存邮件失败: {ex.Message}");
                    }
                    finally
                    {
                        if (item != null)
                        {
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(item); } catch { }
                        }
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
            }
            catch (System.Exception ex)
            {
                AddLog($"下载文件夹失败: {ex.Message}");
            }
        }

        private string GetSafeFileName(string subject, string entryId)
        {
            // 移除非法字符
            string safeName = subject ?? "无主题";
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(c, '_');
            }
            // 限制长度并添加 EntryID 后缀确保唯一性
            if (safeName.Length > 100)
            {
                safeName = safeName.Substring(0, 100);
            }
            // 使用 EntryID 的前8位作为后缀确保唯一性
            string suffix = entryId?.Length > 8 ? entryId.Substring(0, 8) : entryId ?? Guid.NewGuid().ToString().Substring(0, 8);
            return $"{safeName}_{suffix}";
        }

        private void UpdateProgress(int processed, int total, int downloaded, int skipped, int year, string stage = "")
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new System.Action(() => UpdateProgress(processed, total, downloaded, skipped, year, stage)));
                return;
            }

            int percent = total > 0 ? (int)((double)processed / total * 100) : 0;
            progressBar.Value = Math.Min(percent, 100);
            lblProgress.Text = $"{processed} / {total} ({percent}%)";
            string stageText = string.IsNullOrEmpty(stage) ? "" : $"[{stage}] ";
            lblStatus.Text = $"{stageText}{year}年 - 已下载: {downloaded}，跳过: {skipped}";
        }
    }

    /// <summary>
    /// 阻止域对话框
    /// </summary>
    public class BlockDomainDialog : Form
    {
        private string _domain;
        private Label lblMessage;
        private Button btnOK;
        private Button btnCancel;

        public BlockDomainDialog(string domain)
        {
            _domain = domain;
            this.Text = "阻止域";
            this.Width = 400;
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
                Padding = new Padding(20)
            };

            lblMessage = new Label
            {
                Text = $"确定要阻止来自 @{_domain} 的所有邮件吗？",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei", 10),
                Height = 50
            };

            var buttonPanel = new Panel { Height = 50 };
            btnOK = new Button
            {
                Text = "确定",
                Width = 80,
                Height = 30,
                Left = 100,
                Top = 10,
                DialogResult = DialogResult.OK
            };
            btnCancel = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 30,
                Left = 200,
                Top = 10,
                DialogResult = DialogResult.Cancel
            };

            buttonPanel.Controls.Add(btnOK);
            buttonPanel.Controls.Add(btnCancel);

            tableLayout.Controls.Add(lblMessage, 0, 0);
            tableLayout.Controls.Add(buttonPanel, 0, 2);

            this.Controls.Add(tableLayout);
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }
    }

    #endregion
}
