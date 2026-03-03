using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace jimsoutlooktools
{
    public partial class ThisAddIn
    {
        private const string AppVersion = "v1.0.4";

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

        public ProgressForm(string appVersion = "v1.0.3")
        {
            IsCancelled = false;
            this.Text = $"jimsoutlooktools {appVersion} - 保存进度";
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

        public DateRangePickerForm(string appVersion = "v1.0.3")
        {
            this.Text = $"jimsoutlooktools {appVersion} - 保存邮件附件";
            this.Width = 480;
            this.Height = 540;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 主容器
            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 12,
                Padding = new Padding(20)
            };

            // 品牌标题
            var brandLabel = new Label
            {
                Text = $"jimsoutlooktools {appVersion}",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 14, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 35
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
            tableLayout.Controls.Add(brandLabel, 0, 0);
            tableLayout.Controls.Add(pathLabel, 0, 1);
            tableLayout.Controls.Add(pathPanel, 0, 2);
            tableLayout.Controls.Add(startLabel, 0, 3);
            tableLayout.Controls.Add(startDatePicker, 0, 4);
            tableLayout.Controls.Add(spacer, 0, 5);
            tableLayout.Controls.Add(endLabel, 0, 6);
            tableLayout.Controls.Add(endDatePicker, 0, 7);
            tableLayout.Controls.Add(spacer2, 0, 8);
            tableLayout.Controls.Add(folderLabel, 0, 9);
            tableLayout.Controls.Add(chkInbox, 0, 10);
            tableLayout.Controls.Add(chkSentItems, 0, 11);
            tableLayout.Controls.Add(buttonPanel, 0, 12);

            // 设置行高 - 增加间距避免重叠
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45));  // 品牌标题
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
        public SaveResultForm(string appVersion, string resultText)
        {
            this.Text = $"jimsoutlooktools {appVersion} - 保存结果详情";
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
    /// 数据文件选择窗体
    /// </summary>
    public class DataFileSelectForm : Form
    {
        private Outlook.Application _application;
        private ComboBox cmbSourceStore;
        private ComboBox cmbTargetStore;
        private Button btnOK;
        private Button btnCancel;
        private Label lblLoading;
        private ProgressBar progressLoading;

        public Outlook.MAPIFolder SourceRootFolder { get; private set; }
        public Outlook.MAPIFolder TargetRootFolder { get; private set; }

        public DataFileSelectForm(Outlook.Application application)
        {
            _application = application;
            InitializeComponent();
            // 异步加载数据文件
            LoadStoresAsync();
        }

        private void InitializeComponent()
        {
            this.Text = "选择数据文件";
            this.Width = 500;
            this.Height = 320;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(15)
            };

            // 标题
            var titleLabel = new Label
            {
                Text = "选择源数据文件和目标数据文件",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 30
            };

            // 加载状态面板
            var loadingPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 40
            };

            lblLoading = new Label
            {
                Text = "正在加载数据文件列表，请稍候...",
                Left = 0,
                Top = 0,
                Width = 450,
                Height = 20,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                ForeColor = System.Drawing.Color.SteelBlue
            };

            progressLoading = new ProgressBar
            {
                Left = 0,
                Top = 22,
                Width = 450,
                Height = 15,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };

            loadingPanel.Controls.Add(lblLoading);
            loadingPanel.Controls.Add(progressLoading);

            // 源数据文件（联机存档）
            var sourceLabel = new Label
            {
                Text = "源数据文件（联机存档）:",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            cmbSourceStore = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 25,
                Enabled = false
            };

            // 目标数据文件（本地PST）
            var targetLabel = new Label
            {
                Text = "目标数据文件（本地PST）:",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold),
                Height = 25
            };

            cmbTargetStore = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 25,
                Enabled = false
            };

            // 按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };

            btnOK = new Button
            {
                Text = "下一步",
                Width = 80,
                Height = 30,
                Left = 280,
                Top = 5,
                Enabled = false
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 30,
                Left = 380,
                Top = 5,
                DialogResult = DialogResult.Cancel
            };

            buttonPanel.Controls.Add(btnOK);
            buttonPanel.Controls.Add(btnCancel);

            // 说明标签
            var hintLabel = new Label
            {
                Text = "提示：源数据文件是Office 365联机存档，目标数据文件是本地PST文件",
                Dock = DockStyle.Fill,
                ForeColor = System.Drawing.Color.Gray,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8),
                Height = 20
            };

            tableLayout.Controls.Add(titleLabel, 0, 0);
            tableLayout.Controls.Add(loadingPanel, 0, 1);
            tableLayout.Controls.Add(sourceLabel, 0, 2);
            tableLayout.Controls.Add(cmbSourceStore, 0, 3);
            tableLayout.Controls.Add(targetLabel, 0, 4);
            tableLayout.Controls.Add(cmbTargetStore, 0, 5);
            tableLayout.Controls.Add(hintLabel, 0, 6);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));

            this.Controls.Add(tableLayout);
            this.Controls.Add(buttonPanel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private async void LoadStoresAsync()
        {
            lblLoading.Text = "正在连接 Outlook 并加载数据文件列表，请稍候...";
            progressLoading.Visible = true;
            
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

                var sourceList = storeList.Where(s => s.IsArchive).ToList();
                var targetList = storeList.Where(s => !s.IsArchive).ToList();

                cmbSourceStore.DataSource = sourceList;
                cmbSourceStore.DisplayMember = "DisplayName";
                cmbSourceStore.ValueMember = "Store";
                cmbSourceStore.Enabled = sourceList.Count > 0;

                cmbTargetStore.DataSource = targetList;
                cmbTargetStore.DisplayMember = "DisplayName";
                cmbTargetStore.ValueMember = "Store";
                cmbTargetStore.Enabled = targetList.Count > 0;

                btnOK.Enabled = sourceList.Count > 0 && targetList.Count > 0;

                if (sourceList.Count == 0)
                {
                    lblLoading.Text = "未检测到联机存档数据文件";
                    lblLoading.ForeColor = System.Drawing.Color.Red;
                }
                else if (targetList.Count == 0)
                {
                    lblLoading.Text = "未检测到本地 PST 数据文件";
                    lblLoading.ForeColor = System.Drawing.Color.Red;
                }
                else
                {
                    lblLoading.Text = $"已加载 {storeList.Count} 个数据文件";
                    lblLoading.ForeColor = System.Drawing.Color.Green;
                    progressLoading.Visible = false;
                }
            }
            catch (System.Exception ex)
            {
                lblLoading.Text = $"加载失败: {ex.Message}";
                lblLoading.ForeColor = System.Drawing.Color.Red;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (cmbSourceStore.SelectedItem == null || cmbTargetStore.SelectedItem == null)
            {
                MessageBox.Show("请选择源数据文件和目标数据文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var sourceStoreInfo = (StoreInfo)cmbSourceStore.SelectedItem;
            var targetStoreInfo = (StoreInfo)cmbTargetStore.SelectedItem;

            try
            {
                SourceRootFolder = sourceStoreInfo.Store.GetRootFolder();
                TargetRootFolder = targetStoreInfo.Store.GetRootFolder();

                if (SourceRootFolder == null || TargetRootFolder == null)
                {
                    MessageBox.Show("无法获取所选数据文件的根文件夹。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 检查是否是同一个数据文件
                if (sourceStoreInfo.Store.StoreID == targetStoreInfo.Store.StoreID)
                {
                    MessageBox.Show("源数据文件和目标数据文件不能相同。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"获取数据文件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private class StoreInfo
        {
            public Outlook.Store Store { get; set; }
            public string DisplayName { get; set; }
            public bool IsArchive { get; set; }
        }
    }

    /// <summary>
    /// 文件夹差异分析窗体
    /// </summary>
    public class FolderDiffForm : Form
    {
        private List<FolderDiffInfo> _folderDiffs;
        private ListView lvFolders;
        private Button btnOK;
        private Button btnCancel;
        private Button btnSelectAll;
        private Button btnDeselectAll;

        public List<FolderDiffInfo> SelectedFolders { get; private set; }

        public FolderDiffForm(List<FolderDiffInfo> folderDiffs)
        {
            _folderDiffs = folderDiffs;
            SelectedFolders = new List<FolderDiffInfo>();
            InitializeComponent();
            LoadFolderDiffs();
        }

        private void InitializeComponent()
        {
            this.Text = "文件夹差异分析";
            this.Width = 800;
            this.Height = 550;
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

            // 标题面板
            var titlePanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 55
            };

            var titleLabel = new Label
            {
                Text = $"发现 {_folderDiffs.Count} 个文件夹有差异，请选择要同步的文件夹",
                Dock = DockStyle.Top,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 11, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 30
            };

            var hintLabel = new Label
            {
                Text = "勾选要同步的文件夹，差异数为正表示联机存档比本地多的邮件数量",
                Dock = DockStyle.Bottom,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                ForeColor = System.Drawing.Color.Gray,
                Height = 25
            };

            titlePanel.Controls.Add(titleLabel);
            titlePanel.Controls.Add(hintLabel);

            // 按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 45
            };

            // ListView 表格
            lvFolders = new ListView
            {
                Dock = DockStyle.Fill,
                View = System.Windows.Forms.View.Details,
                CheckBoxes = true,
                FullRowSelect = true,
                GridLines = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9)
            };

            // 添加列
            lvFolders.Columns.Add("选择", 50, HorizontalAlignment.Center);
            lvFolders.Columns.Add("文件夹路径", 350, HorizontalAlignment.Left);
            lvFolders.Columns.Add("联机存档", 90, HorizontalAlignment.Right);
            lvFolders.Columns.Add("本地PST", 90, HorizontalAlignment.Right);
            lvFolders.Columns.Add("差异数(待同步)", 110, HorizontalAlignment.Right);

            // 统计信息标签
            int totalDiffCount = _folderDiffs.Sum(f => f.DiffCount);
            var statsLabel = new Label
            {
                Text = $"总计: {_folderDiffs.Count} 个文件夹, {totalDiffCount} 封邮件待同步",
                Dock = DockStyle.Bottom,
                TextAlign = System.Drawing.ContentAlignment.MiddleRight,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 25
            };

            btnSelectAll = new Button
            {
                Text = "全选",
                Width = 70,
                Height = 28,
                Left = 10,
                Top = 8
            };
            btnSelectAll.Click += (s, e) =>
            {
                foreach (ListViewItem item in lvFolders.Items)
                    item.Checked = true;
            };

            btnDeselectAll = new Button
            {
                Text = "全不选",
                Width = 70,
                Height = 28,
                Left = 90,
                Top = 8
            };
            btnDeselectAll.Click += (s, e) =>
            {
                foreach (ListViewItem item in lvFolders.Items)
                    item.Checked = false;
            };

            btnOK = new Button
            {
                Text = "开始同步",
                Width = 90,
                Height = 28,
                Left = 580,
                Top = 8
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Width = 70,
                Height = 28,
                Left = 680,
                Top = 8,
                DialogResult = DialogResult.Cancel
            };

            buttonPanel.Controls.Add(btnSelectAll);
            buttonPanel.Controls.Add(btnDeselectAll);
            buttonPanel.Controls.Add(btnOK);
            buttonPanel.Controls.Add(btnCancel);

            tableLayout.Controls.Add(titlePanel, 0, 0);
            tableLayout.Controls.Add(lvFolders, 0, 1);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            this.Controls.Add(tableLayout);
            this.Controls.Add(statsLabel);
            this.Controls.Add(buttonPanel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadFolderDiffs()
        {
            lvFolders.Items.Clear();
            foreach (var diff in _folderDiffs)
            {
                var item = new ListViewItem("");  // 复选框列
                item.SubItems.Add(diff.FolderPath);
                item.SubItems.Add(diff.SourceCount.ToString("N0"));
                item.SubItems.Add(diff.TargetCount.ToString("N0"));
                item.SubItems.Add(diff.DiffCount.ToString("N0"));
                item.Checked = true;
                item.Tag = diff;  // 存储 FolderDiffInfo 对象
                lvFolders.Items.Add(item);
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SelectedFolders.Clear();
            foreach (ListViewItem item in lvFolders.Items)
            {
                if (item.Checked)
                {
                    SelectedFolders.Add((FolderDiffInfo)item.Tag);
                }
            }

            if (SelectedFolders.Count == 0)
            {
                MessageBox.Show("请至少选择一个文件夹进行同步。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int totalEmails = SelectedFolders.Sum(f => f.DiffCount);
            var result = MessageBox.Show($"已选择 {SelectedFolders.Count} 个文件夹，共 {totalEmails} 封邮件需要同步。\n\n是否开始同步？", 
                "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
    }

    /// <summary>
    /// 分析进度窗体
    /// </summary>
    public class AnalysisProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblProgress;

        public AnalysisProgressForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "jimsoutlooktools - 正在分析文件夹差异";
            this.Width = 500;
            this.Height = 180;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(20)
            };

            // 状态标签
            lblStatus = new Label
            {
                Text = "正在初始化...",
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
                Value = 0,
                Height = 25
            };

            // 进度文字标签
            lblProgress = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                ForeColor = System.Drawing.Color.Gray,
                Height = 25
            };

            tableLayout.Controls.Add(lblStatus, 0, 0);
            tableLayout.Controls.Add(progressBar, 0, 1);
            tableLayout.Controls.Add(lblProgress, 0, 2);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            this.Controls.Add(tableLayout);
        }

        public void UpdateStatus(string status)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<string>(UpdateStatus), status);
                return;
            }
            lblStatus.Text = status;
            System.Windows.Forms.Application.DoEvents();
        }

        public void UpdateProgress(int current, int total, int percent)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<int, int, int>(UpdateProgress), current, total, percent);
                return;
            }
            progressBar.Value = Math.Min(percent, 100);
            lblProgress.Text = $"{current} / {total} ({percent}%)";
            System.Windows.Forms.Application.DoEvents();
        }

        public void Complete(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<string>(Complete), message);
                return;
            }
            lblStatus.Text = message;
            lblStatus.ForeColor = System.Drawing.Color.Green;
            progressBar.Value = 100;
            System.Windows.Forms.Application.DoEvents();
            System.Threading.Thread.Sleep(500); // 显示完成状态500ms
            this.Close();
        }
    }

    /// <summary>
    /// 同步进度窗体
    /// </summary>
    public class SyncProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblCurrentFolder;
        private Button btnCancel;

        public bool IsCancelled { get; private set; }
        private int _totalEmails;

        public SyncProgressForm(string appVersion, int totalEmails)
        {
            _totalEmails = totalEmails;
            IsCancelled = false;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = $"jimsoutlooktools - 正在同步邮件";
            this.Width = 500;
            this.Height = 200;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(20)
            };

            // 状态标签
            lblStatus = new Label
            {
                Text = $"准备同步 {_totalEmails} 封邮件...",
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
                Maximum = _totalEmails,
                Value = 0,
                Height = 25
            };

            // 当前文件夹标签
            lblCurrentFolder = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                ForeColor = System.Drawing.Color.Gray,
                Height = 25
            };

            // 取消按钮
            btnCancel = new Button
            {
                Text = "取消同步",
                Width = 100,
                Height = 30,
                Anchor = AnchorStyles.None
            };
            btnCancel.Click += (s, e) =>
            {
                IsCancelled = true;
                lblStatus.Text = "正在取消，请稍候...";
            };

            var buttonPanel = new Panel { Height = 40 };
            buttonPanel.Controls.Add(btnCancel);
            btnCancel.Left = (buttonPanel.Width - btnCancel.Width) / 2;

            tableLayout.Controls.Add(lblStatus, 0, 0);
            tableLayout.Controls.Add(progressBar, 0, 1);
            tableLayout.Controls.Add(lblCurrentFolder, 0, 2);
            tableLayout.Controls.Add(buttonPanel, 0, 3);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            this.Controls.Add(tableLayout);
        }

        public void UpdateProgress(int processed, int total, string currentFolder)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<int, int, string>(UpdateProgress), processed, total, currentFolder);
                return;
            }

            progressBar.Value = Math.Min(processed, total);
            int percent = total > 0 ? (int)((double)processed / total * 100) : 0;
            lblStatus.Text = $"已处理: {processed} / {total} ({percent}%)";
            lblCurrentFolder.Text = $"当前: {currentFolder}";
        }

        public void Complete()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action(Complete));
                return;
            }

            btnCancel.Text = "确定";
            btnCancel.Click -= (s, e) => { };
            btnCancel.Click += (s, e) => this.Close();
            this.ControlBox = true;
        }
    }

    #endregion
}
