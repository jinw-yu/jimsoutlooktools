using System;
using System.Collections.Generic;
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
        private const string AppVersion = "v1.0.6";

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
            this.Width = 667;  // 扩大1/3: 500 * 4/3 ≈ 667
            this.Height = 427;  // 扩大1/3: 320 * 4/3 ≈ 427
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(20)  // 增加内边距
            };

            // 标题
            var titleLabel = new Label
            {
                Text = "选择源数据文件和目标数据文件",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 35
            };

            // 加载状态面板
            var loadingPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 50
            };

            lblLoading = new Label
            {
                Text = "正在加载数据文件列表，请稍候...",
                Left = 0,
                Top = 0,
                Width = 600,
                Height = 25,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                ForeColor = System.Drawing.Color.SteelBlue
            };

            progressLoading = new ProgressBar
            {
                Left = 0,
                Top = 28,
                Width = 600,
                Height = 18,
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
                Height = 30
            };

            cmbSourceStore = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 28,
                Enabled = false
            };

            // 目标数据文件（本地PST）
            var targetLabel = new Label
            {
                Text = "目标数据文件（本地PST）:",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold),
                Height = 30
            };

            cmbTargetStore = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 28,
                Enabled = false
            };

            // 按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };

            btnOK = new Button
            {
                Text = "下一步",
                Width = 90,
                Height = 32,
                Left = 380,
                Top = 8,
                Enabled = false
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Width = 90,
                Height = 32,
                Left = 490,
                Top = 8,
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
                AutoSize = true,
                MinimumSize = new System.Drawing.Size(0, 30)
            };

            tableLayout.Controls.Add(titleLabel, 0, 0);
            tableLayout.Controls.Add(loadingPanel, 0, 1);
            tableLayout.Controls.Add(sourceLabel, 0, 2);
            tableLayout.Controls.Add(cmbSourceStore, 0, 3);
            tableLayout.Controls.Add(targetLabel, 0, 4);
            tableLayout.Controls.Add(cmbTargetStore, 0, 5);
            tableLayout.Controls.Add(hintLabel, 0, 6);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 55));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

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
    /// 分析进度窗体（带日志显示）
    /// </summary>
    public class AnalysisProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblProgress;
        private TextBox txtLog;
        private Button btnCopyLog;
        private Button btnClose;
        private StringBuilder logBuilder;

        public AnalysisProgressForm()
        {
            logBuilder = new StringBuilder();
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "正在分析文件夹差异";
            this.Width = 700;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = false;

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };

            // 顶部面板：状态和进度
            var topPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 80
            };

            // 状态标签
            lblStatus = new Label
            {
                Text = "正在初始化...",
                Dock = DockStyle.Top,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10),
                Height = 25
            };

            // 进度条
            progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Height = 25,
                Margin = new Padding(0, 5, 0, 5)
            };

            // 进度文字标签
            lblProgress = new Label
            {
                Text = "",
                Dock = DockStyle.Top,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                ForeColor = System.Drawing.Color.Gray,
                Height = 25
            };

            topPanel.Controls.Add(lblProgress);
            topPanel.Controls.Add(progressBar);
            topPanel.Controls.Add(lblStatus);

            // 日志文本框
            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new System.Drawing.Font("Consolas", 9),
                BackColor = System.Drawing.Color.WhiteSmoke,
                ReadOnly = true,
                Text = "=== 分析日志 ===\r\n"
            };

            // 底部按钮面板 - 使用Panel + FlowLayoutPanel对齐按钮
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };

            var flowPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                Width = 200,
                Height = 50,
                Padding = new Padding(0, 9, 10, 0)
            };

            btnCopyLog = new Button
            {
                Text = "复制日志",
                Width = 85,
                Height = 32,
                Margin = new Padding(0, 0, 10, 0)
            };
            btnCopyLog.Click += (s, e) =>
            {
                Clipboard.SetText(txtLog.Text);
                MessageBox.Show("日志已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            btnClose = new Button
            {
                Text = "下一步",
                Width = 85,
                Height = 32,
                Margin = new Padding(0, 0, 0, 0),
                Enabled = false  // 分析期间不可用
            };
            btnClose.Click += (s, e) => { this.DialogResult = DialogResult.OK; this.Close(); };

            flowPanel.Controls.Add(btnCopyLog);
            flowPanel.Controls.Add(btnClose);
            buttonPanel.Controls.Add(flowPanel);

            // 添加到主布局
            mainLayout.Controls.Add(topPanel, 0, 0);
            mainLayout.Controls.Add(txtLog, 0, 1);
            mainLayout.Controls.Add(buttonPanel, 0, 2);

            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            this.Controls.Add(mainLayout);
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

        public void AddLog(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<string>(AddLog), message);
                return;
            }
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string logLine = $"[{timestamp}] {message}\r\n";
            logBuilder.Append(logLine);
            txtLog.AppendText(logLine);

            // 确保滚动到最后一行
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.SelectionLength = 0;
            txtLog.ScrollToCaret();
            txtLog.Refresh();

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
            btnClose.Enabled = true;  // 分析完成，启用"下一步"按钮
            btnClose.Text = "下一步";  // 确保文本正确
            this.ControlBox = true; // 允许关闭窗口
            AddLog($"=== {message} ===");
            System.Windows.Forms.Application.DoEvents();
        }

        public string GetLogText()
        {
            return logBuilder.ToString();
        }

        /// <summary>
        /// 等待用户手动关闭窗体
        /// </summary>
        public void WaitForClose()
        {
            // 启用关闭按钮
            btnClose.Enabled = true;
            
            // 使用ShowDialog模式等待用户关闭
            // 由于在分析前已经Show()过了，这里我们用循环等待的方式
            while (this.Visible)
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
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
        private TextBox txtLog;
        private Button btnCancel;
        private Button btnCopyLog;
        private const int MAX_LOG_LINES = 1000;

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
            this.Text = "正在同步邮件";
            this.Width = 700;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(15)
            };

            // 顶部面板：状态和进度
            var topPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 80
            };

            // 状态标签
            lblStatus = new Label
            {
                Text = $"准备同步 {_totalEmails} 封邮件...",
                Dock = DockStyle.Top,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10),
                Height = 25
            };

            // 进度条
            progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Minimum = 0,
                Maximum = _totalEmails,
                Value = 0,
                Height = 25,
                Margin = new Padding(0, 5, 0, 5)
            };

            // 当前文件夹标签
            lblCurrentFolder = new Label
            {
                Text = "",
                Dock = DockStyle.Top,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                ForeColor = System.Drawing.Color.Gray,
                Height = 25
            };

            topPanel.Controls.Add(lblCurrentFolder);
            topPanel.Controls.Add(progressBar);
            topPanel.Controls.Add(lblStatus);

            // 日志文本框
            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new System.Drawing.Font("Consolas", 9),
                BackColor = System.Drawing.Color.WhiteSmoke,
                ReadOnly = true,
                Text = "=== 同步日志 ===\r\n"
            };

            // 底部按钮面板
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };

            var flowPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                Width = 250,
                Height = 50,
                Padding = new Padding(0, 8, 10, 0)
            };

            btnCopyLog = new Button
            {
                Text = "复制日志",
                Width = 90,
                Height = 32,
                Margin = new Padding(0, 0, 10, 0)
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

            btnCancel = new Button
            {
                Text = "取消同步",
                Width = 90,
                Height = 32,
                Margin = new Padding(0, 0, 0, 0)
            };
            btnCancel.Click += (s, e) =>
            {
                IsCancelled = true;
                lblStatus.Text = "正在取消，请稍候...";
                AddLog("用户请求取消同步...");
            };

            flowPanel.Controls.Add(btnCopyLog);
            flowPanel.Controls.Add(btnCancel);
            buttonPanel.Controls.Add(flowPanel);

            tableLayout.Controls.Add(topPanel, 0, 0);
            tableLayout.Controls.Add(txtLog, 0, 1);
            tableLayout.Controls.Add(buttonPanel, 0, 2);

            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 85));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            this.Controls.Add(tableLayout);
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

            // 添加新日志
            txtLog.AppendText(logLine);

            // 限制日志行数不超过1000行
            var lines = txtLog.Lines;
            if (lines.Length > MAX_LOG_LINES)
            {
                // 保留最新的1000行
                var newLines = new string[MAX_LOG_LINES];
                Array.Copy(lines, lines.Length - MAX_LOG_LINES, newLines, 0, MAX_LOG_LINES);
                txtLog.Lines = newLines;
            }

            // 自动滚动到最后一行
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.SelectionLength = 0;
            txtLog.ScrollToCaret();
            txtLog.Refresh();
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

            btnCancel.Text = "关闭";
            btnCancel.Click -= (s, e) => { };
            btnCancel.Click += (s, e) => this.Close();
            this.ControlBox = true;
            AddLog("同步完成！");
        }
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
                Text = "版本 v1.0.7",
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
                Text = "Outlook 邮件附件管理和数据同步工具\n帮助您高效管理邮件附件和同步数据",
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
    /// 下载联机向导窗口 - 上下分栏布局
    /// </summary>
    public class DownloadOnlineWizardForm : Form
    {
        private const int MAX_LOG_LINES = 1000;

        // 向导步骤
        private enum WizardStep
        {
            SelectDataFiles = 0,    // 选择数据文件
            Analyzing = 1,          // 分析差异
            SelectFolders = 2,      // 选择同步文件夹
            Syncing = 3             // 同步中
        }

        private WizardStep currentStep = WizardStep.SelectDataFiles;

        // 上部分控件
        private Panel wizardPanel;
        private Label lblStepTitle;
        private Button btnPrevious;
        private Button btnNext;
        private Button btnCancel;

        // 步骤1：选择数据文件
        private ComboBox cmbSourceStore;
        private ComboBox cmbTargetStore;
        private ProgressBar progressLoading;
        private Label lblLoading;

        // 步骤2：分析差异（进度条）
        private ProgressBar progressAnalyze;
        private Label lblAnalyzeStatus;

        // 步骤3：选择文件夹
        private ListView lvFolders;
        private Button btnSelectAll;
        private Button btnDeselectAll;
        private Label lblFolderStats;

        // 步骤4：同步进度
        private ProgressBar progressSync;
        private Label lblSyncStatus;

        // 下部分：日志区域
        private TextBox txtLog;
        private Button btnCopyLog;

        // 数据
        private Outlook.Application _application;
        private List<StoreInfo> _sourceStores;
        private List<StoreInfo> _targetStores;
        private Outlook.MAPIFolder _sourceRoot;
        private Outlook.MAPIFolder _targetRoot;
        private List<FolderDiffInfo> _folderDiffs;
        private List<FolderDiffInfo> _selectedFolders;

        public bool IsCancelled { get; private set; }

        public DownloadOnlineWizardForm(Outlook.Application application)
        {
            _application = application;
            _sourceStores = new List<StoreInfo>();
            _targetStores = new List<StoreInfo>();
            _folderDiffs = new List<FolderDiffInfo>();
            _selectedFolders = new List<FolderDiffInfo>();
            IsCancelled = false;

            InitializeComponent();

            // 使用 Load 事件，确保窗口先显示出来再加载数据
            this.Load += (s, e) =>
            {
                // 延迟100毫秒再加载数据，确保窗口完全显示
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
            this.Text = "下载联机存档 - 向导";
            this.Width = 900;
            this.Height = 700;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;  // 允许最大化
            this.MinimizeBox = true;  // 允许最小化
            this.ShowInTaskbar = true;  // 显示在任务栏

            // 添加窗口大小改变事件处理
            this.Resize += (s, e) =>
            {
                // 窗口大小改变时，确保日志滚动到底部
                if (txtLog != null && txtLog.IsHandleCreated)
                {
                    this.BeginInvoke(new System.Action(() =>
                    {
                        txtLog.SelectionStart = txtLog.Text.Length;
                        txtLog.SelectionLength = 0;
                        txtLog.ScrollToCaret();
                        txtLog.Refresh();
                    }));
                }
            };

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(10)
            };

            // 上部分：向导区域
            wizardPanel = new Panel
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

            btnPrevious = new Button
            {
                Text = "上一步",
                Width = 90,
                Height = 32,
                Left = 10,
                Top = 9,
                Enabled = false
            };
            btnPrevious.Click += BtnPrevious_Click;

            btnNext = new Button
            {
                Text = "下一步",
                Width = 90,
                Height = 32,
                Left = 110,
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
                    // 同步过程中取消
                    var result = MessageBox.Show(
                        "确定要取消同步吗？\n\n已同步的邮件将保留，未同步的邮件将不会继续同步。",
                        "确认取消",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        IsCancelled = true;
                        AddLog("用户取消同步操作");
                    }
                }
                else
                {
                    // 其他步骤直接关闭
                    IsCancelled = true;
                    this.Close();
                }
            };

            navPanel.Controls.Add(btnPrevious);
            navPanel.Controls.Add(btnNext);
            navPanel.Controls.Add(btnCancel);

            wizardPanel.Controls.Add(lblStepTitle);
            wizardPanel.Controls.Add(navPanel);
            
            // 添加一个占位符 Panel，用于后续替换为内容面板
            var contentPlaceholder = new Panel
            {
                Dock = DockStyle.Fill,
                Visible = false
            };
            wizardPanel.Controls.Add(contentPlaceholder);

            // 下部分：日志区域
            var logPanel = new Panel
            {
                Dock = DockStyle.Fill,
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
                Text = "=== 操作日志 ===\r\n"
            };

            var logButtonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };

            btnCopyLog = new Button
            {
                Text = "复制日志",
                Width = 90,
                Height = 28,
                Left = 10,
                Top = 6
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
            logPanel.Controls.Add(logButtonPanel);
            logPanel.Controls.Add(txtLog);
            logPanel.Controls.Add(logTitleLabel);

            mainLayout.Controls.Add(wizardPanel, 0, 0);
            mainLayout.Controls.Add(logPanel, 0, 1);

            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 40));

            this.Controls.Add(mainLayout);

            ShowStep(WizardStep.SelectDataFiles);
        }

        private void ShowStep(WizardStep step)
        {
            currentStep = step;

            AddLog($"[DEBUG] ShowStep called: {step}");

            // 移除步骤三特有的按钮（全选、全不选）
            var navPanel = wizardPanel.Controls[1] as Panel;
            if (navPanel != null)
            {
                var buttonsToRemove = new List<Control>();
                foreach (Control ctrl in navPanel.Controls)
                {
                    if (ctrl == btnSelectAll || ctrl == btnDeselectAll)
                        buttonsToRemove.Add(ctrl);
                }
                foreach (var ctrl in buttonsToRemove)
                    navPanel.Controls.Remove(ctrl);
            }

            // 移除旧的 contentPanel（占位符索引为2）
            if (wizardPanel.Controls.Count > 2)
            {
                var oldContent = wizardPanel.Controls[2];
                if (oldContent != null && oldContent is Panel)
                {
                    wizardPanel.Controls.Remove(oldContent);
                    oldContent.Dispose();
                    AddLog($"[DEBUG] Removed old content panel");
                }
            }

            AddLog($"[DEBUG] wizardPanel.Controls.Count after removal: {wizardPanel.Controls.Count}");

            // 根据步骤显示不同内容
            switch (step)
            {
                case WizardStep.SelectDataFiles:
                    ShowSelectDataFilesStep();
                    break;
                case WizardStep.Analyzing:
                    ShowAnalyzingStep();
                    break;
                case WizardStep.SelectFolders:
                    ShowSelectFoldersStep();
                    break;
                case WizardStep.Syncing:
                    ShowSyncingStep();
                    break;
            }

            AddLog($"[DEBUG] wizardPanel.Controls.Count after adding content: {wizardPanel.Controls.Count}");

            AddLog($"[DEBUG] Step content shown");

            // 更新导航按钮状态
            btnPrevious.Enabled = (step != WizardStep.SelectDataFiles && step != WizardStep.Syncing);
            // btnNext 的启用状态由各步骤自己控制，不要在这里覆盖
            // btnCancel 在同步过程中保持可用，让用户可以随时取消
            btnCancel.Enabled = true;
            if (step == WizardStep.Syncing)
            {
                btnCancel.Text = "取消同步";
            }
            else
            {
                btnCancel.Text = "取消";
            }

            AddLog($"[DEBUG] ShowStep completed");
        }

        private void ShowSelectDataFilesStep()
        {
            lblStepTitle.Text = "步骤 1/4: 选择数据文件";

            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 10, 20, 20)  // 左, 上, 右, 下
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
                Text = "提示：源数据文件是Office 365联机存档，目标数据文件是本地PST文件",
                Dock = DockStyle.Bottom,
                Height = 30,
                ForeColor = System.Drawing.Color.Gray,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8)
            };

            contentPanel.Controls.Add(hintLabel);
            contentPanel.Controls.Add(cmbTargetStore);
            contentPanel.Controls.Add(targetLabel);
            contentPanel.Controls.Add(cmbSourceStore);
            contentPanel.Controls.Add(sourceLabel);
            contentPanel.Controls.Add(loadingPanel);

            // 替换占位符 Panel（索引为2）
            if (wizardPanel.Controls.Count > 2)
            {
                var placeholder = wizardPanel.Controls[2];
                wizardPanel.Controls.Remove(placeholder);
                placeholder.Dispose();
            }
            wizardPanel.Controls.Add(contentPanel);

            // 初始化按钮状态
            btnNext.Enabled = false;
        }

        private void ShowAnalyzingStep()
        {
            lblStepTitle.Text = "步骤 2/4: 分析文件夹差异";

            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 10, 20, 20)  // 左, 上, 右, 下
            };

            lblAnalyzeStatus = new Label
            {
                Text = "正在分析...",
                Dock = DockStyle.Top,
                Height = 30,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10)
            };

            progressAnalyze = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 30,
                Minimum = 0,
                Maximum = 100
            };

            contentPanel.Controls.Add(progressAnalyze);
            contentPanel.Controls.Add(lblAnalyzeStatus);

            // 替换占位符 Panel（索引为2）
            if (wizardPanel.Controls.Count > 2)
            {
                var placeholder = wizardPanel.Controls[2];
                wizardPanel.Controls.Remove(placeholder);
                placeholder.Dispose();
            }
            wizardPanel.Controls.Add(contentPanel);

            // 分析过程中禁用下一步按钮
            btnNext.Enabled = false;
        }

        private void ShowSelectFoldersStep()
        {
            AddLog("[DEBUG] ShowSelectFoldersStep started");
            lblStepTitle.Text = "步骤 3/4: 选择要同步的文件夹";

            // 创建主内容面板
            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill
            };

            // 创建全选和全不选按钮，添加到导航面板
            btnSelectAll = new Button
            {
                Text = "全选",
                Width = 80,
                Height = 32,
                Left = 300,
                Top = 9
            };
            btnSelectAll.Click += (s, e) =>
            {
                foreach (ListViewItem item in lvFolders.Items)
                    item.Checked = true;
                UpdateFolderStats();
            };

            btnDeselectAll = new Button
            {
                Text = "全不选",
                Width = 80,
                Height = 32,
                Left = 390,
                Top = 9
            };
            btnDeselectAll.Click += (s, e) =>
            {
                foreach (ListViewItem item in lvFolders.Items)
                    item.Checked = false;
                UpdateFolderStats();
            };

            // 将按钮添加到导航面板
            var navPanel = wizardPanel.Controls[1] as Panel;
            if (navPanel != null)
            {
                navPanel.Controls.Add(btnSelectAll);
                navPanel.Controls.Add(btnDeselectAll);
            }

            // 创建 ListView
            lvFolders = new ListView
            {
                View = System.Windows.Forms.View.Details,
                CheckBoxes = true,
                FullRowSelect = true,
                GridLines = true,
                Location = new System.Drawing.Point(20, 50),  // 增加顶部间距到50
                Size = new System.Drawing.Size(840, 260)      // 相应减少高度
            };

            lvFolders.Columns.Add("选择", 50, HorizontalAlignment.Center);
            lvFolders.Columns.Add("文件夹路径", 400, HorizontalAlignment.Left);
            lvFolders.Columns.Add("联机存档", 90, HorizontalAlignment.Right);
            lvFolders.Columns.Add("本地PST", 90, HorizontalAlignment.Right);
            lvFolders.Columns.Add("差异数", 90, HorizontalAlignment.Right);

            lvFolders.ItemChecked += (s, e) => UpdateFolderStats();

            // 创建统计标签
            lblFolderStats = new Label
            {
                Location = new System.Drawing.Point(20, 320),  // 保持位置不变
                Size = new System.Drawing.Size(840, 25),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };

            // 创建表头标签
            var headerLabel = new Label
            {
                Text = "文件夹列表：",
                Location = new System.Drawing.Point(20, 5),
                Size = new System.Drawing.Size(200, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Bold)
            };

            // 添加控件到内容面板
            contentPanel.Controls.Add(headerLabel);
            contentPanel.Controls.Add(lvFolders);
            contentPanel.Controls.Add(lblFolderStats);

            // 替换占位符 Panel
            if (wizardPanel.Controls.Count > 2)
            {
                var placeholder = wizardPanel.Controls[2];
                wizardPanel.Controls.Remove(placeholder);
                placeholder.Dispose();
            }
            wizardPanel.Controls.Add(contentPanel);

            AddLog($"[DEBUG] ListView Location: {lvFolders.Location}, Size: {lvFolders.Size}");

            // 加载文件夹列表
            LoadFolderDiffs();

            AddLog("[DEBUG] ShowSelectFoldersStep completed");
        }

        private void ShowSyncingStep()
        {
            lblStepTitle.Text = "步骤 4/4: 同步邮件";

            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 10, 20, 20)  // 左, 上, 右, 下
            };

            lblSyncStatus = new Label
            {
                Text = "准备同步...",
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

            // 替换占位符 Panel（索引为2）
            if (wizardPanel.Controls.Count > 2)
            {
                var placeholder = wizardPanel.Controls[2];
                wizardPanel.Controls.Remove(placeholder);
                placeholder.Dispose();
            }
            wizardPanel.Controls.Add(contentPanel);

            // 同步过程中禁用下一步按钮
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

        private void BtnPrevious_Click(object sender, EventArgs e)
        {
            if (currentStep == WizardStep.SelectFolders)
                ShowStep(WizardStep.SelectDataFiles);
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            switch (currentStep)
            {
                case WizardStep.SelectDataFiles:
                    StartAnalysis();
                    break;
                case WizardStep.SelectFolders:
                    StartSync();
                    break;
            }
        }

        private async void StartAnalysis()
        {
            try
            {
                _sourceRoot = ((StoreInfo)cmbSourceStore.SelectedItem).Store.GetRootFolder();
                _targetRoot = ((StoreInfo)cmbTargetStore.SelectedItem).Store.GetRootFolder();

                ShowStep(WizardStep.Analyzing);
                AddLog($"开始分析文件夹差异");
                AddLog($"源数据文件: {_sourceRoot.Name}");
                AddLog($"目标数据文件: {_targetRoot.Name}");
                AddLog("");
                AddLog("正在启动分析任务，请稍候...");
                UpdateAnalyzeProgress(0, 100, 0, "正在初始化...");

                // 在后台线程执行分析
                var analysisTask = System.Threading.Tasks.Task.Run(() =>
                {
                    var result = new List<FolderDiffInfo>();

                    // 获取所有文件夹
                    this.Invoke(new System.Action(() =>
                    {
                        AddLog("正在获取源数据文件文件夹列表...");
                        UpdateAnalyzeProgress(0, 100, 5, "正在获取源数据文件文件夹列表...");
                    }));

                    var sourceFolders = GetAllFolders(_sourceRoot);
                    this.Invoke(new System.Action(() =>
                    {
                        AddLog($"[源] 找到 {sourceFolders.Count} 个文件夹");
                        AddLog("正在获取目标数据文件文件夹列表...");
                        UpdateAnalyzeProgress(0, 100, 20, $"源数据文件: {sourceFolders.Count} 个文件夹");
                    }));

                    var targetFolders = GetAllFolders(_targetRoot);
                    this.Invoke(new System.Action(() =>
                    {
                        AddLog($"[目标] 找到 {targetFolders.Count} 个文件夹");
                        AddLog("正在构建文件夹索引...");
                        UpdateAnalyzeProgress(0, 100, 40, $"目标数据文件: {targetFolders.Count} 个文件夹");
                    }));

                    // 构建目标文件夹字典
                    var targetFolderDict = new Dictionary<string, Outlook.MAPIFolder>(System.StringComparer.OrdinalIgnoreCase);
                    foreach (var folder in targetFolders)
                    {
                        string path = GetSimpleFolderPath(folder, _targetRoot);
                        if (!string.IsNullOrEmpty(path))
                            targetFolderDict[path] = folder;
                    }

                    this.Invoke(new System.Action(() =>
                    {
                        AddLog($"目标字典构建完成，包含 {targetFolderDict.Count} 个路径");
                        AddLog("");
                        AddLog("开始对比文件夹差异...");
                        UpdateAnalyzeProgress(0, 100, 50, "正在对比文件夹差异...");
                    }));

                    // 对比文件夹
                    int processedCount = 0;
                    int totalFolders = sourceFolders.Count;
                    int diffFound = 0;
                    var diffLogs = new List<string>();

                    foreach (var sourceFolder in sourceFolders)
                    {
                        processedCount++;
                        string simplePath = GetSimpleFolderPath(sourceFolder, _sourceRoot);
                        if (string.IsNullOrEmpty(simplePath))
                            continue;

                        // 在目标中查找匹配
                        Outlook.MAPIFolder targetFolder = null;
                        targetFolderDict.TryGetValue(simplePath, out targetFolder);

                        // 获取邮件数量
                        int sourceCount = GetMailItemCount(sourceFolder);
                        int targetCount = targetFolder != null ? GetMailItemCount(targetFolder) : 0;

                        // 如果源比目标多，记录差异
                        if (sourceCount > targetCount)
                        {
                            diffFound++;
                            result.Add(new FolderDiffInfo
                            {
                                FolderPath = simplePath,
                                SourceFolder = sourceFolder,
                                TargetFolder = targetFolder,
                                SourceCount = sourceCount,
                                TargetCount = targetCount,
                                DiffCount = sourceCount - targetCount
                            });

                            diffLogs.Add($"  发现差异: {simplePath} (源:{sourceCount} 目标:{targetCount} 差:{sourceCount - targetCount})");
                        }

                        // 每5个文件夹更新一次进度和日志
                        if (processedCount % 5 == 0 || processedCount == totalFolders)
                        {
                            int percent = 50 + (int)((double)processedCount / totalFolders * 50);
                            var logsToOutput = new List<string>(diffLogs);
                            diffLogs.Clear();

                            this.Invoke(new System.Action(() =>
                            {
                                UpdateAnalyzeProgress(processedCount, totalFolders, percent, $"正在分析: {simplePath}");
                                foreach (var log in logsToOutput)
                                {
                                    AddLog(log);
                                }
                            }));
                        }
                    }

                    this.Invoke(new System.Action(() =>
                    {
                        // 输出剩余的日志
                        foreach (var log in diffLogs)
                        {
                            AddLog(log);
                        }
                        AddLog($"对比完成，共检查 {processedCount} 个文件夹，发现 {diffFound} 个差异");
                    }));

                    return result;
                });

                // 等待分析完成
                _folderDiffs = await analysisTask;

                AddLog("");
                AddLog($"分析完成，发现 {_folderDiffs.Count} 个有差异的文件夹");

                if (_folderDiffs.Count == 0)
                {
                    MessageBox.Show("没有发现需要同步的文件夹差异。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ShowStep(WizardStep.SelectDataFiles);
                }
                else
                {
                    AddLog("正在切换到文件夹选择界面...");
                    ShowStep(WizardStep.SelectFolders);
                    AddLog("已切换到文件夹选择界面");
                }
            }
            catch (System.Exception ex)
            {
                AddLog($"✗ 分析失败: {ex.Message}");
                MessageBox.Show($"分析失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowStep(WizardStep.SelectDataFiles);
            }
        }

        private List<Outlook.MAPIFolder> GetAllFolders(Outlook.MAPIFolder rootFolder)
        {
            var result = new List<Outlook.MAPIFolder>();
            GetFoldersRecursive(rootFolder, result);
            return result;
        }

        private void GetFoldersRecursive(Outlook.MAPIFolder folder, List<Outlook.MAPIFolder> folderList)
        {
            folderList.Add(folder);

            try
            {
                foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                {
                    GetFoldersRecursive(subFolder, folderList);
                }
            }
            catch
            {
                // 忽略无法访问的文件夹
            }
        }

        private string GetSimpleFolderPath(Outlook.MAPIFolder folder, Outlook.MAPIFolder rootFolder)
        {
            if (folder == null)
                return null;

            if (folder.EntryID == rootFolder.EntryID)
                return "__ROOT__";

            try
            {
                var pathParts = new List<string>();
                Outlook.MAPIFolder current = folder;

                while (current != null && current.EntryID != rootFolder.EntryID)
                {
                    pathParts.Insert(0, current.Name);
                    current = current.Parent as Outlook.MAPIFolder;
                }

                return string.Join("\\", pathParts);
            }
            catch
            {
                return folder.Name;
            }
        }

        private int GetMailItemCount(Outlook.MAPIFolder folder)
        {
            try
            {
                // 直接使用 Items.Count，速度快得多
                // 注意：这包含所有项目（邮件、日历、联系人等）
                // 但对于邮件文件夹，大部分都是邮件
                return folder.Items.Count;
            }
            catch
            {
                return 0;
            }
        }

        private void LoadFolderDiffs()
        {
            lvFolders.Items.Clear();
            foreach (var diff in _folderDiffs)
            {
                var item = new ListViewItem("");
                item.SubItems.Add(diff.FolderPath);
                item.SubItems.Add(diff.SourceCount.ToString("N0"));
                item.SubItems.Add(diff.TargetCount.ToString("N0"));
                item.SubItems.Add(diff.DiffCount.ToString("N0"));
                item.Checked = true;
                item.Tag = diff;
                lvFolders.Items.Add(item);
            }
            UpdateFolderStats();
            btnNext.Enabled = (_folderDiffs.Count > 0);
        }

        private void UpdateFolderStats()
        {
            int selectedCount = lvFolders.Items.Cast<ListViewItem>().Count(i => i.Checked);
            int totalEmails = lvFolders.Items.Cast<ListViewItem>()
                .Where(i => i.Checked)
                .Sum(i => ((FolderDiffInfo)i.Tag).DiffCount);

            lblFolderStats.Text = $"已选择 {selectedCount} 个文件夹，共 {totalEmails} 封邮件待同步";
            btnNext.Enabled = (selectedCount > 0);
        }

        private void StartSync()
        {
            _selectedFolders.Clear();
            foreach (ListViewItem item in lvFolders.Items)
            {
                if (item.Checked)
                    _selectedFolders.Add((FolderDiffInfo)item.Tag);
            }

            ShowStep(WizardStep.Syncing);
            AddLog($"开始同步 {_selectedFolders.Count} 个文件夹");

            // 启动异步同步任务
            System.Threading.Tasks.Task.Run(() => PerformSync());
        }

        private async void PerformSync()
        {
            try
            {
                int totalFolders = _selectedFolders.Count;
                int currentFolder = 0;
                int totalEmails = _selectedFolders.Sum(f => f.DiffCount);
                int syncedEmails = 0;
                var syncStats = new List<FolderSyncStat>();  // 记录每个文件夹的同步统计

                AddLog($"总共需要同步 {totalEmails} 封邮件");

                foreach (var folderDiff in _selectedFolders)
                {
                    if (IsCancelled)
                        break;

                    currentFolder++;
                    var folderStat = new FolderSyncStat
                    {
                        FolderPath = folderDiff.FolderPath,
                        TotalEmails = folderDiff.DiffCount,
                        SyncedEmails = 0
                    };

                    this.Invoke(new System.Action(() =>
                    {
                        AddLog($"[{currentFolder}/{totalFolders}] 正在同步: {folderDiff.FolderPath}");
                        UpdateSyncProgress(currentFolder, totalFolders, syncedEmails, totalEmails, $"正在同步: {folderDiff.FolderPath}");
                    }));

                    // 同步单个文件夹
                    int synced = await SyncFolderAsync(folderDiff, syncedEmails, totalEmails);
                    syncedEmails += synced;
                    folderStat.SyncedEmails = synced;
                    syncStats.Add(folderStat);

                    this.Invoke(new System.Action(() =>
                    {
                        AddLog($"  已同步 {synced} 封邮件");
                    }));
                }

                this.Invoke(new System.Action(() =>
                {
                    if (IsCancelled)
                    {
                        // 显示详细的取消统计信息
                        AddLog("");
                        AddLog("=== 同步已取消 ===");
                        AddLog($"已同步: {syncedEmails} 封邮件");
                        AddLog($"未同步: {totalEmails - syncedEmails} 封邮件");
                        AddLog("");
                        AddLog("已同步的文件夹:");

                        foreach (var stat in syncStats)
                        {
                            if (stat.SyncedEmails > 0)
                            {
                                AddLog($"  ✓ {stat.FolderPath}: {stat.SyncedEmails}/{stat.TotalEmails} 封");
                            }
                        }

                        // 显示未同步的文件夹
                        var unsyncedFolders = _selectedFolders.Where(f => !syncStats.Any(s => s.FolderPath == f.FolderPath && s.SyncedEmails > 0)).ToList();
                        if (unsyncedFolders.Count > 0)
                        {
                            AddLog("");
                            AddLog("未同步的文件夹:");
                            foreach (var folder in unsyncedFolders)
                            {
                                AddLog($"  ✗ {folder.FolderPath}: 0/{folder.DiffCount} 封");
                            }
                        }

                        // 显示提示框
                        MessageBox.Show(
                            $"同步已取消\n\n" +
                            $"已同步: {syncedEmails} 封邮件\n" +
                            $"未同步: {totalEmails - syncedEmails} 封邮件\n\n" +
                            $"已同步的邮件已保留在目标文件夹中。",
                            "同步已取消",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        AddLog($"同步完成，共同步 {syncedEmails} 封邮件");
                        UpdateSyncProgress(totalFolders, totalFolders, totalEmails, totalEmails, "同步完成");
                    }

                    // 更新按钮状态
                    btnNext.Text = "完成";
                    btnNext.Enabled = true;
                    btnNext.Click -= BtnNext_Click;
                    btnNext.Click += (s, e) => this.Close();
                }));
            }
            catch (System.Exception ex)
            {
                this.Invoke(new System.Action(() =>
                {
                    AddLog($"✗ 同步失败: {ex.Message}");
                    MessageBox.Show($"同步失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            }
        }

        // 文件夹同步统计类
        private class FolderSyncStat
        {
            public string FolderPath { get; set; }
            public int TotalEmails { get; set; }
            public int SyncedEmails { get; set; }
        }

        private async System.Threading.Tasks.Task<int> SyncFolderAsync(FolderDiffInfo folderDiff, int alreadySynced, int totalEmails)
        {
            return await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    int syncedCount = 0;
                    var sourceFolder = folderDiff.SourceFolder;
                    var targetFolder = folderDiff.TargetFolder;

                    // 如果目标文件夹不存在，创建它
                    if (targetFolder == null)
                    {
                        this.Invoke(new System.Action(() =>
                        {
                            AddLog($"  目标文件夹不存在，需要创建: {folderDiff.FolderPath}");
                        }));
                        return 0;
                    }

                    // 使用已知的差异数作为目标
                    int targetCount = folderDiff.DiffCount;
                    int totalSourceCount = folderDiff.SourceCount;  // 总邮件数

                    // 先添加一行占位符，后续更新这一行
                    this.Invoke(new System.Action(() =>
                    {
                        AddLog($"  开始逐封检查并同步，共需对比 {totalSourceCount} 封邮件，目标同步 {targetCount} 封邮件...");
                        AddLog($"  对比进度: 0/{totalSourceCount} - 已同步: 0/{targetCount}");  // 占位符行
                    }));

                    // 使用 Items 集合的 GetFirst/GetNext 方法
                    var sourceItems = sourceFolder.Items;
                    var targetItems = targetFolder.Items;

                    int current = 0;
                    int lastProgressUpdate = 0;
                    object sourceItem = null;

                    try
                    {
                        sourceItem = sourceItems.GetFirst();

                        while (sourceItem != null && syncedCount < targetCount)
                        {
                            if (IsCancelled)
                            {
                                this.Invoke(new System.Action(() =>
                                {
                                    // 更新进度行
                                    UpdateProgressLine(current, totalSourceCount, syncedCount, targetCount);
                                    AddLog($"  用户取消，已同步 {syncedCount}/{targetCount} 封邮件");
                                }));
                                break;
                            }

                            current++;

                            // 每10封邮件更新一次进度行（避免UI卡顿）
                            if (current % 10 == 0)
                            {
                                this.Invoke(new System.Action(() =>
                                {
                                    UpdateProgressLine(current, totalSourceCount, syncedCount, targetCount);
                                }));
                            }

                            // 每100封邮件更新一次进度条
                            if (current - lastProgressUpdate >= 100)
                            {
                                lastProgressUpdate = current;
                                int totalSynced = alreadySynced + syncedCount;
                                this.Invoke(new System.Action(() =>
                                {
                                    UpdateSyncProgress(0, 0, totalSynced, totalEmails, $"正在检查: {folderDiff.FolderPath} ({current}/{totalSourceCount})");
                                }));
                            }

                            if (sourceItem is Outlook.MailItem sourceMail)
                            {
                                bool shouldSync = false;

                                try
                                {
                                    // 使用邮件的 EntryID 和主题作为唯一标识
                                    string subject = sourceMail.Subject ?? "";
                                    DateTime receivedTime = sourceMail.ReceivedTime;

                                    // 检查目标文件夹中是否存在相同主题和接收时间的邮件
                                    bool exists = false;

                                    try
                                    {
                                        // 使用过滤器查找相同主题的邮件
                                        var filter = $"[Subject] = '{subject.Replace("'", "''")}'";
                                        var filteredItems = targetItems.Restrict(filter);

                                        if (filteredItems.Count > 0)
                                        {
                                            // 检查接收时间是否匹配
                                            object filteredItem = filteredItems.GetFirst();
                                            while (filteredItem != null)
                                            {
                                                if (filteredItem is Outlook.MailItem targetMail)
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

                                                var nextFiltered = filteredItems.GetNext();
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItem);
                                                filteredItem = nextFiltered;
                                            }
                                        }

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
                                    }
                                    catch
                                    {
                                        // 如果过滤失败，假设不存在
                                        exists = false;
                                    }

                                    shouldSync = !exists;

                                    if (shouldSync)
                                    {
                                        // 复制邮件到目标文件夹
                                        var copiedMail = sourceMail.Copy();
                                        copiedMail.Move(targetFolder);

                                        // 释放复制的邮件对象
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedMail);

                                        syncedCount++;

                                        // 每10封邮件更新一次进度
                                        if (syncedCount % 10 == 0)
                                        {
                                            int totalSynced = alreadySynced + syncedCount;
                                            this.Invoke(new System.Action(() =>
                                            {
                                                UpdateSyncProgress(0, 0, totalSynced, totalEmails, $"正在同步: {folderDiff.FolderPath} (已同步 {syncedCount}/{targetCount})");
                                            }));
                                        }
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    // 忽略单个邮件的处理错误
                                    if (syncedCount % 100 == 0)
                                    {
                                        this.Invoke(new System.Action(() =>
                                        {
                                            AddLog($"  警告: 部分邮件处理失败: {ex.Message}");
                                        }));
                                    }
                                }
                                finally
                                {
                                    // 释放源邮件 COM 对象
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMail);
                                }
                            }

                            // 获取下一封邮件
                            var nextItem = sourceItems.GetNext();

                            // 释放当前项
                            if (sourceItem != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceItem);
                            }

                            sourceItem = nextItem;
                        }
                    }
                    finally
                    {
                        // 释放最后的项
                        if (sourceItem != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceItem);
                        }

                        // 释放 items 集合
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceItems);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(targetItems);
                    }

                    return syncedCount;
                }
                catch (System.Exception ex)
                {
                    this.Invoke(new System.Action(() =>
                    {
                        AddLog($"  错误: 同步文件夹失败: {ex.Message}");
                    }));
                    return 0;
                }
            });
        }

        private void UpdateSyncProgress(int currentFolder, int totalFolders, int currentEmail, int totalEmails, string status)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<int, int, int, int, string>(UpdateSyncProgress), currentFolder, totalFolders, currentEmail, totalEmails, status);
                return;
            }

            // 使用邮件进度而不是文件夹进度
            if (totalEmails > 0)
            {
                int emailPercent = (int)((double)currentEmail / totalEmails * 100);
                progressSync.Value = emailPercent;
            }
            else if (totalFolders > 0)
            {
                // 如果没有邮件进度，使用文件夹进度
                int folderPercent = (int)((double)currentFolder / totalFolders * 100);
                progressSync.Value = folderPercent;
            }

            lblSyncStatus.Text = status;
        }

        private void UpdateProgressLine(int current, int total, int synced, int target)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<int, int, int, int>(UpdateProgressLine), current, total, synced, target);
                return;
            }

            try
            {
                // 更新最后一行日志（进度行）
                if (txtLog.Lines.Length > 0)
                {
                    string progressText = $"  对比进度: {current}/{total} - 已同步: {synced}/{target}";

                    // 使用更高效的方式更新最后一行
                    int lastLineIndex = txtLog.GetFirstCharIndexFromLine(txtLog.Lines.Length - 1);
                    if (lastLineIndex >= 0)
                    {
                        txtLog.SelectionStart = lastLineIndex;
                        txtLog.SelectionLength = txtLog.Text.Length - lastLineIndex;
                        txtLog.SelectedText = progressText;
                    }

                    // 滚动到底部
                    txtLog.SelectionStart = txtLog.Text.Length;
                    txtLog.SelectionLength = 0;
                    txtLog.ScrollToCaret();
                }
            }
            catch
            {
                // 忽略更新错误
            }
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

            // 确保滚动到最后一行
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.SelectionLength = 0;
            txtLog.ScrollToCaret();

            // 强制刷新显示
            txtLog.Refresh();

            // 使用 Windows API 确保滚动到底部
            System.Windows.Forms.Application.DoEvents();
        }

        public void UpdateAnalyzeProgress(int current, int total, int percent, string status)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<int, int, int, string>(UpdateAnalyzeProgress), current, total, percent, status);
                return;
            }

            progressAnalyze.Value = Math.Min(percent, 100);
            lblAnalyzeStatus.Text = status;
        }

        public void UpdateSyncProgress(int current, int total, int percent, string status)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action<int, int, int, string>(UpdateSyncProgress), current, total, percent, status);
                return;
            }

            progressSync.Value = Math.Min(percent, 100);
            lblSyncStatus.Text = status;
        }
    }

    #endregion

    #region 阻止域对话框

    // 确认对话框
    public class BlockDomainConfirmForm : Form
    {
        public BlockDomainConfirmForm(string domain, string logContent)
        {
            this.Text = $"JTools-outlook - 确认阻止域 *@{domain}";
            this.Width = 600;
            this.Height = 450;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 标题标签
            var lblTitle = new Label
            {
                Text = $"确认阻止域: *@{domain}",
                Dock = DockStyle.Top,
                Height = 40,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 14, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue
            };

            // 日志文本框
            var txtLog = new TextBox
            {
                Text = logContent,
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 10),
                BackColor = System.Drawing.Color.White
            };

            // 按钮面板
            var panelButtons = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60
            };

            var btnOK = new Button
            {
                Text = "确定执行",
                DialogResult = DialogResult.OK,
                Width = 120,
                Height = 35,
                Left = 150,
                Top = 12
            };

            var btnCancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Width = 120,
                Height = 35,
                Left = 320,
                Top = 12
            };

            panelButtons.Controls.Add(btnOK);
            panelButtons.Controls.Add(btnCancel);

            this.Controls.Add(txtLog);
            this.Controls.Add(lblTitle);
            this.Controls.Add(panelButtons);

            this.AcceptButton = btnCancel; // 默认取消按钮
            this.CancelButton = btnCancel;
        }
    }

    // 结果对话框
    public class BlockDomainResultForm : Form
    {
        public BlockDomainResultForm(string logContent)
        {
            this.Text = "JTools-outlook - 操作结果";
            this.Width = 600;
            this.Height = 450;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 标题标签
            var lblTitle = new Label
            {
                Text = "操作日志",
                Dock = DockStyle.Top,
                Height = 40,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 14, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue
            };

            // 日志文本框
            var txtLog = new TextBox
            {
                Text = logContent,
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 10),
                BackColor = System.Drawing.Color.White
            };

            // 按钮面板
            var panelButtons = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60
            };

            var btnCopy = new Button
            {
                Text = "复制日志",
                Width = 120,
                Height = 35,
                Left = 150,
                Top = 12
            };
            btnCopy.Click += (s, e) =>
            {
                try
                {
                    Clipboard.SetText(logContent);
                    MessageBox.Show("日志已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch { }
            };

            var btnClose = new Button
            {
                Text = "关闭",
                DialogResult = DialogResult.OK,
                Width = 120,
                Height = 35,
                Left = 320,
                Top = 12
            };

            panelButtons.Controls.Add(btnCopy);
            panelButtons.Controls.Add(btnClose);

            this.Controls.Add(txtLog);
            this.Controls.Add(lblTitle);
            this.Controls.Add(panelButtons);

            this.AcceptButton = btnClose;
            this.CancelButton = btnClose;
        }
    }

    #endregion
}
