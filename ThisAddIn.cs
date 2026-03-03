using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;

namespace jimsoutlooktools
{
    public partial class ThisAddIn
    {
        private CommandBarButton _downloadButton;
        private CommandBar _toolbar;
        private const string ToolbarName = "jimsoutlooktools";
        private const string AppVersion = "v1.0.2";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // 等待Outlook完全加载
                ((ApplicationEvents_11_Event)Application).Startup += Application_Startup;
            }
            catch
            {
                // 如果Startup事件已触发，直接创建工具栏
                CreateToolbar();
            }
        }

        private void Application_Startup()
        {
            CreateToolbar();
        }

        private void CreateToolbar()
        {
            try
            {
                if (Application.ActiveExplorer() == null)
                    return;

                CommandBars commandBars = Application.ActiveExplorer().CommandBars;

                // 如果工具栏已存在，先删除
                try
                {
                    _toolbar = commandBars[ToolbarName];
                    _toolbar.Delete();
                }
                catch { }

                // 创建新工具栏
                _toolbar = commandBars.Add(ToolbarName, MsoBarPosition.msoBarTop, false, true);

                // 添加按钮
                _downloadButton = (CommandBarButton)_toolbar.Controls.Add(
                    MsoControlType.msoControlButton,
                    System.Type.Missing,
                    System.Type.Missing,
                    1,
                    true);

                _downloadButton.Caption = "保存附件";
                _downloadButton.Style = MsoButtonStyle.msoButtonCaption;
                _downloadButton.Click += new _CommandBarButtonEvents_ClickEventHandler(DownloadButton_Click);

                _toolbar.Visible = true;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"创建工具栏失败: {ex.Message}");
            }
        }

        private void DownloadButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                string saveRoot;
                DateTime startDate, endDate;

                if (!SelectSaveOptions(out saveRoot, out startDate, out endDate))
                {
                    MessageBox.Show("操作取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                MAPIFolder inbox = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                Items items = inbox.Items;
                // 限制只获取必要字段，减少内存占用
                items.IncludeRecurrences = false;

                int savedCount = 0;
                int skippedCount = 0;
                int processedCount = 0;

                using (var progressForm = new ProgressForm(AppVersion))
                {
                    progressForm.Show();
                    progressForm.SetProgress(0, items.Count);

                    // 使用 for 循环代替 foreach，更好地控制 COM 对象释放
                    for (int i = 1; i <= items.Count; i++)
                    {
                        object item = null;
                        MailItem mailItem = null;
                        Attachments attachments = null;

                        try
                        {
                            item = items[i];
                            mailItem = item as MailItem;

                            if (mailItem != null && mailItem.ReceivedTime >= startDate && mailItem.ReceivedTime <= endDate)
                            {
                                string monthFolder = Path.Combine(saveRoot, mailItem.ReceivedTime.ToString("yyyyMM"));
                                Directory.CreateDirectory(monthFolder);

                                attachments = mailItem.Attachments;
                                for (int j = 1; j <= attachments.Count; j++)
                                {
                                    Attachment attachment = null;
                                    try
                                    {
                                        attachment = attachments[j];

                                        // 跳过内联图片（小于100KB的图片文件通常是邮件正文中的图标、表情等）
                                        string ext = Path.GetExtension(attachment.FileName).ToLower();
                                        bool isImage = ext == ".png" || ext == ".jpg" || ext == ".jpeg" ||
                                                       ext == ".gif" || ext == ".bmp" || ext == ".ico" || ext == ".webp";

                                        if (isImage && attachment.Size < 102400) // 小于100KB的图片跳过
                                        {
                                            continue;
                                        }

                                        string safeFileName = SanitizeFileName(attachment.FileName);
                                        // 使用邮件接收时间戳+原文件名作为唯一标识
                                        string timestamp = mailItem.ReceivedTime.ToString("yyyyMMdd_HHmmss_fff");
                                        string uniqueFileName = $"{timestamp}_{safeFileName}";
                                        string targetPath = Path.Combine(monthFolder, uniqueFileName);

                                        if (File.Exists(targetPath))
                                        {
                                            skippedCount++;
                                        }
                                        else
                                        {
                                            attachment.SaveAsFile(targetPath);
                                            savedCount++;
                                        }
                                    }
                                    finally
                                    {
                                        // 释放附件 COM 对象
                                        if (attachment != null)
                                        {
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(attachment);
                                        }
                                    }
                                }
                                processedCount++;

                                // 每处理 50 封邮件强制垃圾回收一次
                                if (processedCount % 50 == 0)
                                {
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                }
                            }
                        }
                        finally
                        {
                            // 释放 COM 对象
                            if (attachments != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(attachments);
                            }
                            if (mailItem != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);
                            }
                            if (item != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                            }
                        }

                        progressForm.SetProgress(i, items.Count);

                        // 每 100 封邮件让 UI 刷新一下，避免假死
                        if (i % 100 == 0)
                        {
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }
                }

                    MessageBox.Show($"保存完成！已保存 {savedCount} 个附件，跳过 {skippedCount} 个已存在附件。", $"jimsoutlooktools {AppVersion}", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", $"jimsoutlooktools {AppVersion}", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool SelectSaveOptions(out string saveRoot, out DateTime startDate, out DateTime endDate)
        {
            saveRoot = null;
            startDate = DateTime.MinValue;
            endDate = DateTime.MaxValue;

            using (var form = new DateRangePickerForm(AppVersion))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    saveRoot = form.SavePath;
                    // 起始日期设为当天00:00:00，结束日期设为当天23:59:59，确保包含整天
                    startDate = form.StartDate.Date;
                    endDate = form.EndDate.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

                    if (startDate > endDate)
                    {
                        MessageBox.Show("起始日期不能晚于结束日期！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }

                    return true;
                }
            }

            return false;
        }

        private string SanitizeFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '-');
            }

            return fileName.Length > 180 ? fileName.Substring(0, 180) : fileName;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                if (_toolbar != null)
                {
                    _toolbar.Delete();
                }
            }
            catch { }
        }

        #region VSTO 生成的代码

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

        public ProgressForm(string appVersion = "v1.0.2")
        {
            this.Text = $"jimsoutlooktools {appVersion} - 保存进度";
            this.Width = 400;
            this.Height = 100;

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = 100
            };

            this.Controls.Add(progressBar);
        }

        public void SetProgress(int current, int total)
        {
            if (total > 0)
            {
                progressBar.Maximum = total;
                progressBar.Value = current;
            }
        }

        public void IncrementProgress()
        {
            if (progressBar.Value < progressBar.Maximum)
            {
                progressBar.Value++;
            }
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

        public DateTime StartDate { get; private set; }
        public DateTime EndDate { get; private set; }
        public string SavePath { get; private set; }

        public DateRangePickerForm(string appVersion = "v1.0.2")
        {
            this.Text = $"jimsoutlooktools {appVersion} - 保存邮件附件";
            this.Width = 450;
            this.Height = 360;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 主容器
            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 10,
                Padding = new Padding(15)
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

            // 分隔线1
            var separator1 = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                Height = 5,
                BorderStyle = BorderStyle.Fixed3D
            };

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

                SavePath = pathTextBox.Text;
                StartDate = startDatePicker.Value;
                EndDate = endDatePicker.Value;
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
            tableLayout.Controls.Add(separator1, 0, 3);
            tableLayout.Controls.Add(startLabel, 0, 4);
            tableLayout.Controls.Add(startDatePicker, 0, 5);
            tableLayout.Controls.Add(spacer, 0, 6);
            tableLayout.Controls.Add(endLabel, 0, 7);
            tableLayout.Controls.Add(endDatePicker, 0, 8);
            tableLayout.Controls.Add(buttonPanel, 0, 9);

            // 设置行高
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40));
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 10));
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 10));
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30));
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35));
            tableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50));

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
}
