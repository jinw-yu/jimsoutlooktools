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
        private const string ToolbarName = "邮件附件工具";

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
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "请选择附件保存的根文件夹";
                    if (folderDialog.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(folderDialog.SelectedPath))
                    {
                        MessageBox.Show("未选择保存路径，操作取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    string saveRoot = folderDialog.SelectedPath;

                    DateTime startDate, endDate;
                    if (!SelectDateRange(out startDate, out endDate))
                    {
                        MessageBox.Show("未选择日期范围，操作取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    MAPIFolder inbox = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                    Items items = inbox.Items;

                    int savedCount = 0;
                    int skippedCount = 0;

                    using (var progressForm = new ProgressForm())
                    {
                        progressForm.Show();
                        progressForm.SetProgress(0, items.Count);

                        foreach (object item in items)
                        {
                            if (item is MailItem mailItem)
                            {
                                if (mailItem.ReceivedTime >= startDate && mailItem.ReceivedTime <= endDate)
                                {
                                    string monthFolder = Path.Combine(saveRoot, mailItem.ReceivedTime.ToString("yyyyMM"));
                                    Directory.CreateDirectory(monthFolder);

                                    foreach (Attachment attachment in mailItem.Attachments)
                                    {
                                        string safeFileName = SanitizeFileName(attachment.FileName);
                                        string targetPath = Path.Combine(monthFolder, safeFileName);

                                        if (!File.Exists(targetPath))
                                        {
                                            attachment.SaveAsFile(targetPath);
                                            savedCount++;
                                        }
                                        else
                                        {
                                            skippedCount++;
                                        }
                                    }
                                }
                            }

                            progressForm.IncrementProgress();
                        }
                    }

                    MessageBox.Show($"保存完成！已保存 {savedCount} 个附件，跳过 {skippedCount} 个附件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool SelectDateRange(out DateTime startDate, out DateTime endDate)
        {
            startDate = DateTime.MinValue;
            endDate = DateTime.MaxValue;

            using (var form = new DateRangePickerForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    startDate = form.StartDate;
                    endDate = form.EndDate;

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

        public ProgressForm()
        {
            this.Text = "保存进度";
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
        private Button okButton;
        private Button cancelButton;

        public DateTime StartDate { get; private set; }
        public DateTime EndDate { get; private set; }

        public DateRangePickerForm()
        {
            this.Text = "选择日期范围";
            this.Width = 300;
            this.Height = 200;

            startDatePicker = new DateTimePicker { Dock = DockStyle.Top };
            endDatePicker = new DateTimePicker { Dock = DockStyle.Top };

            okButton = new Button { Text = "确定", Dock = DockStyle.Bottom };
            cancelButton = new Button { Text = "取消", Dock = DockStyle.Bottom };

            okButton.Click += (sender, e) =>
            {
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

            this.Controls.Add(endDatePicker);
            this.Controls.Add(startDatePicker);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);
        }
    }
}
