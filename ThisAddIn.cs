using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.IO;

namespace jimsoutlooktools
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Ribbon API 替代 CommandBar
        }

        private void DownloadAttachmentButton_Click(object sender, EventArgs e)
        {
            try
            {
                // 选择保存附件的根目录
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "请选择附件保存的根文件夹";
                    if (folderDialog.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(folderDialog.SelectedPath))
                    {
                        MessageBox.Show("未选择保存路径，操作取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    string saveRoot = folderDialog.SelectedPath;

                    // 选择日期范围
                    DateTime startDate, endDate;
                    if (!SelectDateRange(out startDate, out endDate))
                    {
                        MessageBox.Show("未选择日期范围，操作取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 获取收件箱邮件
                    MAPIFolder inbox = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                    Items items = inbox.Items;

                    // 初始化统计数据
                    int savedCount = 0;
                    int skippedCount = 0;

                    // 显示进度条
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

                            // 更新进度条
                            progressForm.IncrementProgress();
                        }
                    }

                    // 显示统计结果
                    MessageBox.Show($"保存完成！已保存 {savedCount} 个附件，跳过 {skippedCount} 个附件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex) // 明确使用 System.Exception
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
            // Ribbon API 不需要显式清理
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
