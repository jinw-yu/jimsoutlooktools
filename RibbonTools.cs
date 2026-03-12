using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace jtools_outlook
{
    public partial class RibbonTools
    {
        private const string AppVersion = "v1.1.1";

        private void RibbonTools_Load(object sender, RibbonUIEventArgs e)
        {
        }

        #region 保存附件功能

        private void btnSaveAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string saveRoot;
                DateTime startDate, endDate;
                bool saveInbox, saveSentItems;

                if (!SelectSaveOptions(out saveRoot, out startDate, out endDate, out saveInbox, out saveSentItems))
                {
                    MessageBox.Show("操作取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int savedCount = 0;
                int skippedCount = 0;
                int processedCount = 0;
                var failedAttachments = new List<string>();

                // 计算符合日期范围的总邮件数
                int totalItems = 0;
                string dateFilter = $"[ReceivedTime] >= '{startDate.ToString("g")}' AND [ReceivedTime] <= '{endDate.ToString("g")}'";

                if (saveInbox)
                {
                    Outlook.MAPIFolder inbox = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    Outlook.Items filteredItems = inbox.Items.Restrict(dateFilter);
                    totalItems += filteredItems.Count;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);
                }
                if (saveSentItems)
                {
                    Outlook.MAPIFolder sentItems = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                    Outlook.Items filteredItems = sentItems.Items.Restrict(dateFilter);
                    totalItems += filteredItems.Count;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sentItems);
                }

                int currentItemIndex = 0;

                using (var progressForm = new ProgressForm())
                {
                    progressForm.Show();
                    progressForm.SetProgress(0, totalItems);

                    // 处理收件箱
                    if (saveInbox)
                    {
                        progressForm.UpdateStatus("正在处理收件箱...");
                        ProcessFolderAttachments(
                            Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox),
                            saveRoot, startDate, endDate,
                            ref currentItemIndex, totalItems,
                            ref savedCount, ref skippedCount, ref processedCount,
                            failedAttachments, progressForm);
                    }

                    // 处理已发送邮件
                    if (saveSentItems && !progressForm.IsCancelled)
                    {
                        progressForm.UpdateStatus("正在处理已发送邮件...");
                        ProcessFolderAttachments(
                            Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail),
                            saveRoot, startDate, endDate,
                            ref currentItemIndex, totalItems,
                            ref savedCount, ref skippedCount, ref processedCount,
                            failedAttachments, progressForm);
                    }
                }

            // 显示详细的保存结果
            ShowSaveResult(savedCount, skippedCount, failedAttachments, saveInbox, saveSentItems);
        }
        catch (System.Exception ex)
        {
            MessageBox.Show($"发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ProcessFolderAttachments(
        Outlook.MAPIFolder folder,
        string saveRoot,
        DateTime startDate,
        DateTime endDate,
        ref int currentItemIndex,
        int totalItems,
        ref int savedCount,
        ref int skippedCount,
        ref int processedCount,
        List<string> failedAttachments,
        ProgressForm progressForm)
    {
        string folderName = folder.Name;

        // 使用日期过滤，只处理符合范围的邮件
        string dateFilter = $"[ReceivedTime] >= '{startDate.ToString("g")}' AND [ReceivedTime] <= '{endDate.ToString("g")}'";
        Outlook.Items allItems = folder.Items;
        allItems.IncludeRecurrences = false;
        Outlook.Items filteredItems = allItems.Restrict(dateFilter);

        for (int i = 1; i <= filteredItems.Count; i++)
        {
            currentItemIndex++;

            if (progressForm.IsCancelled)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(allItems);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
                return;
            }

            object item = null;
            Outlook.MailItem mailItem = null;
            Outlook.Attachments attachments = null;

            try
            {
                item = filteredItems[i];
                mailItem = item as Outlook.MailItem;

                if (mailItem != null)
                {
                    string monthFolder = Path.Combine(saveRoot, mailItem.ReceivedTime.ToString("yyyyMM"));
                    Directory.CreateDirectory(monthFolder);

                    attachments = mailItem.Attachments;
                    for (int j = 1; j <= attachments.Count; j++)
                    {
                        Outlook.Attachment attachment = null;
                        try
                        {
                            attachment = attachments[j];

                            string ext = Path.GetExtension(attachment.FileName).ToLower();
                            bool isImage = ext == ".png" || ext == ".jpg" || ext == ".jpeg" ||
                                           ext == ".gif" || ext == ".bmp" || ext == ".ico" || ext == ".webp";

                            if (isImage && attachment.Size < 102400)
                            {
                                continue;
                            }

                            string safeFileName = SanitizeFileName(attachment.FileName);
                            string timestamp = mailItem.ReceivedTime.ToString("yyyyMMdd_HHmmss_fff");
                            string uniqueFileName = $"{timestamp}_{safeFileName}";
                            string targetPath = Path.Combine(monthFolder, uniqueFileName);

                            if (File.Exists(targetPath))
                            {
                                skippedCount++;
                            }
                            else
                            {
                                try
                                {
                                    attachment.SaveAsFile(targetPath);
                                    savedCount++;
                                }
                                catch (System.Exception ex)
                                {
                                    string failedInfo = $"文件: {attachment.FileName} | 邮件: {mailItem.Subject} | 文件夹: {folderName} | 时间: {mailItem.ReceivedTime:yyyy-MM-dd HH:mm:ss} | 错误: {ex.Message}";
                                    failedAttachments.Add(failedInfo);
                                    System.Diagnostics.Debug.WriteLine($"保存附件失败: {failedInfo}");
                                }
                            }
                        }
                        finally
                        {
                            if (attachment != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(attachment);
                            }
                        }
                    }
                    processedCount++;

                    if (processedCount % 50 == 0)
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
            }
            finally
            {
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

            if (currentItemIndex % 10 == 0 || currentItemIndex == totalItems)
            {
                progressForm.SetProgress(currentItemIndex, totalItems);
                progressForm.UpdateStatus($"正在处理 {folderName}: {currentItemIndex} / {totalItems} | 已保存: {savedCount} 个附件");
                System.Windows.Forms.Application.DoEvents();
            }
        }

        System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(allItems);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
    }

    private void ShowSaveResult(int savedCount, int skippedCount, List<string> failedAttachments, bool saveInbox, bool saveSentItems)
        {
            int failedCount = failedAttachments.Count;
            StringBuilder message = new StringBuilder();
            message.AppendLine($"保存完成！");
            message.AppendLine();

            // 显示处理的文件夹
            var folders = new List<string>();
            if (saveInbox) folders.Add("收件箱");
            if (saveSentItems) folders.Add("已发送邮件");
            message.AppendLine($"📁 处理文件夹: {string.Join("、", folders)}");
            message.AppendLine();

            message.AppendLine($"✓ 已保存: {savedCount} 个附件");
            message.AppendLine($"○ 跳过(已存在): {skippedCount} 个附件");
            message.AppendLine($"✗ 保存失败: {failedCount} 个附件");

            if (failedCount > 0)
            {
                message.AppendLine();
                message.AppendLine("失败详情:");
                message.AppendLine("--------------------");
                foreach (var failed in failedAttachments)
                {
                    message.AppendLine($"• {failed}");
                }
            }

            // 如果失败数量较多，使用滚动文本框显示
            if (failedCount > 5)
            {
                using (var resultForm = new SaveResultForm(message.ToString()))
                {
                    resultForm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show(message.ToString(), "保存结果",
                    MessageBoxButtons.OK, failedCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
        }

        private bool SelectSaveOptions(out string saveRoot, out DateTime startDate, out DateTime endDate, out bool saveInbox, out bool saveSentItems)
        {
            saveRoot = null;
            startDate = DateTime.MinValue;
            endDate = DateTime.MaxValue;
            saveInbox = true;
            saveSentItems = false;

            using (var form = new DateRangePickerForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    saveRoot = form.SavePath;
                    // 起始日期设为当天00:00:00，结束日期设为当天23:59:59，确保包含整天
                    startDate = form.StartDate.Date;
                    endDate = form.EndDate.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
                    saveInbox = form.SaveInbox;
                    saveSentItems = form.SaveSentItems;

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

        #endregion

        #region 下载联机功能

        private void btnDownloadOnline_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                using (var form = new DownloadOnlineForm(Globals.ThisAddIn.Application))
                {
                    form.ShowDialog();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region 阻止功能

        private void btnBlockDomain_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前选中的邮件
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null || explorer.Selection == null || explorer.Selection.Count == 0)
                {
                    MessageBox.Show("请先选择一封邮件。", "JTools-outlook - 提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var selectedItem = explorer.Selection[1];
                if (!(selectedItem is Outlook.MailItem mailItem))
                {
                    MessageBox.Show("选中的项目不是邮件。", "JTools-outlook - 提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 获取发件人邮箱地址
                string senderEmail = mailItem.SenderEmailAddress;
                if (string.IsNullOrEmpty(senderEmail))
                {
                    MessageBox.Show("无法获取发件人邮箱地址。", "JTools-outlook - 提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 提取域名
                string domain = ExtractDomain(senderEmail);
                if (string.IsNullOrEmpty(domain))
                {
                    MessageBox.Show($"无法从邮箱地址 '{senderEmail}' 中提取域名。", "JTools-outlook - 提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 显示选择对话框
                using (var selectForm = new BlockSelectForm(senderEmail, domain))
                {
                    var result = selectForm.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        // 根据用户选择执行相应的阻止操作
                        if (selectForm.BlockType == BlockType.Sender)
                        {
                            BlockSender(senderEmail);
                        }
                        else if (selectForm.BlockType == BlockType.Domain)
                        {
                            AddDomainToBlockedSenders(domain);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"阻止操作时发生错误：{ex.Message}",
                    "JTools-outlook - 错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private string ExtractDomain(string email)
        {
            try
            {
                if (string.IsNullOrEmpty(email))
                    return null;

                // 处理 SMTP 格式的邮箱地址
                if (email.StartsWith("SMTP:") || email.StartsWith("smtp:"))
                {
                    email = email.Substring(5);
                }

                // 提取 @ 后面的域名
                int atIndex = email.IndexOf('@');
                if (atIndex >= 0 && atIndex < email.Length - 1)
                {
                    return email.Substring(atIndex + 1).ToLower();
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private void AddDomainToBlockedSenders(string domain)
        {
            try
            {
                // 显示阻止域对话框
                using (var dialog = new BlockDomainDialog(domain))
                {
                    dialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"无法添加阻止域: {ex.Message}",
                    "JTools-outlook - 错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void BlockSender(string senderEmail)
        {
            try
            {
                // 显示阻止发件人对话框
                using (var dialog = new BlockSenderDialog(senderEmail))
                {
                    dialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"无法添加阻止发件人: {ex.Message}",
                    "JTools-outlook - 错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        #endregion

        #region 带附件全部答复功能

        private void btnReplyAllWithAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前选中的邮件
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null || explorer.Selection == null || explorer.Selection.Count == 0)
                {
                    MessageBox.Show("请先选择一封邮件。", "JTools-outlook - 提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var selectedItem = explorer.Selection[1];
                if (!(selectedItem is Outlook.MailItem originalMail))
                {
                    MessageBox.Show("选中的项目不是邮件。", "JTools-outlook - 提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 检查是否有附件
                if (originalMail.Attachments.Count == 0)
                {
                    MessageBox.Show("当前邮件没有附件。", "JTools-outlook - 提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 创建全部答复邮件
                Outlook.MailItem replyMail = originalMail.ReplyAll();

                // 创建临时文件夹
                string tempFolder = Path.Combine(Path.GetTempPath(), $"JTools_Attachments_{Guid.NewGuid():N}");
                Directory.CreateDirectory(tempFolder);

                int copiedCount = 0;
                List<string> tempFiles = new List<string>();

                try
                {
                    // 将原邮件的附件保存到临时文件夹，然后添加到答复邮件中
                    for (int i = 1; i <= originalMail.Attachments.Count; i++)
                    {
                        var attachment = originalMail.Attachments[i];
                        try
                        {
                            // 获取附件的文件名
                            string fileName = attachment.FileName;
                            if (string.IsNullOrEmpty(fileName))
                            {
                                fileName = $"Attachment{i}";
                            }

                            // 检查是否为图片文件
                            string extension = Path.GetExtension(fileName).ToLower();
                            string[] imageExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".ico", ".webp" };
                            bool isImage = imageExtensions.Contains(extension);

                            // 如果是图片且小于 100KB，跳过
                            if (isImage && attachment.Size < 100 * 1024)
                            {
                                System.Diagnostics.Debug.WriteLine($"跳过小图片附件 {fileName} ({attachment.Size} bytes)");
                                continue;
                            }

                            // 保存附件到临时文件夹
                            string tempFilePath = Path.Combine(tempFolder, fileName);
                            attachment.SaveAsFile(tempFilePath);
                            tempFiles.Add(tempFilePath);

                            // 将附件添加到答复邮件
                            replyMail.Attachments.Add(
                                tempFilePath,
                                Outlook.OlAttachmentType.olByValue,
                                Type.Missing,
                                fileName
                            );
                            copiedCount++;
                        }
                        catch (System.Exception ex)
                        {
                            // 某些附件可能无法复制（如内嵌图片），跳过
                            System.Diagnostics.Debug.WriteLine($"无法复制附件 {attachment.FileName}: {ex.Message}");
                        }
                    }

                    // 显示答复邮件编辑窗口
                    replyMail.Display(false);
                }
                finally
                {
                    // 延迟删除临时文件（等待邮件窗口完全打开）
                    System.Threading.Tasks.Task.Run(async () =>
                    {
                        await System.Threading.Tasks.Task.Delay(5000); // 等待5秒
                        try
                        {
                            foreach (var tempFile in tempFiles)
                            {
                                if (File.Exists(tempFile))
                                {
                                    File.Delete(tempFile);
                                }
                            }
                            if (Directory.Exists(tempFolder))
                            {
                                Directory.Delete(tempFolder, true);
                            }
                        }
                        catch { }
                    });
                }

                // 释放资源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(replyMail);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    $"带附件全部答复时发生错误：{ex.Message}",
                    "JTools-outlook - 错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        #endregion

        #region 关于功能

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            using (var aboutForm = new AboutForm())
            {
                aboutForm.ShowDialog();
            }
        }

        #endregion
    }

    #region 阻止功能对话框

    public enum BlockType
    {
        None,
        Sender,
        Domain
    }

    public class BlockSelectForm : Form
    {
        public BlockType BlockType { get; private set; }
        private string senderEmail;
        private string domain;

        public BlockSelectForm(string senderEmail, string domain)
        {
            this.senderEmail = senderEmail;
            this.domain = domain;

            this.Text = "JTools-outlook - 选择阻止方式";
            this.Width = 500;
            this.Height = 250;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };

            var lblTitle = new Label
            {
                Text = "选择要阻止的内容",
                Dock = DockStyle.Top,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold),
                Height = 30,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            var lblInfo = new Label
            {
                Text = $"发件人: {senderEmail}\n发件人域: {domain}",
                Dock = DockStyle.Top,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                Height = 50,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 80
            };

            var btnBlockSender = new Button
            {
                Text = "阻止发件人",
                Width = 160,
                Height = 40,
                Left = 20,
                Top = 20,
                DialogResult = DialogResult.OK
            };
            btnBlockSender.Click += (s, e) => { BlockType = BlockType.Sender; };

            var btnBlockDomain = new Button
            {
                Text = "阻止发件人域",
                Width = 160,
                Height = 40,
                Left = 200,
                Top = 20,
                DialogResult = DialogResult.OK
            };
            btnBlockDomain.Click += (s, e) => { BlockType = BlockType.Domain; };

            var btnCancel = new Button
            {
                Text = "取消",
                Width = 100,
                Height = 40,
                Left = 380,
                Top = 20,
                DialogResult = DialogResult.Cancel
            };

            buttonPanel.Controls.Add(btnBlockSender);
            buttonPanel.Controls.Add(btnBlockDomain);
            buttonPanel.Controls.Add(btnCancel);

            mainPanel.Controls.Add(lblTitle);
            mainPanel.Controls.Add(lblInfo);

            this.Controls.Add(mainPanel);
            this.Controls.Add(buttonPanel);
        }
    }

    public class BlockSenderDialog : Form
    {
        private string senderEmail;
        private string registryPath = @"Software\Microsoft\Office\16.0\Outlook\Options\Mail";
        private string valueName = "BlockedSenders";
        private string fullRegistryPath;
        private string senderEntry;
        private TextBox txtLog;
        private Button btnConfirm;
        private Button btnCancel;
        private Button btnCopy;
        private Button btnClose;

        public BlockSenderDialog(string senderEmail)
        {
            this.senderEmail = senderEmail;
            this.fullRegistryPath = $"HKEY_CURRENT_USER\\{registryPath}";
            this.senderEntry = senderEmail;

            this.Text = $"JTools-outlook - 阻止发件人 {senderEmail}";
            this.Width = 600;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new System.Drawing.Size(600, 400);

            InitializeComponent();
            ShowConfirmLog();
        }

        private void InitializeComponent()
        {
            var lblTitle = new Label
            {
                Text = $"阻止发件人: {senderEmail}",
                Dock = DockStyle.Top,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 14, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.SteelBlue,
                Height = 40,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            txtLog = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Consolas", 9),
                BackColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            var panelButtons = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = System.Drawing.Color.LightGray
            };

            btnConfirm = new Button
            {
                Text = "确认执行",
                Width = 120,
                Height = 35,
                Left = 90,
                Top = 12
            };
            btnConfirm.Click += BtnConfirm_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Width = 120,
                Height = 35,
                Left = 230,
                Top = 12
            };
            btnCancel.Click += BtnCancel_Click;

            btnCopy = new Button
            {
                Text = "复制日志",
                Width = 120,
                Height = 35,
                Left = 90,
                Top = 12,
                Visible = false
            };
            btnCopy.Click += BtnCopy_Click;

            btnClose = new Button
            {
                Text = "关闭",
                Width = 120,
                Height = 35,
                Left = 370,
                Top = 12,
                Visible = false
            };
            btnClose.Click += BtnClose_Click;

            panelButtons.Controls.Add(btnConfirm);
            panelButtons.Controls.Add(btnCancel);
            panelButtons.Controls.Add(btnCopy);
            panelButtons.Controls.Add(btnClose);

            this.Controls.Add(txtLog);
            this.Controls.Add(lblTitle);
            this.Controls.Add(panelButtons);

            this.CancelButton = btnCancel;
        }

        private void ShowConfirmLog()
        {
            var log = new System.Text.StringBuilder();
            log.AppendLine("【操作内容】");
            log.AppendLine($"将发件人 '{senderEmail}' 添加到 Outlook 阻止发件人列表");
            log.AppendLine();
            log.AppendLine("【注册表修改】");
            log.AppendLine($"位置: {fullRegistryPath}");
            log.AppendLine($"值名: {valueName}");
            log.AppendLine($"类型: REG_MULTI_SZ (多字符串值)");
            log.AppendLine($"添加内容: {senderEntry}");
            log.AppendLine();
            log.AppendLine("【效果】");
            log.AppendLine("• 来自该发件人的所有邮件将被自动移动到垃圾邮件文件夹");
            log.AppendLine("• 当前邮件也会被移动到垃圾邮件文件夹");
            log.AppendLine("• 可能需要重启 Outlook 使设置生效");
            log.AppendLine();
            log.AppendLine("请确认是否继续执行？");

            txtLog.Text = log.ToString();
        }

        private void BtnConfirm_Click(object sender, EventArgs e)
        {
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

                // 检查是否已包含该发件人
                bool alreadyExists = false;
                foreach (string value in existingValues)
                {
                    if (value.Equals(senderEntry, StringComparison.OrdinalIgnoreCase))
                    {
                        alreadyExists = true;
                        break;
                    }
                }

                if (!alreadyExists)
                {
                    // 添加新发件人到列表
                    var newValues = new string[existingValues.Length + 1];
                    existingValues.CopyTo(newValues, 0);
                    newValues[existingValues.Length] = senderEntry;

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
            log.AppendLine($"已将发件人 '{senderEmail}' 添加到阻止发件人列表");
            log.AppendLine();
            log.AppendLine("【注册表修改】");
            log.AppendLine($"位置: {fullRegistryPath}");
            log.AppendLine($"值名: {valueName}");
            log.AppendLine($"类型: REG_MULTI_SZ (多字符串值)");
            log.AppendLine($"添加内容: {senderEntry}");
            log.AppendLine();
            log.AppendLine("【效果】");
            log.AppendLine("来自该发件人的所有邮件将被自动移动到垃圾邮件文件夹");
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
            log.AppendLine($"发件人 '{senderEmail}' 已在阻止发件人列表中");
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
            log.AppendLine("【建议】");
            log.AppendLine("1. 检查是否有足够的权限修改注册表");
            log.AppendLine("2. 尝试以管理员身份运行 Outlook");
            log.AppendLine("3. 手动在 Outlook 中添加阻止发件人");

            txtLog.Text = log.ToString();
            SwitchToResultMode();
        }

        private void SwitchToResultMode()
        {
            this.Text = $"JTools-outlook - 阻止结果 {senderEmail}";
            btnConfirm.Visible = false;
            btnCancel.Visible = false;
            btnCopy.Visible = true;
            btnClose.Visible = true;
            btnClose.Left = 230;
            this.CancelButton = btnClose;
            this.AcceptButton = btnClose;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("=== 点击了取消按钮 ===");
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
