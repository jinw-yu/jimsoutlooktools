using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace jimsoutlooktools
{
    public partial class RibbonTools
    {
        private const string AppVersion = "v1.0.4";

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

                using (var progressForm = new ProgressForm(AppVersion))
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
            MessageBox.Show($"发生错误: {ex.Message}", $"jimsoutlooktools {AppVersion}", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                using (var resultForm = new SaveResultForm(AppVersion, message.ToString()))
                {
                    resultForm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show(message.ToString(), $"jimsoutlooktools {AppVersion} - 保存结果", 
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

            using (var form = new DateRangePickerForm(AppVersion))
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
                using (var selectForm = new DataFileSelectForm(Globals.ThisAddIn.Application))
                {
                    if (selectForm.ShowDialog() != DialogResult.OK)
                        return;

                    Outlook.MAPIFolder sourceRoot = selectForm.SourceRootFolder;
                    Outlook.MAPIFolder targetRoot = selectForm.TargetRootFolder;

                    if (sourceRoot == null || targetRoot == null)
                    {
                        MessageBox.Show("请选择有效的源数据文件和目标数据文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 分析文件夹差异（显示进度）
                    var folderDiffs = AnalyzeFolderDifferencesWithProgress(sourceRoot, targetRoot);

                    if (folderDiffs.Count == 0)
                    {
                        MessageBox.Show("两个数据文件的文件夹结构相同，没有需要同步的差异。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 显示差异并让用户选择要同步的文件夹
                    using (var diffForm = new FolderDiffForm(folderDiffs))
                    {
                        if (diffForm.ShowDialog() != DialogResult.OK)
                            return;

                        var selectedFolders = diffForm.SelectedFolders;
                        if (selectedFolders.Count == 0)
                        {
                            MessageBox.Show("未选择任何文件夹进行同步。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        // 开始同步
                        SyncFolders(selectedFolders, sourceRoot, targetRoot);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", $"jimsoutlooktools {AppVersion}", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<FolderDiffInfo> AnalyzeFolderDifferencesWithProgress(Outlook.MAPIFolder sourceRoot, Outlook.MAPIFolder targetRoot)
        {
            var result = new List<FolderDiffInfo>();

            using (var progressForm = new AnalysisProgressForm())
            {
                progressForm.Show();
                System.Windows.Forms.Application.DoEvents();

                // 第一步：获取所有文件夹
                progressForm.UpdateStatus("正在获取源数据文件文件夹列表...");
                var sourceFolders = GetAllFolders(sourceRoot);
                progressForm.UpdateStatus($"找到 {sourceFolders.Count} 个文件夹，正在获取目标数据文件文件夹列表...");
                
                var targetFolders = GetAllFolders(targetRoot);
                progressForm.UpdateStatus($"找到 {targetFolders.Count} 个文件夹，开始分析差异...");

                // 第二步：分析差异
                int processedCount = 0;
                int totalFolders = sourceFolders.Count;

                // 构建目标文件夹路径字典，使用相对路径作为键
                var targetFolderDict = new Dictionary<string, Outlook.MAPIFolder>(System.StringComparer.OrdinalIgnoreCase);
                foreach (var folder in targetFolders)
                {
                    string relativePath = GetRelativeFolderPath(folder, targetRoot);
                    if (!string.IsNullOrEmpty(relativePath))
                    {
                        targetFolderDict[relativePath] = folder;
                    }
                }

                foreach (var sourceFolder in sourceFolders)
                {
                    processedCount++;
                    
                    // 获取相对于源根的路径
                    string relativePath = GetRelativeFolderPath(sourceFolder, sourceRoot);
                    if (string.IsNullOrEmpty(relativePath))
                        continue;

                    // 每5个文件夹更新一次进度
                    if (processedCount % 5 == 0)
                    {
                        int percent = (int)((double)processedCount / totalFolders * 100);
                        progressForm.UpdateProgress(processedCount, totalFolders, percent);
                        progressForm.UpdateStatus($"正在分析: {relativePath}");
                        System.Windows.Forms.Application.DoEvents();
                    }

                    // 在目标文件夹中查找匹配的相对路径
                    Outlook.MAPIFolder targetFolder = null;
                    targetFolderDict.TryGetValue(relativePath, out targetFolder);

                    // 使用更准确的方法获取邮件数量
                    int sourceCount = GetMailItemCount(sourceFolder);
                    int targetCount = targetFolder != null ? GetMailItemCount(targetFolder) : 0;
                    
                    // 调试信息
                    System.Diagnostics.Debug.WriteLine($"[分析] 相对路径: {relativePath}, 源: {sourceCount}, 目标: {targetCount}");
                    
                    int diffCount = sourceCount - targetCount;

                    if (diffCount > 0)
                    {
                        result.Add(new FolderDiffInfo
                        {
                            FolderPath = relativePath,
                            SourceFolder = sourceFolder,
                            TargetFolder = targetFolder,
                            SourceCount = sourceCount,
                            TargetCount = targetCount,
                            DiffCount = diffCount
                        });
                    }
                }

                progressForm.Complete($"分析完成！发现 {result.Count} 个有差异的文件夹");
            }

            return result;
        }

        private List<FolderDiffInfo> AnalyzeFolderDifferences(Outlook.MAPIFolder sourceRoot, Outlook.MAPIFolder targetRoot)
        {
            var result = new List<FolderDiffInfo>();
            var sourceFolders = GetAllFolders(sourceRoot);
            var targetFolders = GetAllFolders(targetRoot);

            // 构建目标文件夹路径字典，使用相对路径作为键
            var targetFolderDict = new Dictionary<string, Outlook.MAPIFolder>(System.StringComparer.OrdinalIgnoreCase);
            foreach (var folder in targetFolders)
            {
                string relativePath = GetRelativeFolderPath(folder, targetRoot);
                if (!string.IsNullOrEmpty(relativePath))
                {
                    targetFolderDict[relativePath] = folder;
                }
            }

            foreach (var sourceFolder in sourceFolders)
            {
                // 获取相对于源根的路径
                string relativePath = GetRelativeFolderPath(sourceFolder, sourceRoot);
                if (string.IsNullOrEmpty(relativePath))
                    continue;

                // 在目标文件夹中查找匹配的相对路径
                Outlook.MAPIFolder targetFolder = null;
                targetFolderDict.TryGetValue(relativePath, out targetFolder);

                // 使用更准确的方法获取邮件数量
                int sourceCount = GetMailItemCount(sourceFolder);
                int targetCount = targetFolder != null ? GetMailItemCount(targetFolder) : 0;
                int diffCount = sourceCount - targetCount;

                if (diffCount > 0)
                {
                    result.Add(new FolderDiffInfo
                    {
                        FolderPath = relativePath,
                        SourceFolder = sourceFolder,
                        TargetFolder = targetFolder,
                        SourceCount = sourceCount,
                        TargetCount = targetCount,
                        DiffCount = diffCount
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// 获取文件夹中邮件项目的准确数量
        /// </summary>
        private int GetMailItemCount(Outlook.MAPIFolder folder)
        {
            if (folder == null) return 0;

            // 对于 PST 文件，直接使用 Items.Count 通常是最准确的
            // 因为 PST 文件夹通常只包含邮件项目
            try
            {
                int rawCount = folder.Items.Count;
                System.Diagnostics.Debug.WriteLine($"文件夹 '{folder.Name}' Items.Count: {rawCount}");
                return rawCount;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取计数失败: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 检查文件夹是否应该被排除（系统文件夹）
        /// </summary>
        private bool ShouldExcludeFolder(Outlook.MAPIFolder folder)
        {
            if (folder == null) return true;

            string folderName = folder.Name.ToLower();
            
            // 排除的系统文件夹列表
            string[] excludedFolders = new[]
            {
                // 中文名称
                "日历", "任务", "便签", "垃圾邮件", "发件箱", 
                "同步文件", "已删除邮件", "已发送邮件", "草稿",
                "联系人", "日记", "便笺", " rss源", "建议的联系人",
                "收件箱", "outbox", "sent items", "deleted items",
                "junk email", "drafts", "contacts", "calendar",
                "tasks", "notes", "journal", "sync issues",
                "conflicts", "local failures", "server failures",
                "search folders", "conversation action settings",
                "quick step settings", "rss 源"
            };

            foreach (var excluded in excludedFolders)
            {
                if (folderName.Contains(excluded))
                    return true;
            }

            return false;
        }

        private List<Outlook.MAPIFolder> GetAllFolders(Outlook.MAPIFolder root)
        {
            var result = new List<Outlook.MAPIFolder>();
            
            // 检查根文件夹是否被排除
            if (ShouldExcludeFolder(root))
                return result;
                
            result.Add(root);
            
            foreach (Outlook.MAPIFolder folder in root.Folders)
            {
                // 跳过被排除的文件夹
                if (ShouldExcludeFolder(folder))
                    continue;
                result.AddRange(GetAllFolders(folder));
            }
            
            return result;
        }

        private string GetFolderPath(Outlook.MAPIFolder folder)
        {
            var parts = new List<string>();
            Outlook.MAPIFolder current = folder;

            while (current != null)
            {
                parts.Insert(0, current.Name);
                try
                {
                    current = current.Parent as Outlook.MAPIFolder;
                }
                catch
                {
                    break;
                }
            }

            return string.Join("\\", parts);
        }

        /// <summary>
        /// 获取相对于根文件夹的路径
        /// </summary>
        private string GetRelativeFolderPath(Outlook.MAPIFolder folder, Outlook.MAPIFolder root)
        {
            if (folder == null) return null;
            if (folder == root) return "";

            var parts = new List<string>();
            Outlook.MAPIFolder current = folder;

            // 向上遍历直到到达根文件夹
            while (current != null && current != root)
            {
                parts.Insert(0, current.Name);
                try
                {
                    current = current.Parent as Outlook.MAPIFolder;
                }
                catch
                {
                    break;
                }
            }

            // 如果没有到达根文件夹，说明不在同一棵树中
            if (current != root)
                return null;

            return string.Join("\\", parts);
        }

        private Outlook.MAPIFolder FindFolderByPath(Outlook.MAPIFolder root, string path)
        {
            string rootPath = GetFolderPath(root);
            if (path == rootPath)
                return root;

            if (!path.StartsWith(rootPath + "\\"))
                return null;

            string relativePath = path.Substring(rootPath.Length + 1);
            string[] parts = relativePath.Split('\\');

            Outlook.MAPIFolder current = root;
            foreach (string part in parts)
            {
                bool found = false;
                foreach (Outlook.MAPIFolder folder in current.Folders)
                {
                    if (folder.Name == part)
                    {
                        current = folder;
                        found = true;
                        break;
                    }
                }
                if (!found)
                    return null;
            }

            return current;
        }

        private void SyncFolders(List<FolderDiffInfo> selectedFolders, Outlook.MAPIFolder sourceRoot, Outlook.MAPIFolder targetRoot)
        {
            int totalEmails = selectedFolders.Sum(f => f.DiffCount);
            int processedEmails = 0;
            int successCount = 0;
            int failedCount = 0;
            var failedEmails = new List<string>();

            using (var syncProgress = new SyncProgressForm(AppVersion, totalEmails))
            {
                syncProgress.Show();

                foreach (var folderDiff in selectedFolders)
                {
                    if (syncProgress.IsCancelled)
                        break;

                    Outlook.MAPIFolder sourceFolder = folderDiff.SourceFolder;
                    Outlook.MAPIFolder targetFolder = folderDiff.TargetFolder;

                    // 如果目标文件夹不存在，创建它
                    if (targetFolder == null)
                    {
                        targetFolder = CreateFolderStructure(sourceRoot, targetRoot, folderDiff.FolderPath);
                        if (targetFolder == null)
                        {
                            failedEmails.Add($"文件夹 {folderDiff.FolderPath} - 创建目标文件夹失败");
                            continue;
                        }
                    }

                    // 获取目标文件夹中已有的邮件EntryID，避免重复
                    var existingEntryIds = new HashSet<string>();
                    foreach (object item in targetFolder.Items)
                    {
                        if (item is Outlook.MailItem mail)
                        {
                            try
                            {
                                existingEntryIds.Add(mail.EntryID);
                            }
                            finally
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                            }
                        }
                        else
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                        }
                    }

                    // 复制邮件
                    Outlook.Items sourceItems = sourceFolder.Items;
                    for (int i = 1; i <= sourceItems.Count; i++)
                    {
                        if (syncProgress.IsCancelled)
                            break;

                        object item = null;
                        Outlook.MailItem sourceMail = null;

                        try
                        {
                            item = sourceItems[i];
                            sourceMail = item as Outlook.MailItem;

                            if (sourceMail == null)
                                continue;

                            // 检查是否已存在
                            if (existingEntryIds.Contains(sourceMail.EntryID))
                                continue;

                            // 复制邮件
                            try
                            {
                                // 使用 Copy() 方法复制邮件，然后移动到目标文件夹
                                var copiedMail = sourceMail.Copy();
                                copiedMail.Move(targetFolder);
                                successCount++;
                            }
                            catch (System.Exception ex)
                            {
                                failedCount++;
                                failedEmails.Add($"邮件: {sourceMail.Subject} | 文件夹: {folderDiff.FolderPath} | 错误: {ex.Message}");
                                System.Diagnostics.Debug.WriteLine($"复制邮件失败: {sourceMail.Subject} - {ex.Message}");
                            }

                            processedEmails++;

                            // 每处理10封邮件更新一次进度
                            if (processedEmails % 10 == 0)
                            {
                                syncProgress.UpdateProgress(processedEmails, totalEmails, folderDiff.FolderPath);
                                System.Windows.Forms.Application.DoEvents();
                            }

                            // 每50封邮件垃圾回收
                            if (processedEmails % 50 == 0)
                            {
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                            }
                        }
                        finally
                        {
                            if (sourceMail != null)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMail);
                            if (item != null)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                        }
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceItems);
                }

                syncProgress.Complete();
            }

            // 显示结果
            ShowSyncResult(successCount, failedCount, failedEmails);
        }

        private Outlook.MAPIFolder CreateFolderStructure(Outlook.MAPIFolder sourceRoot, Outlook.MAPIFolder targetRoot, string folderPath)
        {
            string rootPath = GetFolderPath(sourceRoot);
            if (!folderPath.StartsWith(rootPath))
                return null;

            string relativePath = folderPath.Substring(rootPath.Length).TrimStart('\\');
            string[] parts = relativePath.Split('\\');

            Outlook.MAPIFolder current = targetRoot;
            string currentPath = GetFolderPath(targetRoot);

            foreach (string part in parts)
            {
                currentPath = currentPath + "\\" + part;
                Outlook.MAPIFolder nextFolder = null;

                foreach (Outlook.MAPIFolder folder in current.Folders)
                {
                    if (folder.Name == part)
                    {
                        nextFolder = folder;
                        break;
                    }
                }

                if (nextFolder == null)
                {
                    // 创建新文件夹
                    try
                    {
                        nextFolder = current.Folders.Add(part, Outlook.OlDefaultFolders.olFolderInbox);
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"创建文件夹失败: {currentPath} - {ex.Message}");
                        return null;
                    }
                }

                current = nextFolder;
            }

            return current;
        }

        private void ShowSyncResult(int successCount, int failedCount, List<string> failedEmails)
        {
            StringBuilder message = new StringBuilder();
            message.AppendLine($"同步完成！");
            message.AppendLine();
            message.AppendLine($"✓ 成功复制: {successCount} 封邮件");
            message.AppendLine($"✗ 复制失败: {failedCount} 封邮件");

            if (failedCount > 0)
            {
                message.AppendLine();
                message.AppendLine("失败详情:");
                message.AppendLine("--------------------");
                foreach (var failed in failedEmails)
                {
                    message.AppendLine($"• {failed}");
                }
            }

            if (failedCount > 5)
            {
                using (var resultForm = new SaveResultForm(AppVersion, message.ToString()))
                {
                    resultForm.Text = $"jimsoutlooktools {AppVersion} - 同步结果详情";
                    resultForm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show(message.ToString(), $"jimsoutlooktools {AppVersion} - 同步结果", 
                    MessageBoxButtons.OK, failedCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
        }

        #endregion
    }

    #region 数据类

    public class FolderDiffInfo
    {
        public string FolderPath { get; set; }
        public Outlook.MAPIFolder SourceFolder { get; set; }
        public Outlook.MAPIFolder TargetFolder { get; set; }
        public int SourceCount { get; set; }
        public int TargetCount { get; set; }
        public int DiffCount { get; set; }
    }

    #endregion
}
