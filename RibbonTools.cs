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
        private const string AppVersion = "v1.0.5";

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
                // 使用新的向导窗口
                using (var wizardForm = new DownloadOnlineWizardForm(Globals.ThisAddIn.Application))
                {
                    wizardForm.ShowDialog();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 分析两个数据文件的文件夹差异 - 简单直接对比，无过滤
        /// </summary>
        private List<FolderDiffInfo> AnalyzeFolderDifferencesWithProgress(Outlook.MAPIFolder sourceRoot, Outlook.MAPIFolder targetRoot)
        {
            var result = new List<FolderDiffInfo>();

            // 不使用using，让窗体保持打开直到用户手动关闭
            var progressForm = new AnalysisProgressForm();
            
            progressForm.Show();
            System.Windows.Forms.Application.DoEvents();

            progressForm.AddLog($"开始分析文件夹差异");
            progressForm.AddLog($"源数据文件: {sourceRoot.Name}");
            progressForm.AddLog($"目标数据文件: {targetRoot.Name}");
            progressForm.AddLog("");

            // 第一步：递归获取所有文件夹
            progressForm.UpdateStatus("正在获取源数据文件文件夹列表...");
            var sourceFolders = GetAllFoldersSimple(sourceRoot);
            progressForm.UpdateStatus($"源数据文件: {sourceFolders.Count} 个文件夹");
            progressForm.AddLog($"[源] 找到 {sourceFolders.Count} 个文件夹");

            progressForm.UpdateStatus("正在获取目标数据文件文件夹列表...");
            var targetFolders = GetAllFoldersSimple(targetRoot);
            progressForm.UpdateStatus($"目标数据文件: {targetFolders.Count} 个文件夹");
            progressForm.AddLog($"[目标] 找到 {targetFolders.Count} 个文件夹");

            // 记录源文件夹详情到日志
            progressForm.AddLog("");
            progressForm.AddLog("=== 源文件夹列表 ===");
            foreach (var folder in sourceFolders)
            {
                string simplePath = GetSimpleFolderPath(folder, sourceRoot);
                progressForm.AddLog($"  {folder.Name} -> [{simplePath ?? "null"}] ({folder.Items.Count}封)");
            }

            // 第二步：构建目标文件夹字典
            progressForm.AddLog("");
            progressForm.AddLog("=== 目标文件夹列表 ===");
            var targetFolderDict = new Dictionary<string, Outlook.MAPIFolder>(System.StringComparer.OrdinalIgnoreCase);
            
            foreach (var folder in targetFolders)
            {
                string simplePath = GetSimpleFolderPath(folder, targetRoot);
                progressForm.AddLog($"  {folder.Name} -> [{simplePath ?? "null"}] ({folder.Items.Count}封)");
                
                if (!string.IsNullOrEmpty(simplePath))
                {
                    targetFolderDict[simplePath] = folder;
                }
            }

            progressForm.AddLog("");
            progressForm.AddLog($"目标字典构建完成，包含 {targetFolderDict.Count} 个路径");

            // 第三步：对比源文件夹与目标
            int processedCount = 0;
            int totalFolders = sourceFolders.Count;
            int matchedCount = 0;
            int unmatchedCount = 0;

            foreach (var sourceFolder in sourceFolders)
            {
                processedCount++;

                // 获取简化路径
                string simplePath = GetSimpleFolderPath(sourceFolder, sourceRoot);
                if (string.IsNullOrEmpty(simplePath))
                    continue;

                // 每5个更新一次进度，或者最后一个
                    if (processedCount % 5 == 0 || processedCount == totalFolders)
                    {
                        int percent = (int)((double)processedCount / totalFolders * 100);
                        progressForm.UpdateProgress(processedCount, totalFolders, percent);
                        progressForm.UpdateStatus($"正在分析: {simplePath}");
                        System.Windows.Forms.Application.DoEvents();
                    }

                // 在目标中查找匹配
                Outlook.MAPIFolder targetFolder = null;
                bool foundInTarget = targetFolderDict.TryGetValue(simplePath, out targetFolder);

                if (foundInTarget) matchedCount++;
                else unmatchedCount++;

                // 获取邮件数量
                int sourceCount = GetMailItemCount(sourceFolder);
                int targetCount = foundInTarget ? GetMailItemCount(targetFolder) : 0;

                // 如果源比目标多，记录差异
                if (sourceCount > targetCount)
                {
                    result.Add(new FolderDiffInfo
                    {
                        FolderPath = simplePath,
                        SourceFolder = sourceFolder,
                        TargetFolder = targetFolder,
                        SourceCount = sourceCount,
                        TargetCount = targetCount,
                        DiffCount = sourceCount - targetCount
                    });
                }
            }

            // 输出分析结果统计到日志
            progressForm.AddLog("");
            progressForm.AddLog("=== 分析结果统计 ===");
            progressForm.AddLog($"源文件夹总数: {sourceFolders.Count}");
            progressForm.AddLog($"目标文件夹总数: {targetFolders.Count}");
            progressForm.AddLog($"匹配成功: {matchedCount}");
            progressForm.AddLog($"未匹配: {unmatchedCount}");
            progressForm.AddLog($"有差异的文件夹: {result.Count}");

            if (result.Count > 0)
            {
                progressForm.AddLog("");
                progressForm.AddLog("=== 差异详情 ===");
                foreach (var diff in result)
                {
                    progressForm.AddLog($"  {diff.FolderPath}: 源{diff.SourceCount} -> 目标{diff.TargetCount} (差{diff.DiffCount})");
                }
            }

            // 完成分析，等待用户关闭窗体
            progressForm.Complete($"发现 {result.Count} 个有差异的文件夹，请查看日志后关闭此窗口");
            
            // 使用ShowDialog方式等待用户手动关闭
            System.Windows.Forms.Application.DoEvents();
            progressForm.WaitForClose();

            return result;
        }

        /// <summary>
        /// 简单递归获取所有文件夹，无任何过滤
        /// </summary>
        private List<Outlook.MAPIFolder> GetAllFoldersSimple(Outlook.MAPIFolder root)
        {
            var result = new List<Outlook.MAPIFolder>();
            if (root == null) return result;

            result.Add(root);

            // 安全地遍历子文件夹（不释放COM对象，让调用方统一管理）
            try
            {
                foreach (Outlook.MAPIFolder folder in root.Folders)
                {
                    result.AddRange(GetAllFoldersSimple(folder));
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取子文件夹失败: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 获取文件夹的简化路径（用于跨PST文件匹配）
        /// 规则：忽略根文件夹名称，只使用其子路径
        /// 例如：PST1/收件箱/项目A 和 PST2/收件箱/项目A 都返回 "收件箱\项目A"
        /// </summary>
        private string GetSimpleFolderPath(Outlook.MAPIFolder folder, Outlook.MAPIFolder root)
        {
            if (folder == null) return null;
            
            // 使用EntryID比较来判断是否是同一个文件夹（COM对象引用比较不可靠）
            try
            {
                if (folder.EntryID == root.EntryID)
                    return "[根文件夹]";
            }
            catch { }

            var parts = new List<string>();
            Outlook.MAPIFolder current = folder;
            string rootEntryID = null;
            
            try
            {
                rootEntryID = root.EntryID;
            }
            catch { }

            // 向上遍历收集路径
            int depth = 0;
            while (current != null)
            {
                depth++;
                if (depth > 100) // 防止无限循环
                {
                    System.Diagnostics.Debug.WriteLine($"[警告] 路径遍历超过100层，可能有问题: {folder.Name}");
                    return null;
                }
                
                // 检查是否到达根（使用EntryID比较）
                try
                {
                    if (current.EntryID == rootEntryID)
                        break; // 到达根，停止遍历
                }
                catch { }
                
                parts.Insert(0, current.Name);
                
                try
                {
                    current = current.Parent as Outlook.MAPIFolder;
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[错误] 获取父文件夹失败 {folder.Name}: {ex.Message}");
                    break;
                }
            }

            if (parts.Count == 0)
                return "[根文件夹]";

            return string.Join("\\", parts);
        }

        /// <summary>
        /// 简单的差异分析（无进度条）
        /// </summary>
        private List<FolderDiffInfo> AnalyzeFolderDifferences(Outlook.MAPIFolder sourceRoot, Outlook.MAPIFolder targetRoot)
        {
            var result = new List<FolderDiffInfo>();
            var sourceFolders = GetAllFoldersSimple(sourceRoot);
            var targetFolders = GetAllFoldersSimple(targetRoot);

            // 构建目标文件夹字典
            var targetFolderDict = new Dictionary<string, Outlook.MAPIFolder>(System.StringComparer.OrdinalIgnoreCase);
            foreach (var folder in targetFolders)
            {
                string simplePath = GetSimpleFolderPath(folder, targetRoot);
                if (!string.IsNullOrEmpty(simplePath))
                {
                    targetFolderDict[simplePath] = folder;
                }
            }

            foreach (var sourceFolder in sourceFolders)
            {
                string simplePath = GetSimpleFolderPath(sourceFolder, sourceRoot);
                if (string.IsNullOrEmpty(simplePath))
                    continue;

                // 在目标中查找匹配
                Outlook.MAPIFolder targetFolder = null;
                targetFolderDict.TryGetValue(simplePath, out targetFolder);

                int sourceCount = GetMailItemCount(sourceFolder);
                int targetCount = targetFolder != null ? GetMailItemCount(targetFolder) : 0;

                if (sourceCount > targetCount)
                {
                    result.Add(new FolderDiffInfo
                    {
                        FolderPath = simplePath,
                        SourceFolder = sourceFolder,
                        TargetFolder = targetFolder,
                        SourceCount = sourceCount,
                        TargetCount = targetCount,
                        DiffCount = sourceCount - targetCount
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

            using (var syncProgress = new SyncProgressForm(AppVersion, totalEmails))
            {
                syncProgress.Show();
                syncProgress.AddLog($"开始同步 {selectedFolders.Count} 个文件夹，共 {totalEmails} 封邮件");

                // 在后台线程执行同步操作
                var syncTask = System.Threading.Tasks.Task.Run(() =>
                {
                    int processedEmails = 0;
                    int successCount = 0;
                    int failedCount = 0;
                    var failedEmails = new List<string>();

                    try
                    {
                        foreach (var folderDiff in selectedFolders)
                        {
                            if (syncProgress.IsCancelled)
                            {
                                syncProgress.AddLog("同步已取消");
                                break;
                            }

                            Outlook.MAPIFolder sourceFolder = folderDiff.SourceFolder;
                            Outlook.MAPIFolder targetFolder = folderDiff.TargetFolder;

                            syncProgress.AddLog($"开始处理文件夹: {folderDiff.FolderPath}");

                            // 如果目标文件夹不存在，创建它
                            if (targetFolder == null)
                            {
                                syncProgress.AddLog($"  目标文件夹不存在，正在创建...");
                                targetFolder = CreateFolderStructure(targetRoot, folderDiff.FolderPath);
                                if (targetFolder == null)
                                {
                                    string errorMsg = $"文件夹 {folderDiff.FolderPath} - 创建目标文件夹失败";
                                    failedEmails.Add(errorMsg);
                                    syncProgress.AddLog($"  ✗ {errorMsg}");
                                    continue;
                                }
                                syncProgress.AddLog($"  ✓ 目标文件夹创建成功");
                            }

                            // 获取目标文件夹中已有的邮件EntryID，避免重复
                            syncProgress.AddLog($"  正在扫描目标文件夹已有邮件...");
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
                            syncProgress.AddLog($"  已扫描 {existingEntryIds.Count} 封已有邮件");

                            // 复制邮件
                            int folderCopiedCount = 0;
                            int folderSkippedCount = 0;
                            int folderFailedCount = 0;
                            Outlook.Items sourceItems = sourceFolder.Items;
                            syncProgress.AddLog($"  开始复制邮件，源文件夹共 {sourceItems.Count} 封邮件");

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
                                    {
                                        folderSkippedCount++;
                                        continue;
                                    }

                                    // 复制邮件
                                    try
                                    {
                                        // 使用 Copy() 方法复制邮件，然后移动到目标文件夹
                                        var copiedMail = sourceMail.Copy();
                                        copiedMail.Move(targetFolder);
                                        successCount++;
                                        folderCopiedCount++;

                                        // 每20封邮件记录一次日志
                                        if (folderCopiedCount % 20 == 0)
                                        {
                                            syncProgress.AddLog($"    已复制 {folderCopiedCount} 封邮件");
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                        failedCount++;
                                        folderFailedCount++;
                                        string errorMsg = $"邮件: {sourceMail.Subject} | 错误: {ex.Message}";
                                        failedEmails.Add(errorMsg);
                                        syncProgress.AddLog($"    ✗ 复制失败: {sourceMail.Subject} - {ex.Message}");
                                        System.Diagnostics.Debug.WriteLine($"复制邮件失败: {sourceMail.Subject} - {ex.Message}");
                                    }

                                    processedEmails++;

                                    // 每处理5封邮件更新一次进度（提高响应性）
                                    if (processedEmails % 5 == 0)
                                    {
                                        syncProgress.UpdateProgress(processedEmails, totalEmails, folderDiff.FolderPath);
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
                            syncProgress.AddLog($"  ✓ 文件夹处理完成: 复制 {folderCopiedCount} 封, 跳过 {folderSkippedCount} 封, 失败 {folderFailedCount} 封");
                        }

                        syncProgress.AddLog("");
                        syncProgress.AddLog("=== 同步统计 ===");
                        syncProgress.AddLog($"总计: 成功 {successCount} 封, 失败 {failedCount} 封");
                        syncProgress.Complete();

                        // 返回结果
                        return new { SuccessCount = successCount, FailedCount = failedCount, FailedEmails = failedEmails };
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"同步过程出错: {ex.Message}");
                        syncProgress.AddLog($"✗ 同步过程出错: {ex.Message}");
                        syncProgress.Complete();
                        return new { SuccessCount = successCount, FailedCount = failedCount, FailedEmails = failedEmails };
                    }
                });

                // 等待任务完成，同时保持UI响应
                while (!syncTask.IsCompleted)
                {
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(50);
                }

                // 获取结果
                var result = syncTask.Result;

                // 显示结果
                ShowSyncResult(result.SuccessCount, result.FailedCount, result.FailedEmails);
            }
        }

        private Outlook.MAPIFolder CreateFolderStructure(Outlook.MAPIFolder targetRoot, string relativeFolderPath)
        {
            // 处理根文件夹的特殊标识符
            if (string.IsNullOrEmpty(relativeFolderPath) || relativeFolderPath == "__ROOT__")
                return targetRoot;

            string[] parts = relativeFolderPath.Split('\\');

            Outlook.MAPIFolder current = targetRoot;

            foreach (string part in parts)
            {
                Outlook.MAPIFolder nextFolder = null;

                foreach (Outlook.MAPIFolder folder in current.Folders)
                {
                    if (folder.Name.Equals(part, System.StringComparison.OrdinalIgnoreCase))
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
                        System.Diagnostics.Debug.WriteLine($"创建文件夹失败: {part} - {ex.Message}");
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
                using (var resultForm = new SaveResultForm(message.ToString()))
                {
                    resultForm.Text = $"jimsoutlooktools - 同步结果详情";
                    resultForm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show(message.ToString(), "jimsoutlooktools - 同步结果",
                    MessageBoxButtons.OK, failedCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            using (var aboutForm = new AboutForm())
            {
                aboutForm.ShowDialog();
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
