' ---------------- 文件信息 ----------------
' 文件名: outlook.bas
' 描述: 包含用于保存附件的 VBA 脚本
' 作者: 未知
' 创建日期: 未知
' 最后修改日期: 2025-12-15
' 版本: 1.0.0

Option Explicit

' ---------------- 配置 ----------------
' PROGRESS_SCRIPT_NAME: PowerShell 脚本的文件名
' PROGRESS_STATUS_NAME: 进度状态文件的文件名
Private Const PROGRESS_SCRIPT_NAME As String = "vba_save_attachments_progress.ps1"
Private Const PROGRESS_STATUS_NAME As String = "progress_status.txt"

' 全局变量
' gPsScriptPath: PowerShell 脚本路径
' gStatusFilePath: 状态文件路径
' gLastAction: 上一次操作的描述
' gLogPath: 日志文件路径
Private gPsScriptPath As String
Private gStatusFilePath As String
Private gLastAction As String
Private gLogPath As String

' ---------------- 日志函数 ----------------
' LogAppend: 将日志信息追加到日志文件中
' 参数:
'   txt - 要追加的日志内容
Private Sub LogAppend(ByVal txt As String)
    On Error Resume Next
    If Len(gLogPath) = 0 Then gLogPath = Environ("TEMP") & "\vba_progress_log.txt"
    Dim ff As Integer
    ff = FreeFile
    Open gLogPath For Append As #ff
    Print #ff, Format(Now, "yyyy-mm-dd HH:nn:ss") & "  " & txt
    Close #ff
    On Error GoTo 0
End Sub

' ---------------- 写文本文件（覆盖） ----------------
' WriteTextFile: 将内容写入指定路径的文本文件（覆盖模式）
' 参数:
'   path - 文件路径
'   content - 要写入的内容
Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, content
    Close #ff
    On Error GoTo 0
End Sub

' ---------------- 生成 PowerShell 脚本 ----------------
' BuildPSScriptForProgress: 生成用于显示附件保存进度的 PowerShell 脚本
' 参数:
'   statusFile - 状态文件路径
' 返回值:
'   生成的 PowerShell 脚本内容
Private Function BuildPSScriptForProgress(ByVal statusFile As String) As String
    Dim s As String
    s = ""
    s = s & "$statusFile = '" & Replace(statusFile, "'", "''") & "'" & vbCrLf
    s = s & "Write-Host '附件保存进度（按附件计数）' -ForegroundColor Cyan" & vbCrLf
    s = s & "Write-Host 'PS 窗口将在任务完成后保留，关闭窗口可结束查看。'" & vbCrLf
    s = s & "if (!(Test-Path $statusFile)) { New-Item -Path $statusFile -ItemType File -Force | Out-Null }" & vbCrLf
    s = s & "while ($true) {" & vbCrLf
    s = s & "  try {" & vbCrLf
    s = s & "    $txt = Get-Content -Path $statusFile -Raw -ErrorAction SilentlyContinue" & vbCrLf
    s = s & "    if ($txt) {" & vbCrLf
    s = s & "      $parts = $txt -split '\|\|'" & vbCrLf
    s = s & "      $current = if ($parts.Length -ge 1 -and $parts[0] -ne '') { [int]$parts[0] } else { 0 }" & vbCrLf
    s = s & "      $msg = if ($parts.Length -ge 2) { $parts[1] } else { '' }" & vbCrLf
    s = s & "      Write-Host ('[{0}] {1}' -f $current, $msg)" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "  } catch {}" & vbCrLf
    s = s & "  Start-Sleep -Milliseconds 700" & vbCrLf
    s = s & "}" & vbCrLf
    BuildPSScriptForProgress = s
End Function

' ---------------- 更新进度文件 ----------------
' UpdateProgressFile: 更新进度文件的内容
' 参数:
'   current - 当前进度（已保存的附件数量）
'   message - 进度消息
Private Sub UpdateProgressFile(ByVal current As Long, ByVal message As String)
    On Error Resume Next
    If Len(gStatusFilePath) = 0 Then Exit Sub
    Dim ff As Integer
    ff = FreeFile
    Open gStatusFilePath For Output As #ff
    Print #ff, CStr(current) & "||" & message
    Close #ff
    On Error GoTo 0
End Sub

' ---------------- 文件名安全化 ----------------
' SanitizeFileName: 将文件名安全化，替换非法字符
' 参数:
'   name - 原始文件名
' 返回值:
'   安全化后的文件名
Private Function SanitizeFileName(ByVal name As String) As String
    Dim s As String
    s = name
    s = Replace(s, "\", "-")
    s = Replace(s, "/", "-")
    s = Replace(s, ":", "-")
    s = Replace(s, "*", "-")
    s = Replace(s, "?", "-")
    s = Replace(s, """", "-")
    s = Replace(s, "<", "-")
    s = Replace(s, ">", "-")
    s = Replace(s, "|", "-")
    If Len(s) > 180 Then s = Left(s, 180)
    SanitizeFileName = s
End Function

' ---------------- 递归创建文件夹 ----------------
' MkDirRecursive: 递归创建文件夹
' 参数:
'   fullPath - 文件夹完整路径
Private Sub MkDirRecursive(ByVal fullPath As String)
    On Error Resume Next
    If fullPath = "" Then Exit Sub
    If Dir(fullPath, vbDirectory) <> "" Then Exit Sub ' 如果文件夹已存在，直接退出
    Dim parent As String
    parent = Left(fullPath, InStrRev(fullPath, "\") - 1)
    If parent <> "" And Dir(parent, vbDirectory) = "" Then MkDirRecursive parent
    If Dir(fullPath, vbDirectory) = "" Then MkDir fullPath
    On Error GoTo 0
End Sub

' ---------------- 选择文件夹（Shell 对话框） ----------------
' SelectFolderOutlook: 显示选择文件夹对话框
' 参数:
'   prompt - 提示信息
'   initialPath - 初始路径
' 返回值:
'   选择的文件夹路径
Private Function SelectFolderOutlook(prompt As String, initialPath As String) As String
    On Error Resume Next
    Dim shell As Object, folder As Object, selectedPath As String
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, prompt, 1, initialPath)
    If Not folder Is Nothing Then
        selectedPath = folder.Self.path
        If Right(selectedPath, 1) <> "\" Then selectedPath = selectedPath & "\"
    Else
        selectedPath = ""
    End If
    Set folder = Nothing
    Set shell = Nothing
    SelectFolderOutlook = selectedPath
    On Error GoTo 0
End Function

' ---------------- 选择日期范围（输入框） ----------------
' SelectDateRange: 选择日期范围
' 参数:
'   prompt - 提示信息
'   startDate - 起始日期（输出参数）
'   endDate - 结束日期（输出参数）
' 返回值:
'   是否选择成功
Private Function SelectDateRange(prompt As String, ByRef startDate As Date, ByRef endDate As Date) As Boolean
    On Error Resume Next
    Dim inputStart As String, inputEnd As String
    inputStart = InputBox(prompt & "\n请输入起始日期 (格式: yyyy-mm-dd):", "选择日期范围")
    If inputStart = "" Then Exit Function
    inputEnd = InputBox(prompt & "\n请输入结束日期 (格式: yyyy-mm-dd):", "选择日期范围")
    If inputEnd = "" Then Exit Function

    Dim tempStart As Date, tempEnd As Date
    tempStart = CDate(inputStart)
    tempEnd = CDate(inputEnd)
    If tempStart > tempEnd Then
        MsgBox "起始日期不能晚于结束日期！", vbExclamation
        Exit Function
    End If

    startDate = tempStart
    endDate = tempEnd
    SelectDateRange = True
    On Error GoTo 0
End Function

' ---------------- 主过程：保存收件箱附件并启动 PowerShell 进度（不可取消） ----------------
' SaveInboxAttachments_WithPSProgress: 主过程，保存收件箱附件并显示 PowerShell 进度
Public Sub SaveInboxAttachments_WithPSProgress()
    On Error GoTo ErrHandler
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olInbox As Outlook.MAPIFolder
    Dim itm As Object
    Dim mail As Outlook.MailItem
    Dim att As Outlook.attachment
    Dim saveRoot As String
    Dim folderYM As String
    Dim targetFolder As String
    Dim totalAttachments As Long
    Dim savedCount As Long
    Dim tmpPSPath As String
    Dim psScript As String
    Dim wsh As Object
    Dim startTime As Date

    ' 立即初始化日志与路径（确保早期变量被设置）
    gLogPath = Environ("TEMP") & "\vba_progress_log.txt"
    On Error Resume Next
    Dim ff0 As Integer
    ff0 = FreeFile
    Open gLogPath For Output As #ff0
    Print #ff0, "Log start: " & Format(Now, "yyyy-mm-dd HH:nn:ss")
    Close #ff0
    On Error GoTo ErrHandler

    gLastAction = "进入主过程"
    LogAppend gLastAction

    ' 预设 PS 路径与状态文件路径
    gPsScriptPath = Environ("TEMP") & "\" & PROGRESS_SCRIPT_NAME
    gStatusFilePath = Left(gPsScriptPath, InStrRev(gPsScriptPath, "\") - 1) & "\" & PROGRESS_STATUS_NAME
    LogAppend "PS 脚本路径预设: " & gPsScriptPath
    LogAppend "进度文件路径预设: " & gStatusFilePath

    ' 1) 选择根文件夹
    gLastAction = "选择保存根文件夹"
    LogAppend gLastAction
    saveRoot = SelectFolderOutlook("请选择附件保存的根文件夹", Environ("USERPROFILE") & "\Desktop\附件\")
    If saveRoot = "" Then
        MsgBox "未选择保存路径，操作取消。", vbInformation
        Exit Sub
    End If
    If Right(saveRoot, 1) <> "\" Then saveRoot = saveRoot & "\"
    LogAppend "选择保存根路径: " & saveRoot

    ' 2) 初始化 Outlook 对象
    gLastAction = "初始化 Outlook 对象"
    LogAppend gLastAction
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    If olApp Is Nothing Then Err.Raise vbObjectError + 1000, , "无法创建或获取 Outlook.Application"
    On Error GoTo ErrHandler
    Set olNS = olApp.GetNamespace("MAPI")
    Set olInbox = olNS.GetDefaultFolder(olFolderInbox)
    LogAppend "Outlook 初始化完成"

    ' 3) 统计总附件数（预扫描）
    gLastAction = "预扫描统计附件总数"
    LogAppend gLastAction
    totalAttachments = 0
    For Each itm In olInbox.items
        If TypeName(itm) = "MailItem" Then
            Set mail = itm
            totalAttachments = totalAttachments + mail.Attachments.Count
        End If
    Next itm
    LogAppend "总附件数: " & CStr(totalAttachments)

    ' 4) 生成并写入 PowerShell 脚本及初始状态文件
    gLastAction = "写入 PowerShell 脚本"
    LogAppend gLastAction
    psScript = BuildPSScriptForProgress(gStatusFilePath)
    WriteTextFile gPsScriptPath, psScript
    LogAppend "已写入 PS 脚本: " & gPsScriptPath

    gLastAction = "创建初始进度文件"
    LogAppend gLastAction
    UpdateProgressFile 0, "准备中"

    ' 5) 启动 PowerShell（显示窗口并保持）
    gLastAction = "启动 PowerShell 窗口"
    LogAppend gLastAction
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "powershell -NoProfile -ExecutionPolicy Bypass -File """ & gPsScriptPath & """", 1, False
    Set wsh = Nothing
    LogAppend "PowerShell 已启动（窗口应已打开）"

    ' 6) 选择日期范围
    Dim startDate As Date, endDate As Date
    If Not SelectDateRange("请选择附件保存的日期范围", startDate, endDate) Then
        MsgBox "未选择日期范围，操作取消。", vbInformation
        Exit Sub
    End If
    LogAppend "选择日期范围: " & Format(startDate, "yyyy-mm-dd") & " 至 " & Format(endDate, "yyyy-mm-dd")

    ' 7) 遍历并保存附件（按 yyyyMM）
    startTime = Now
    savedCount = 0
    gLastAction = "开始保存附件"
    LogAppend gLastAction

    Dim items As Object
    Set items = olInbox.items
    items.Sort "[ReceivedTime]", True
    Dim itmObj As Object

    For Each itmObj In items
        If TypeName(itmObj) = "MailItem" Then
            Set mail = itmObj
            If mail.ReceivedTime >= startDate And mail.ReceivedTime <= endDate Then
                If mail.Attachments.Count > 0 Then
                    On Error Resume Next
                    folderYM = Format(CDate(mail.ReceivedTime), "yyyyMM")
                    If folderYM = "" Then folderYM = Format(Now, "yyyyMM")
                    On Error GoTo ErrHandler

                    targetFolder = saveRoot & folderYM
                    If Dir(targetFolder, vbDirectory) = "" Then
                        gLastAction = "创建文件夹: " & targetFolder
                        LogAppend gLastAction
                        MkDirRecursive targetFolder
                    End If

                    For Each att In mail.Attachments
                        gLastAction = "准备保存附件: " & att.FileName
                        LogAppend gLastAction
                        Dim origName As String, safeName As String, baseName As String, extName As String
                        origName = att.FileName
                        safeName = SanitizeFileName(origName)
                        Dim posDot As Long
                        posDot = InStrRev(safeName, ".")
                        If posDot > 0 Then
                            baseName = Left(safeName, posDot - 1)
                            extName = Mid(safeName, posDot)
                        Else
                            baseName = safeName
                            extName = ""
                        End If

                        Dim targetPath As String
                        targetPath = targetFolder & "\" & safeName
                        If Dir(targetPath) <> "" Then
                            LogAppend "文件已存在，跳过: " & targetPath
                            GoTo NextAttachment ' 跳过当前附件
                        End If

                        ' 保存附件
                        On Error Resume Next
                        att.SaveAsFile targetPath
                        If Err.Number <> 0 Then
                            LogAppend "保存失败: Err=" & CStr(Err.Number) & " Desc=" & Err.Description & " File=" & targetPath
                            Err.Clear
                        Else
                            savedCount = savedCount + 1
                            LogAppend "已保存: " & targetPath & " (已保存 " & CStr(savedCount) & ")"
                            UpdateProgressFile savedCount, "保存: " & targetPath
                        End If
                        On Error GoTo ErrHandler
NextAttachment:
                    Next att
                End If
            End If
        End If
    Next itmObj

    ' 8) 写入完成状态（PowerShell 窗口将显示最终信息并保持打开）
    gLastAction = "写入完成状态"
    LogAppend gLastAction
    UpdateProgressFile savedCount, "完成，已保存 " & CStr(savedCount) & " 个附件，开始时间：" & Format(startTime, "yyyy-mm-dd HH:nn:ss")
    MsgBox "保存完成，已保存附件：" & CStr(savedCount), vbInformation

    Exit Sub

ErrHandler:
    On Error Resume Next
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    If errNum = 0 Then errDesc = "（Err.Number = 0，可能被 On Error Resume Next 清除或无详细描述）"
    LogAppend "运行时错误: " & CStr(errNum) & " - " & errDesc
    LogAppend "最后操作: " & gLastAction
    LogAppend "PS 脚本路径: " & gPsScriptPath
    LogAppend "进度文件路径: " & gStatusFilePath
    MsgBox "运行时错误：" & CStr(errNum) & " - " & errDesc & vbCrLf & "最后操作：" & gLastAction & vbCrLf & "请查看日志：" & gLogPath, vbCritical
    Resume Next
End Sub
