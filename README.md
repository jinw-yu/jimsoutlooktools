# JTools

[![版本](https://img.shields.io/badge/版本-v1.0.5-blue)](https://gitee.com/jim/jimsoutlooktools)
[![协议](https://img.shields.io/badge/协议-MIT-green)](LICENSE)

> ⚠️ **Vibe Coding 警告**
> 
> 本项目采用 **Vibe Coding** 方式开发（即 AI 辅助编程，人机协作快速迭代）。代码可能包含未充分考虑边界情况、未完整测试或不符合传统软件工程规范的部分。
> 
> **使用者后果自负**：在使用本工具前，请务必备份重要数据。开发者不对因使用本工具导致的任何数据丢失、损坏或其他损失承担责任。

---

**Jim的outlook工具集。**

## 功能特性

### 保存附件
- 📧 **批量保存附件**：一键保存指定日期范围内的邮件附件，支持选择**收件箱**和/或**已发送邮件**
- 📁 **智能分类**：按 `年月(yyyyMM)` 自动创建子文件夹，如 `202602`、`202603`
- 🏷️ **时间戳命名**：文件名格式为 `时间戳(精确到毫秒)_原文件名`，确保唯一性
  - 示例：`20250302_143052_123_invoice.pdf`
- 🖼️ **智能过滤**：自动跳过小于 100KB 的图片文件（邮件内联图标、表情等），只保存真正的附件
- 🔄 **防重复**：自动检测已存在的文件，避免重复保存
- 📅 **日期范围选择**：精确选择起始和结束日期，包含整天数据
- 📊 **进度显示**：实时显示保存进度和统计结果，支持中途取消
- ⛔ **容错处理**：单个附件保存失败不影响其他附件，最后统一显示失败详情
- ℹ️ **关于版本信息**：点击工具栏"关于"按钮查看完整版本信息

### 下载联机
- ☁️ **联机存档同步**：从 Office 365 联机存档同步邮件到本地 PST
- 📂 **数据文件选择**：自动识别联机存档和本地 PST 文件
- 📈 **差异分析**：自动分析并显示每个文件夹的邮件数量差异
- ✅ **选择性同步**：支持选择特定文件夹进行同步
- 📊 **进度显示**：实时显示同步进度，支持中途取消
- 🛡️ **错误隔离**：个别邮件同步失败不影响整体流程

## 安装要求

- Microsoft Outlook 2010 或更高版本
- .NET Framework 4.7.2 或更高版本
- VSTO Runtime（Visual Studio Tools for Office）

## 安装方法

### 方法一：使用安装包
1. 下载最新版本的 [Releases](../../releases)
2. 运行 `setup.exe` 进行安装
3. 重启 Outlook 即可使用

### 方法二：手动安装（开发者）
1. 克隆仓库
   ```bash
   git clone https://gitee.com/jim/jimsoutlooktools.git
   ```
2. 使用 Visual Studio 打开 `jimsoutlooktools.slnx`
3. 编译并发布项目
4. 在 Outlook 中启用加载项

## 使用方法

### 保存附件

1. 打开 Outlook，在工具栏找到 **"JTools"** 工具栏
2. 点击 **"保存附件"** 按钮
3. 在弹出的对话框中：
   - 点击 **"浏览..."** 选择附件保存的根文件夹
   - 选择 **起始日期** 和 **结束日期**
   - 勾选要处理的文件夹（**收件箱** 和/或 **已发送邮件**）
   - 点击 **"确定"** 开始保存
4. 等待进度完成，查看保存统计结果

> **注意**：如果同时选择"收件箱"和"已发送邮件"，某些邮件（如发送给自己的邮件）可能同时存在于两个文件夹中，第二次遇到时会因为文件已存在而被跳过，这是正常的防重复机制。

### 下载联机（同步邮件）

1. 打开 Outlook，在工具栏找到 **"JTools"** 工具栏
2. 点击 **"下载联机"** 按钮
3. 选择源数据文件（联机存档）和目标数据文件（本地 PST）
4. 等待差异分析完成
5. 在差异分析结果中勾选要同步的文件夹
6. 点击 **"开始同步"** 开始下载邮件
7. 等待同步完成，查看同步结果

## 文件命名规则

保存的附件文件名格式：
```
{邮件接收时间}_{原文件名}
```

例如：
- `20250302_143052_123_invoice.pdf`
- `20250302_143052_456_report.xlsx`

文件夹结构：
```
📁 您选择的根文件夹/
├── 📁 202601/
│   ├── 20260115_093012_001_document.pdf
│   └── 20260120_143052_123_invoice.pdf
├── 📁 202602/
│   └── 20260228_163022_456_report.xlsx
└── 📁 202603/
    └── 20260301_101512_789_photo.jpg
```

## 技术栈

- **语言**：C# 
- **框架**：.NET Framework 4.7.2
- **平台**：VSTO (Visual Studio Tools for Office)
- **目标应用**：Microsoft Outlook

## 项目结构

```
jimsoutlooktools/
├── ThisAddIn.cs          # 主程序代码
├── ThisAddIn.Designer.cs # 设计器文件
├── jimsoutlooktools.csproj   # 项目文件
├── README.md             # 项目说明
└── CHANGELOG.md          # 更新日志
```

## 开发说明

### 环境配置
1. 安装 Visual Studio 2019/2022
2. 安装 Office 开发工具（VSTO）
3. 安装 .NET Framework 4.7.2 目标包

### 调试运行
1. 设置项目为启动项目
2. 按 F5 启动调试
3. Visual Studio 会自动启动 Outlook 并加载插件

## 更新日志

详见 [CHANGELOG.md](CHANGELOG.md) 文件。

## 贡献指南

欢迎提交 Issue 和 Pull Request！

1. Fork 本仓库
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

## 开源协议

本项目基于 [MIT](LICENSE) 协议开源。

## 作者

**Jim** - [jinw.yu@qq.com](mailto:jinw.yu@qq.com)

## 致谢

- 感谢 Microsoft 提供 VSTO 开发平台
- 感谢所有贡献者和用户的支持

---

如果这个项目对您有帮助，请给个 ⭐ Star 支持一下！谢谢！
