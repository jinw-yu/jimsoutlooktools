# Jim's Outlook Tools

[![版本](https://img.shields.io/badge/版本-v1.0.0-blue)](https://gitee.com/jim/jimsoutlooktools)
[![协议](https://img.shields.io/badge/协议-MIT-green)](LICENSE)

Jim的outlook工具集，当前仅开发完成一个简单易用的 Outlook 附件保存工具，帮助您快速批量保存邮件附件，并按年月自动分类整理。

## 功能特性

- 📧 **批量保存附件**：一键保存指定日期范围内的所有邮件附件
- 📁 **智能分类**：按 `年月(yyyyMM)` 自动创建子文件夹，如 `202602`、`202603`
- 🏷️ **时间戳命名**：文件名格式为 `时间戳(精确到毫秒)_原文件名`，确保唯一性
  - 示例：`20250302_143052_123_invoice.pdf`
- 🖼️ **智能过滤**：自动跳过小于 100KB 的图片文件（邮件内联图标、表情等），只保存真正的附件
- 🔄 **防重复**：自动检测已存在的文件，避免重复保存
- 📅 **日期范围选择**：精确选择起始和结束日期，包含整天数据
- 📊 **进度显示**：实时显示保存进度和统计结果

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

1. 打开 Outlook，在工具栏找到 **"邮件附件工具"** 工具栏
2. 点击 **"保存附件"** 按钮
3. 在弹出的对话框中：
   - 点击 **"浏览..."** 选择附件保存的根文件夹
   - 选择 **起始日期** 和 **结束日期**
   - 点击 **"确定"** 开始保存
4. 等待进度完成，查看保存统计结果

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

### 最新版本：v1.0.2
- 优化内存管理，解决大日期范围处理时内存不足问题

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
