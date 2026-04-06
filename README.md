# CC 实施工具箱

> 高效、安全的数据实施辅助工具，专为实施工程师打造的 Windows 桌面工具箱。

![.NET](https://img.shields.io/badge/.NET-10.0-blue?logo=dotnet)
![Platform](https://img.shields.io/badge/Platform-Windows_x64-0078d4?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-1.0.0-orange)

---

## ✨ 功能概览

CC 实施工具箱集成了实施工作中常用的数据处理、开发辅助和系统工具，采用左侧导航 + 右侧内容区的经典布局。

### 📊 数据导入临时表

将 Excel / CSV 文件快速转换为数据库临时表 SQL 语句，**无需直接连接数据库**。

- 支持 `.xlsx`、`.xls`、`.csv` 文件导入
- 支持 Excel 多 Sheet 选择
- 自动检测 CSV 文件编码（UTF-8 / GBK / GB2312 等）
- 自动识别列名和数据类型
- 生成 `CREATE TEMP TABLE` + `INSERT INTO` 语句
- 支持 **PostgreSQL** 和 **SQL Server** 两种数据库方言
- 可配置每批 INSERT 行数
- 一键复制 SQL 或保存为 `.sql` 文件

### 📋 CSV 预览工具

无损打开 CSV 文件，避免 Excel 打开 CSV 时自动转换数据格式（如身份证号、长数字等）。

- 以只读方式打开，不修改原文件
- 使用 `FileShare.ReadWrite` 模式，不锁定文件
- 自动检测文件编码
- 支持关键字搜索查找数据
- DataGrid 表格展示，支持排序

### 📄 JSON 处理工具

集编辑、校验、格式化、嵌套网格查看于一体的 JSON 工具，对标 [jsongrid.com](https://jsongrid.com)。

- **格式化**：一键美化 / 压缩 JSON
- **校验**：实时语法检查，定位错误位置
- **嵌套网格**：将 JSON 渲染为可视化嵌套表格
  - 递归渲染多层嵌套结构
  - 简单数组直接显示值列表（不显示索引）
  - 支持懒加载展开，避免大数据卡顿
  - 每个嵌套节点支持 `···` 按钮导出为 CSV

### 🔀 JSON 对比工具

深度对比两个 JSON 的结构与值差异，精确到每个字段级别。

- **深度递归对比**：逐层递归比较 Object、Array、基本类型
- **四种差异识别**：
  - 🟢 **新增** — B 中有而 A 中没有的字段
  - 🔴 **删除** — A 中有而 B 中没有的字段
  - 🟠 **值修改** — 同一路径下值发生变化
  - 🟣 **类型变化** — 同一路径下数据类型改变（如 string → number）
- **忽略字段顺序**：JSON 对象按 key 比较，不受书写顺序影响
- **JSONPath 路径定位**：每条差异显示完整路径（如 `$.user.address[0].city`）
- **差异筛选**：按类型勾选过滤，快速聚焦关注的变更
- **差异报告**：一键复制或导出为 `.txt` 文件，方便分享
- **便捷操作**：支持文件导入、一键美化、左右交换、清空

### ⚙️ 系统设置

全局配置管理 + 7 个实用开发小工具。

#### 配置管理
- SQL 生成偏好（默认数据库类型、临时表前缀、批量行数）
- 默认导出路径设置
- 配置持久化保存为 `settings.json`

#### 实用小工具

| 工具 | 说明 |
|------|------|
| 🕐 **时间戳转换** | Unix 时间戳 ↔ 日期时间互转，实时显示当前时间戳，自动识别秒/毫秒 |
| 🔐 **Base64 编解码** | 文本 ↔ Base64 双向转换 |
| 🆔 **UUID 生成器** | 批量生成 UUID/GUID，支持大写、无连字符选项 |
| 🔒 **文本哈希** | 一键计算 MD5 / SHA1 / SHA256 |
| 🔗 **URL 编解码** | URL Percent-Encoding 编码与解码 |
| 🔍 **正则测试** | 正则表达式匹配测试，显示匹配位置与捕获组 |
| 📝 **文本对比** | 逐行 Diff 对比两段文本，红绿色标记差异 |

---

## 🖼️ 界面预览

```
┌──────────────────────────────────────────────────────┐
│  CC 实施工具箱                                        │
├────────────┬─────────────────────────────────────────┤
│            │                                         │
│  数据工具   │                                         │
│  🏠 首页概览│         右 侧 内 容 区                    │
│  📊 数据导入│                                         │
│  📋 CSV预览 │         (根据左侧导航切换页面)             │
│            │                                         │
│  开发工具   │                                         │
│  📄 JSON工具│                                         │
│  🔀 JSON对比│                                         │
│            │                                         │
│  系统      │                                         │
│  ⚙️ 系统设置│                                         │
│            │                                         │
│ v1.0.0     │                                         │
└────────────┴─────────────────────────────────────────┘
```

---

## 🏗️ 技术栈

| 技术 | 版本 | 用途 |
|------|------|------|
| **C# / WPF** | .NET 10 | 桌面 UI 框架 |
| **EPPlus** | 8.5.1 | Excel 文件读取（.xlsx） |
| **CsvHelper** | 33.1.0 | CSV 文件解析 |
| **Npgsql** | 10.0.2 | PostgreSQL 数据类型支持 |
| **Microsoft.Data.SqlClient** | 7.0.0 | SQL Server 数据类型支持 |

---

## 📁 项目结构

```
WpfApp1/
├── WpfApp1.slnx                    # 解决方案文件
└── WpfApp1/
    ├── WpfApp1.csproj               # 项目文件
    ├── favicon.ico                  # 应用图标
    ├── App.xaml / App.xaml.cs        # 应用入口 & 全局异常捕获
    ├── MainWindow.xaml / .cs         # 主窗口（左侧导航 + 右侧 Frame）
    ├── Views/
    │   ├── HomePage.xaml / .cs       # 首页概览（功能卡片 + 系统信息）
    │   ├── DataImportPage.xaml / .cs # 数据导入临时表（SQL 生成器）
    │   ├── CsvViewerPage.xaml / .cs  # CSV 无损预览工具
    │   ├── JsonToolPage.xaml / .cs   # JSON 处理工具（嵌套网格）
    │   ├── JsonDiffPage.xaml / .cs   # JSON 深度对比工具
    │   └── SettingsPage.xaml / .cs   # 系统设置 & 实用小工具
    ├── Services/
    │   ├── FileParserService.cs      # Excel / CSV 文件解析服务
    │   ├── SqlGeneratorService.cs    # SQL 语句生成服务（PG / SQL Server）
    │   ├── JsonToolService.cs        # JSON 格式化 / 校验 / 树形解析
    │   └── JsonGridParser.cs         # JSON 嵌套网格节点模型与解析器
    └── Models/
        └── (数据模型)
```

---

## 🚀 快速开始

### 环境要求

- **运行**：Windows 10/11 x64（使用发布版 EXE 无需安装 .NET）
- **开发**：.NET 10 SDK

### 直接运行

下载 `publish/` 目录下的 `WpfApp1.exe`，双击即可运行，无需安装任何依赖。

### 从源码构建

```bash
# 克隆项目
git clone <repository-url>
cd WpfApp1

# 调试运行
dotnet run --project WpfApp1/WpfApp1.csproj

# 发布为单文件 EXE（自包含，约 68MB）
cd WpfApp1
dotnet publish -c Release -r win-x64 --self-contained true ^
    -p:PublishSingleFile=true ^
    -p:IncludeNativeLibrariesForSelfExtract=true ^
    -p:EnableCompressionInSingleFile=true ^
    -o ..\publish
```

发布后的 EXE 位于 `publish/WpfApp1.exe`，包含完整 .NET 运行时，可在任意 Windows x64 机器上直接运行。

---

## 🎨 设计规范

- **主色调**：`#4E6EF2`（百度蓝）
- **背景色**：`#F5F6F8`（浅灰）
- **卡片背景**：`#FFFFFF`（纯白）
- **边框色**：`#E5E5E5`
- **选中高亮**：`#EDF0FF`
- **字体**：系统默认 + `Consolas`（代码场景）

---

## 🔧 技术细节

### 剪贴板方案

使用 Win32 原生 API（`OpenClipboard` / `SetClipboardData`）+ WPF `Clipboard` 回退，解决 OLE 剪贴板锁定导致的闪退问题。

```csharp
[DllImport("user32.dll")] static extern bool OpenClipboard(IntPtr hWndNewOwner);
[DllImport("user32.dll")] static extern bool CloseClipboard();
[DllImport("user32.dll")] static extern bool EmptyClipboard();
[DllImport("user32.dll")] static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
```

### 文件读取

所有文件读取使用 `FileShare.ReadWrite` 模式，避免文件被其他程序占用时报错：

```csharp
new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
```

### 全局异常捕获

在 `App.xaml.cs` 中统一捕获三类异常，防止程序闪退：

- `DispatcherUnhandledException` — UI 线程异常
- `AppDomain.UnhandledException` — 非 UI 线程异常
- `TaskScheduler.UnobservedTaskException` — Task 异步异常

### JSON 深度对比算法

递归遍历 JSON 树，按 JSONPath 路径逐节点比较：

- **Object**：取两侧 key 集合，差集为新增/删除，交集递归比较
- **Array**：按索引逐元素对比，长度差异部分标记为新增/删除
- **基本类型**：直接值比较，类型不同时标记为类型变化

---

## 📄 License

MIT License © 2026 CC Team
