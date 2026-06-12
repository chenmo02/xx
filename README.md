# CC 实施工具箱

> 一款基于 WPF 的本地桌面工具，面向数据实施、数据校验、JSON 处理、办公输出等日常场景。离线可用，操作直观，旨在降低实施人员的数据处理和问题排查成本。

[![.NET](https://img.shields.io/badge/.NET-10-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2B-blue)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Version](https://img.shields.io/badge/version-2.1.2-orange)](https://github.com)

---

## 功能模块

<<<<<<< HEAD
| 模块 | 说明 |
|---|---|
| **首页概览** | 版本信息、快捷入口、系统信息面板，统一导航中枢 |
| **数据导入临时表** | Excel / CSV / DBF → SQL 临时表脚本，适配 SQL Server / PostgreSQL |
| **CSV 预览工具** | 稳定打开大文件分隔文本，支持搜索定位与编码识别 |
| **CSV 对比工具** | 按行号或复合主键比较两个 CSV 数据集的差异，结果可筛选导出 |
| **数据验证排查** | DDL 建表 + INSERT 数据 → 自动字段映射 → 逐行类型校验 → 导出错误报告 |
| **JSON 处理工具** | 格式化、校验、搜索、嵌套 GRID 浏览、表格编辑同步回写、导出 JSON/CSV |
| **JSON 对比工具** | 两个 JSON 文本的结构与内容差异比较，快速定位新增、删除和变更字段 |
| **发票打印工具** | 支持 PDF / OFD / 图片的排版与打印，多种模板、纸张方向、边距配置 |
| **Excalidraw 画板** | 流程图、草图和示意图绘制，适合与实施方案沟通配套使用 |
| **系统设置** | 应用级配置项与通用参数维护 |
=======
| 模块                | 说明                                                         |
| ------------------- | ------------------------------------------------------------ |
| **首页概览**        | 版本信息、快捷入口、系统信息面板，统一导航中枢               |
| **数据导入临时表**  | Excel / CSV / DBF → SQL 临时表脚本，适配 SQL Server / PostgreSQL |
| **CSV 预览工具**    | 稳定打开大文件分隔文本，支持搜索定位与编码识别               |
| **CSV 对比工具**    | 按行号或复合主键比较两个 CSV 数据集的差异，结果可筛选导出    |
| **数据验证排查**    | DDL 建表 + INSERT 数据 → 自动字段映射 → 逐行类型校验 → 导出错误报告 |
| **JSON 处理工具**   | 格式化、校验、搜索、嵌套 GRID 浏览、表格编辑同步回写、导出 JSON/CSV |
| **JSON 对比工具**   | 两个 JSON 文本的结构与内容差异比较，快速定位新增、删除和变更字段 |
| **发票打印工具**    | 支持 PDF / OFD / 图片的排版与打印，多种模板、纸张方向、边距配置 |
| **Excalidraw 画板** | 流程图、草图和示意图绘制，适合与实施方案沟通配套使用         |
| **系统设置**        | 应用级配置项与通用参数维护                                   |
>>>>>>> e60fac4f12d5712f3b1432b48e5441a1fc1b86ea

## 数据验证排查（核心模块）

这是当前功能最完整、迭代最多的模块，完整流程 4 步走：

```
结构输入 → 数据输入 → 字段映射 → 校验结果
```

### 1. 结构输入

- **DDL 粘贴**：直接粘贴 CREATE TABLE 语句，自动解析字段名、类型、必填、长度/精度
- **SQL 查询导入**：在数据库执行查询 → 导出 Excel → 导入结构
- 支持 SQL Server 和 PostgreSQL 两种数据库类型

### 2. 数据输入

- **INSERT 语句**：直接粘贴 INSERT INTO 语句，支持多表、批量 VALUES、字符串中特殊字符、SQL 注释
- **Excel 导入**：导入中间表数据 Excel，首行为表头
- 内置增强版 INSERT 解析器，按字符扫描而非正则截取，稳定处理千行级别的批量 SQL

### 3. 字段映射

- 源字段与目标字段自动匹配（支持精确/标准化/前缀剥离/语义优先/包含等多种匹配方式）
- 支持源字段映射、固定值、忽略三种模式
- 必填字段未映射红色高亮提醒
- 一键自动映射、全部确认、忽略自动生成 UUID

### 4. 校验结果

- 总行数 / 异常记录（去重）/ 错误项数（明细）/ 警告记录 / 耗时 概览
- 逐行逐字段校验：必填为空、字符超长、整数溢出、数值精度溢出、日期/时间/GUID/布尔/JSON 格式错误
- 可配置忽略项：整数格式、UUID 格式、日期时间格式、指定实际值
- 结果表格支持排序、右键复制单元格/行、Ctrl+C
- 导出 Excel 错误报告（含主键定位信息）

## 技术栈

```
.NET 10.0  |  WPF / XAML  |  CsvHelper  |  EPPlus  |  ExcelDataReader
Microsoft.Data.SqlClient  |  Npgsql  |  WebView2  |  System.Text.Json
```

## 项目结构

```
xx
├── WpfApp1/
│   ├── Views/              # 所有页面 (XAML + code-behind)
│   │   ├── HomePage.xaml              # 首页概览
│   │   ├── DataImportPage.xaml        # 数据导入临时表
│   │   ├── CsvViewerPage.xaml         # CSV 预览工具
│   │   ├── CsvComparePage.xaml        # CSV 对比工具
│   │   ├── DataValidationPage.xaml    # 数据验证排查（核心）
│   │   ├── JsonToolPage.xaml          # JSON 处理工具
│   │   ├── JsonDiffPage.xaml          # JSON 对比工具
│   │   ├── DrawBoardPage.xaml         # Excalidraw 画板
│   │   ├── InvoicePrintPage.xaml      # 发票打印工具
│   │   └── SettingsPage.xaml          # 系统设置
│   ├── Services/           # 业务逻辑层
│   │   ├── ValidationEngine.cs        # 核心校验引擎
│   │   ├── DdlParser.cs               # DDL 建表语句解析器
│   │   ├── InsertStatementParser.cs   # INSERT 语句解析器
│   │   ├── ValidationReportService.cs # Excel 校验报告生成
│   │   ├── JsonToolService.cs         # JSON 处理服务
│   │   ├── JsonGridParser.cs          # JSON → DataTable 解析
│   │   ├── InvoicePrintService.cs     # 票据打印服务
│   │   ├── CsvCompareService.cs       # CSV 对比服务
│   │   ├── SqlGeneratorService.cs     # SQL 临时表生成
│   │   ├── FieldMatcherService.cs     # 字段自动匹配
│   │   └── ...
│   ├── Models/             # 数据模型
│   ├── Themes/             # 主题与控件样式
│   ├── Converters/         # 值转换器
│   ├── MainWindow.xaml     # 主窗口（导航框架）
│   └── App.xaml            # 应用入口
├── docs/                   # 文档
└── README.md
```

## 本地运行

### 环境要求

- Windows 10 19041 及以上
- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)（含 Windows Desktop Runtime）

### 启动

```powershell
# 还原依赖
dotnet restore .\WpfApp1\WpfApp1.csproj --configfile .\NuGet.Config

# 运行
dotnet run --project .\WpfApp1\WpfApp1.csproj --launch-profile CCToolbox-DebugLocal
```

### 从本地 SDK 运行（无需全局安装 dotnet）

```powershell
# 安装 .NET SDK 到本地目录
powershell -NoProfile -ExecutionPolicy Bypass -File C:\tmp\dotnet-install.ps1 -Channel 10.0 -InstallDir C:\tmp\dotnet-10 -Architecture x64

# 编译运行
C:\tmp\dotnet-10\dotnet.exe restore .\WpfApp1\WpfApp1.csproj --configfile .\NuGet.Config
C:\tmp\dotnet-10\dotnet.exe build .\WpfApp1\WpfApp1.csproj --no-restore
$env:DOTNET_ROOT="C:\tmp\dotnet-10"
$env:PATH="C:\tmp\dotnet-10;$env:PATH"
.\WpfApp1\bin\Debug\net10.0-windows10.0.19041.0\CCToolbox.exe
```

### 发布单文件

```powershell
dotnet publish .\WpfApp1\WpfApp1.csproj -c Release -p:PublishProfile=SingleFile-win-x64
```

## NuGet 依赖

<<<<<<< HEAD
| 包 | 用途 |
|---|---|
| [CsvHelper](https://joshclose.github.io/CsvHelper/) | CSV 读写 |
| [EPPlus](https://www.epplussoftware.com/) | Excel 导出 |
| [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) | Excel 导入 |
| [Microsoft.Data.SqlClient](https://github.com/dotnet/SqlClient) | SQL Server 连接 |
| [Npgsql](https://www.npgsql.org/) | PostgreSQL 连接 |
=======
| 包                                                           | 用途                |
| ------------------------------------------------------------ | ------------------- |
| [CsvHelper](https://joshclose.github.io/CsvHelper/)          | CSV 读写            |
| [EPPlus](https://www.epplussoftware.com/)                    | Excel 导出          |
| [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) | Excel 导入          |
| [Microsoft.Data.SqlClient](https://github.com/dotnet/SqlClient) | SQL Server 连接     |
| [Npgsql](https://www.npgsql.org/)                            | PostgreSQL 连接     |
>>>>>>> e60fac4f12d5712f3b1432b48e5441a1fc1b86ea
| [WebView2](https://www.nuget.org/packages/Microsoft.Web.WebView2) | Excalidraw 画板嵌入 |

## 更新记录

### v2.1.3 (2026-06-12)

- 优化已知问题
- 调整UI显示
<<<<<<< HEAD
=======
- 本次主要改了数据导入页：
修复 JSON 导出不使用默认导出路径的问题。
删除 ApplyDatabaseHint 里被覆盖的重复提示赋值。
从数据导入页面移除“每批行数”输入框。
生成 SQL 时改为使用设置页里的默认批量行数，异常时回退 1000。
>>>>>>> e60fac4f12d5712f3b1432b48e5441a1fc1b86ea

---

**by 悲伤番茄** · 持续迭代中，欢迎 Issue & PR。
