# CC 实施工具箱

面向实施、数据处理和日常开发辅助场景的 Windows 桌面工具箱。

![.NET](https://img.shields.io/badge/.NET-10.0-blue?logo=dotnet)
![Platform](https://img.shields.io/badge/Platform-Windows_x64-0078d4?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

## 项目简介

`CC 实施工具箱` 是一个基于 WPF 的桌面应用，聚合了数据导入、CSV 预览、JSON 处理、JSON 对比、Excalidraw 画板、发票打印和若干实用小工具。

当前主程序程序集名称为 `CCToolbox`，支持打包为 `win-x64` 自包含单文件 EXE。

## 主要功能

### 数据导入临时表

- 支持导入 `.xlsx`、`.xls`、`.csv`、`.dbf`
- 支持 Excel 多 Sheet
- CSV 自动识别编码和分隔符
- DBF 支持手动切换编码
- 支持在表格中编辑数据后重新生成 SQL
- 支持导出当前数据为 CSV / JSON

支持的目标数据库：

- PostgreSQL
- SQL Server
- MySQL
- Oracle

当前 SQL 生成规则：

- SQL Server 临时表名会自动规范为 `#TempTable`
- Oracle 使用 `CREATE GLOBAL TEMPORARY TABLE`
- Oracle 日期值输出为 `TO_DATE(...)`
- SQL Server 单批 `INSERT` 最多 1000 行
- Oracle 批量插入会生成 `INSERT ALL`
- 勾选“限制文本字段长度”后，所有列统一生成文本类型
  - PostgreSQL / MySQL：`VARCHAR(1000)`
  - SQL Server：`NVARCHAR(1000)`
  - Oracle：`VARCHAR2(1000)`

### CSV 预览工具

- 只读方式预览 CSV，避免 Excel 自动改格式
- 支持搜索、排序和大文件预览
- 使用共享读取方式打开文件，不强锁文件

### JSON 处理工具

- JSON 美化、压缩、校验
- 左侧原始 JSON 编辑，右侧 GRID 可视化浏览
- 支持展开 / 折叠嵌套对象和数组
- 支持导出 JSON / CSV

搜索能力：

- 左右两侧都支持搜索
- 左右两侧都支持 `Aa` 大小写敏感开关
- 左右两侧搜索已统一使用同一套节点命中规则
- 搜索范围覆盖：
  - Key
  - Value
  - 表格列名
  - 折叠节点标题
  - 嵌套摘要
- 右侧搜索会按需展开命中的折叠节点
- 点击右侧 GRID 的 key / value 可联动定位到左侧原始 JSON

### JSON 对比工具

- 深度对比两个 JSON
- 展示新增、删除、值变化、类型变化
- 支持导入、美化和导出对比结果

### Excalidraw 画板

- 基于 WebView2 嵌入 Excalidraw
- 支持打开 `.excalidraw` / `.json`
- 支持保存、导出 PNG、导出 SVG
- 使用 Base64 传输数据，避免特殊字符转义问题

说明：

- 该页面依赖 `Microsoft Edge WebView2 Runtime`
- 单文件 EXE 可以成立，但 Excalidraw 功能仍要求目标机器具备 WebView2 Runtime

### 发票打印工具

- 支持 PDF / OFD / JPG / PNG / BMP / TIFF
- 支持拖拽导入和目录递归导入
- 支持旋转、分页预览、批量打印
- 支持模板保存、打印历史记录

### 系统设置与小工具

设置页支持保存以下默认配置：

- 默认数据库类型
- 默认表名
- 默认批量行数
- 是否默认生成 DROP TABLE
- 是否默认启用批量 INSERT
- 是否默认限制字段长度
- 默认导出路径

内置小工具包括：

- 身份信息生成
- Base64 编解码
- UUID / GUID 生成
- 文本哈希计算
- URL 编解码
- 正则测试
- 文本对比

## 技术栈

- .NET 10
- WPF
- EPPlus
- ExcelDataReader
- CsvHelper
- Microsoft.Data.SqlClient
- Npgsql
- Microsoft.Web.WebView2

## 目录结构

```text
.
├─ WpfApp1.slnx
├─ README.md
├─ artifacts/
│  └─ publish/
├─ publish/
└─ WpfApp1/
   ├─ WpfApp1.csproj
   ├─ Services/
   ├─ Views/
   ├─ Models/
   ├─ App.xaml
   └─ MainWindow.xaml
```

## 开发环境

- Windows 10/11 x64
- .NET 10 SDK

## 运行项目

```powershell
dotnet run --project .\WpfApp1\WpfApp1.csproj
```

## 打包单文件 EXE

项目已经内置发布配置 `SingleFile-win-x64`。

执行命令：

```powershell
dotnet publish .\WpfApp1\WpfApp1.csproj -c Release -p:PublishProfile=SingleFile-win-x64
```

输出目录：

```text
artifacts/publish/win-x64-single/CCToolbox.exe
```

当前打包策略：

- `win-x64`
- `Release`
- `Self-contained`
- `PublishSingleFile=true`
- `IncludeNativeLibrariesForSelfExtract=true`
- `EnableCompressionInSingleFile=true`
- 不启用 `PublishTrimmed`

说明：

- 输出目录已切换到 `artifacts/publish/win-x64-single/`
- 该目录已加入忽略规则，不会污染源码目录
- 发布目录默认只保留一个主交付文件 `CCToolbox.exe`

## 运行时文件

程序运行后，可能会在 EXE 同目录生成这些文件：

- `settings.json`
- `invoice_templates.json`
- `invoice_print_history.json`

这是当前设计的一部分，便于绿色便携分发，不视为异常。

Excalidraw 的 WebView2 用户数据目录位于当前用户本地应用数据目录下，不写入仓库。

## 已知前置条件

- 主程序单文件 EXE 不依赖额外安装 `.NET Desktop Runtime`
- Excalidraw 页面依赖 `Microsoft Edge WebView2 Runtime`

如果目标机器未安装 WebView2 Runtime，画板页面会初始化失败，并提示下载安装。

## 最近更新

### 文档同步

- 更新了单文件打包命令和输出目录说明
- 补充了运行时生成的配置文件说明
- 修正文档中的数据库支持范围为 PostgreSQL / SQL Server / MySQL / Oracle

### 数据导入

- SQL Server 临时表名自动补 `#`
- Oracle 生成全局临时表语法
- 文本列按统一文本类型输出

### JSON 工具

- 左右搜索支持统一的大小写规则
- 修复右侧节点搜索遗漏
- 修复右侧点击定位左侧 JSON 时的高亮偏移问题

## 许可

MIT
